"""
PDF → DOCX converter that places every element at its **exact PDF coordinate**.

Instead of relying on pdf2docx's layout reconstruction (which causes overlap
and line-shift artefacts), this module:

1. Uses PyMuPDF to extract every text span, image, and drawing with its
   exact bounding box, font, size, colour, and flags.
2. Creates a DOCX whose page size matches the PDF page exactly.
3. Places each text span as an **invisible floating text box** (DrawingML
   ``wp:anchor``) at the precise (x, y) position.  The text inside each
   box is fully editable in Word.
4. Extracts and inserts images as floating pictures at their exact position.
5. Draws table borders / rectangles as shapes.
6. Renders complex graphics (path-based text, filled shapes) as rasterized images.

The result is an editable DOCX that visually matches the PDF character-by-
character.
"""

from __future__ import annotations

import html
import io
import os
import sys
import tempfile
from pathlib import Path
from typing import Optional, Sequence, List, Tuple, Dict, Any

import fitz  # PyMuPDF
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.shared import Emu, Pt
from PIL import Image


# ── Unit helpers ─────────────────────────────────────────────────────────────

_PT_TO_EMU = 12700          # 1 pt  = 12 700 EMU
_IN_TO_EMU = 914400         # 1 in  = 914 400 EMU
_PDF_PT_TO_EMU = _PT_TO_EMU  # PDF points → EMU (1:1 since both use 72 dpi)

def _pt2emu(pt: float) -> int:
    return int(pt * _PT_TO_EMU)


# ── Character width estimation for better text box sizing ─────────────────────

# Approximate character width multipliers for common font categories
_CHAR_WIDTH_FACTORS = {
    # Narrow characters
    'i': 0.35, 'l': 0.35, 'I': 0.35, '1': 0.55, '.': 0.35, ',': 0.35, 
    ':': 0.35, ';': 0.35, '!': 0.35, "'": 0.25, '"': 0.45, '|': 0.35,
    'j': 0.40, 'f': 0.40, 't': 0.45, 'r': 0.45,
    # Wide characters  
    'w': 0.85, 'm': 0.90, 'W': 1.00, 'M': 0.95, '@': 0.95, '%': 0.90,
    # Default for normal width characters
}
_DEFAULT_CHAR_WIDTH = 0.60


def _estimate_text_width(text: str, font_size: float, font_name: str = "") -> float:
    """Estimate text width in points based on character composition."""
    if not text:
        return 0
    
    # Check if it's a monospace font
    mono_indicators = ['mono', 'courier', 'consolas', 'menlo', 'fixed']
    is_mono = any(ind in font_name.lower() for ind in mono_indicators)
    
    if is_mono:
        # Monospace: all characters have same width (approx 0.6 * font_size)
        return len(text) * font_size * 0.62
    
    # Variable width: estimate per character
    total_width = 0.0
    for char in text:
        factor = _CHAR_WIDTH_FACTORS.get(char, _DEFAULT_CHAR_WIDTH)
        total_width += font_size * factor
    
    # Add a small buffer to prevent clipping
    return total_width * 1.05


# ── Floating text box builder ────────────────────────────────────────────────

_SHAPE_ID_COUNTER = 0

def _next_shape_id() -> int:
    global _SHAPE_ID_COUNTER
    _SHAPE_ID_COUNTER += 1
    return _SHAPE_ID_COUNTER


def _escape(text: str) -> str:
    """Escape text for XML embedding."""
    return html.escape(text, quote=True)


def _make_run_xml(
    text: str,
    font_name: str = "Arial",
    font_size_half_pt: int = 24,
    bold: bool = False,
    italic: bool = False,
    color_hex: str = "000000",
    superscript: bool = False,
    subscript: bool = False,
    char_spacing_twips: int = 0,
) -> str:
    """Build a ``<w:r>`` XML fragment for one styled text span.
    
    Parameters
    ----------
    char_spacing_twips:
        Character spacing in twips (1/20 of a point). Positive = expand, negative = condense.
    """
    flags = ""
    if bold:
        flags += "<w:b/>"
    if italic:
        flags += "<w:i/>"
    if superscript:
        flags += '<w:vertAlign w:val="superscript"/>'
    elif subscript:
        flags += '<w:vertAlign w:val="subscript"/>'
    if char_spacing_twips != 0:
        flags += f'<w:spacing w:val="{char_spacing_twips}"/>'
    
    return (
        "<w:r><w:rPr>"
        '<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}" w:eastAsia="{font}"/>'
        "{flags}"
        '<w:color w:val="{color}"/>'
        '<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
        "</w:rPr>"
        '<w:t xml:space="preserve">{text}</w:t>'
        "</w:r>"
    ).format(
        font=_escape(font_name),
        flags=flags,
        color=color_hex,
        sz=font_size_half_pt,
        text=_escape(text),
    )


def _add_textbox(
    paragraph,
    runs_xml: str,
    x_emu: int,
    y_emu: int,
    w_emu: int,
    h_emu: int,
    shape_id: int,
) -> None:
    """Append an invisible floating text box to *paragraph* at (x, y)."""
    xml = (
        '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        '     xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        '     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        '     xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
        '     xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        "<mc:AlternateContent>"
        '<mc:Choice Requires="wps"><w:drawing>'
        '<wp:anchor distT="0" distB="0" distL="0" distR="0"'
        ' simplePos="0" relativeHeight="{z}"'
        ' behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="page">'
        "  <wp:posOffset>{x}</wp:posOffset>"
        "</wp:positionH>"
        '<wp:positionV relativeFrom="page">'
        "  <wp:posOffset>{y}</wp:posOffset>"
        "</wp:positionV>"
        '<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        "<wp:wrapNone/>"
        '<wp:docPr id="{sid}" name="TB{sid}"/>'
        "<wp:cNvGraphicFramePr/>"
        "<a:graphic>"
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        "<wps:wsp>"
        '<wps:cNvSpPr txBox="1"/>'
        "<wps:spPr>"
        '  <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        "  <a:noFill/><a:ln><a:noFill/></a:ln>"
        "</wps:spPr>"
        "<wps:txbx><w:txbxContent>"
        '<w:p><w:pPr><w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/>'
        "</w:pPr>"
        "{runs}"
        "</w:p>"
        "</w:txbxContent></wps:txbx>"
        '<wps:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"'
        ' anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
        "</wps:wsp>"
        "</a:graphicData></a:graphic>"
        "</wp:anchor>"
        "</w:drawing></mc:Choice>"
        "<mc:Fallback><w:pict/></mc:Fallback>"
        "</mc:AlternateContent>"
        "</w:r>"
    ).format(
        x=x_emu,
        y=y_emu,
        cx=w_emu,
        cy=h_emu,
        sid=shape_id,
        z=251659264 + shape_id,
        runs=runs_xml,
    )

    run_element = parse_xml(xml)
    paragraph._element.append(run_element)


# ── Complex graphics detection and rasterization ─────────────────────────────

def _has_complex_graphics(page: fitz.Page) -> List[fitz.Rect]:
    """
    Detect regions with complex graphics (filled paths, gradients, etc.)
    that need to be rasterized for proper rendering.
    
    Returns a list of rectangles containing complex graphics.
    """
    complex_regions = []
    
    drawings = page.get_drawings()
    
    for drawing in drawings:
        draw_fill = drawing.get("fill")
        draw_color = drawing.get("color")
        items = drawing.get("items", [])
        
        # Check for complex path elements (curves, filled shapes)
        has_curves = False
        has_fill = draw_fill is not None
        has_complex_path = False
        
        for item in items:
            kind = item[0]
            if kind in ("c", "qu"):  # curves, quadratic bezier
                has_curves = True
            if kind == "re" and has_fill:
                has_complex_path = True
        
        # If this drawing has fills (not just strokes), mark its region
        if has_fill or has_curves:
            rect = drawing.get("rect")
            if rect and rect.width > 5 and rect.height > 5:
                complex_regions.append(fitz.Rect(rect))
    
    return complex_regions


def _merge_overlapping_rects(rects: List[fitz.Rect], margin: float = 5) -> List[fitz.Rect]:
    """Merge overlapping or nearby rectangles into larger regions."""
    if not rects:
        return []
    
    merged = []
    used = [False] * len(rects)
    
    for i, rect in enumerate(rects):
        if used[i]:
            continue
        
        current = fitz.Rect(rect)
        current.x0 -= margin
        current.y0 -= margin
        current.x1 += margin
        current.y1 += margin
        
        changed = True
        while changed:
            changed = False
            for j, other in enumerate(rects):
                if used[j] or i == j:
                    continue
                other_expanded = fitz.Rect(other)
                other_expanded.x0 -= margin
                other_expanded.y0 -= margin
                other_expanded.x1 += margin
                other_expanded.y1 += margin
                
                if current.intersects(other_expanded):
                    current = current | other_expanded
                    used[j] = True
                    changed = True
        
        used[i] = True
        merged.append(current)
    
    return merged


def _render_region_as_image(
    page: fitz.Page,
    rect: fitz.Rect,
    dpi: int = 300
) -> Tuple[bytes, int, int]:
    """
    Render a specific region of the page as a PNG image.
    
    Returns:
        Tuple of (image_bytes, width_emu, height_emu)
    """
    # Calculate the transformation matrix for the clip
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    
    # Clip and render just this region
    clip = fitz.Rect(rect)
    pix = page.get_pixmap(matrix=mat, clip=clip, alpha=False)
    
    # Convert to PNG bytes
    img_bytes = pix.tobytes("png")
    
    # Return dimensions in EMU
    w_emu = _pt2emu(rect.width)
    h_emu = _pt2emu(rect.height)
    
    return img_bytes, w_emu, h_emu


# ── Floating image builder ───────────────────────────────────────────────────

def _add_floating_image(
    doc: Document,
    paragraph,
    image_bytes: bytes,
    x_emu: int,
    y_emu: int,
    w_emu: int,
    h_emu: int,
    shape_id: int,
) -> None:
    """Insert a floating image at an exact page position."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    # Add image to the document's media and get the relationship ID
    image_part, rId = doc.part.get_or_add_image_part(io.BytesIO(image_bytes))

    xml = (
        '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        '     xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        '     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        '     xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
        '     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        '     xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        "<mc:AlternateContent>"
        '<mc:Choice Requires="wps"><w:drawing>'
        '<wp:anchor distT="0" distB="0" distL="0" distR="0"'
        ' simplePos="0" relativeHeight="{z}"'
        ' behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="page">'
        "  <wp:posOffset>{x}</wp:posOffset>"
        "</wp:positionH>"
        '<wp:positionV relativeFrom="page">'
        "  <wp:posOffset>{y}</wp:posOffset>"
        "</wp:positionV>"
        '<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        "<wp:wrapNone/>"
        '<wp:docPr id="{sid}" name="Img{sid}"/>'
        "<wp:cNvGraphicFramePr>"
        '  <a:graphicFrameLocks noChangeAspect="1"/>'
        "</wp:cNvGraphicFramePr>"
        "<a:graphic>"
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        "<pic:pic>"
        "<pic:nvPicPr>"
        '  <pic:cNvPr id="{sid}" name="Img{sid}"/>'
        "  <pic:cNvPicPr/>"
        "</pic:nvPicPr>"
        "<pic:blipFill>"
        '  <a:blip r:embed="{rId}"/>'
        "  <a:stretch><a:fillRect/></a:stretch>"
        "</pic:blipFill>"
        "<pic:spPr>"
        '  <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        "</pic:spPr>"
        "</pic:pic>"
        "</a:graphicData></a:graphic>"
        "</wp:anchor>"
        "</w:drawing></mc:Choice>"
        "<mc:Fallback><w:pict/></mc:Fallback>"
        "</mc:AlternateContent>"
        "</w:r>"
    ).format(
        x=x_emu,
        y=y_emu,
        cx=w_emu,
        cy=h_emu,
        sid=shape_id,
        z=251650000 + shape_id,
        rId=rId,
    )

    run_element = parse_xml(xml)
    paragraph._element.append(run_element)


# ── Page conversion ──────────────────────────────────────────────────────────

def _color_to_hex(color) -> str:
    """Convert a PyMuPDF colour (int, tuple, list, or None) to 6-char hex."""
    if color is None:
        return "000000"
    if isinstance(color, (tuple, list)):
        if len(color) >= 3:
            r, g, b = int(color[0] * 255), int(color[1] * 255), int(color[2] * 255)
            return f"{r:02X}{g:02X}{b:02X}"
        if len(color) == 1:
            v = int(color[0] * 255)
            return f"{v:02X}{v:02X}{v:02X}"
        return "000000"
    if isinstance(color, float):
        v = int(color * 255)
        return f"{v:02X}{v:02X}{v:02X}"
    # int (sRGB packed)
    r = (color >> 16) & 0xFF
    g = (color >> 8) & 0xFF
    b = color & 0xFF
    return f"{r:02X}{g:02X}{b:02X}"


def _process_page(
    pdf_doc: fitz.Document,
    word_doc: Document,
    page: fitz.Page,
    is_first: bool,
    verbose: bool,
    dpi: int = 300,
) -> None:
    """Convert one PDF page into a DOCX section with positioned elements."""
    rect = page.rect
    w_emu = _pt2emu(rect.width)
    h_emu = _pt2emu(rect.height)

    # ── Set up DOCX section matching PDF page size ───────────────────────
    if is_first:
        section = word_doc.sections[0]
    else:
        section = word_doc.add_section()

    landscape = rect.width > rect.height
    if landscape:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Emu(max(w_emu, h_emu))
        section.page_height = Emu(min(w_emu, h_emu))
    else:
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width = Emu(w_emu)
        section.page_height = Emu(h_emu)

    section.top_margin = Emu(0)
    section.bottom_margin = Emu(0)
    section.left_margin = Emu(0)
    section.right_margin = Emu(0)
    section.header_distance = Emu(0)
    section.footer_distance = Emu(0)

    # One anchor paragraph per page — all floating boxes attach here
    anchor_para = word_doc.add_paragraph()
    anchor_para.paragraph_format.space_before = Pt(0)
    anchor_para.paragraph_format.space_after = Pt(0)

    # ── Extract and place images FIRST (behind text) ─────────────────────
    # Use multiple extraction methods to capture all image types
    
    # Method 1: Standard image extraction with full=True
    image_list = page.get_images(full=True)
    processed_rects = []  # Track processed image regions to avoid duplicates
    
    for img_info in image_list:
        xref = img_info[0]
        
        try:
            # Get ALL rectangles where this image appears (not just the first)
            img_rects = page.get_image_rects(xref)
            if not img_rects:
                continue
            
            base_image = pdf_doc.extract_image(xref)
            if not base_image or not base_image.get("image"):
                continue
            
            img_bytes = base_image["image"]
            
            # Process each instance of the image
            for img_rect in img_rects:
                # Skip if this rect overlaps significantly with already processed rect
                rect_tuple = (round(img_rect.x0, 1), round(img_rect.y0, 1), 
                             round(img_rect.x1, 1), round(img_rect.y1, 1))
                if rect_tuple in processed_rects:
                    continue
                processed_rects.append(rect_tuple)
                
                ix_emu = _pt2emu(img_rect.x0)
                iy_emu = _pt2emu(img_rect.y0)
                iw_emu = _pt2emu(img_rect.width)
                ih_emu = _pt2emu(img_rect.height)

                if iw_emu <= 0 or ih_emu <= 0:
                    continue

                _add_floating_image(
                    word_doc,
                    anchor_para,
                    img_bytes,
                    ix_emu,
                    iy_emu,
                    iw_emu,
                    ih_emu,
                    _next_shape_id(),
                )
        except Exception:
            # Skip problematic images rather than failing the whole page
            continue
    
    # Method 2: Extract images from XObject Forms (nested images)
    try:
        xref_list = page.get_xobjects()
        for xref_info in xref_list:
            try:
                xref = xref_info[0]
                if xref in [img[0] for img in image_list]:
                    continue  # Already processed
                
                # Try to extract as image
                base_image = pdf_doc.extract_image(xref)
                if base_image and base_image.get("image"):
                    # Try to get position info
                    img_rects = page.get_image_rects(xref)
                    for img_rect in (img_rects or []):
                        rect_tuple = (round(img_rect.x0, 1), round(img_rect.y0, 1), 
                                     round(img_rect.x1, 1), round(img_rect.y1, 1))
                        if rect_tuple in processed_rects:
                            continue
                        processed_rects.append(rect_tuple)
                        
                        img_bytes = base_image["image"]
                        ix_emu = _pt2emu(img_rect.x0)
                        iy_emu = _pt2emu(img_rect.y0)
                        iw_emu = _pt2emu(img_rect.width)
                        ih_emu = _pt2emu(img_rect.height)
                        
                        if iw_emu > 0 and ih_emu > 0:
                            _add_floating_image(
                                word_doc,
                                anchor_para,
                                img_bytes,
                                ix_emu,
                                iy_emu,
                                iw_emu,
                                ih_emu,
                                _next_shape_id(),
                            )
            except Exception:
                continue
    except Exception:
        pass

    # ── Extract text with exact positions ────────────────────────────────
    # Use rawdict for more precise character-level positioning
    blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE | fitz.TEXT_PRESERVE_LIGATURES)["blocks"]

    for block in blocks:
        if block["type"] != 0:  # 0 = text block
            continue

        for line in block["lines"]:
            # Process each span individually with proper width calculation
            for span in line["spans"]:
                text = span["text"]
                if not text or text.isspace():
                    continue

                bbox = span["bbox"]  # (x0, y0, x1, y1) in PDF points
                font = span["font"]
                size = span["size"]
                flags = span["flags"]
                color = span["color"]

                # Convert font flags
                is_bold = bool(flags & 2 ** 4)  # bit 4 = bold
                is_italic = bool(flags & 2 ** 1)  # bit 1 = italic
                is_superscript = bool(flags & 2 ** 0)  # bit 0 = superscript

                # Font size in half-points (OOXML uses half-points)
                size_half_pt = max(int(round(size * 2)), 2)

                color_hex = _color_to_hex(color)

                # Clean font name (remove subset prefix like "ABCDEF+")
                clean_font = font
                if "+" in clean_font:
                    clean_font = clean_font.split("+", 1)[1]

                # Position and size in EMU
                x_emu = _pt2emu(bbox[0])
                y_emu = _pt2emu(bbox[1])
                
                # Calculate width from PDF bbox first
                pdf_width = bbox[2] - bbox[0]
                pdf_height = bbox[3] - bbox[1]
                
                # Use improved character-based width estimation as a minimum
                estimated_width = _estimate_text_width(text, size, clean_font)
                
                # Use the larger of PDF-reported width or estimated width
                # Add 15% padding to prevent overlap with adjacent text
                final_width = max(pdf_width, estimated_width) * 1.15
                
                # Height should accommodate descenders and ascenders
                final_height = max(pdf_height, size * 1.5)
                
                box_w = _pt2emu(final_width)
                box_h = _pt2emu(final_height)

                run_xml = _make_run_xml(
                    text,
                    font_name=clean_font,
                    font_size_half_pt=size_half_pt,
                    bold=is_bold,
                    italic=is_italic,
                    color_hex=color_hex,
                    superscript=is_superscript,
                )

                _add_textbox(
                    anchor_para,
                    run_xml,
                    x_emu,
                    y_emu,
                    box_w,
                    box_h,
                    _next_shape_id(),
                )

    # ── Draw rectangles / lines for table borders ────────────────────────
    # First, identify complex graphics regions that need rasterization
    complex_regions = _has_complex_graphics(page)
    
    # Track regions we've rendered as images
    rasterized_regions = []
    
    drawings = page.get_drawings()
    
    # Collect all drawings with complex paths (curves, fills) for potential rasterization
    complex_drawings = []
    simple_drawings = []
    
    for drawing in drawings:
        draw_fill = drawing.get("fill")
        items = drawing.get("items", [])
        
        # Check if drawing has complex elements
        has_curves = any(item[0] in ("c", "qu", "curve") for item in items)
        has_fill = draw_fill is not None
        
        if has_curves or (has_fill and len(items) > 2):
            complex_drawings.append(drawing)
        else:
            simple_drawings.append(drawing)
    
    # Render complex graphics regions as images
    if complex_regions:
        merged_regions = _merge_overlapping_rects(complex_regions)
        for region in merged_regions:
            # Skip very small regions
            if region.width < 10 or region.height < 10:
                continue
            
            try:
                img_bytes, w_emu, h_emu = _render_region_as_image(page, region, dpi=min(dpi, 200))
                
                _add_floating_image(
                    word_doc,
                    anchor_para,
                    img_bytes,
                    _pt2emu(region.x0),
                    _pt2emu(region.y0),
                    w_emu,
                    h_emu,
                    _next_shape_id(),
                )
                rasterized_regions.append(region)
            except Exception:
                pass
    
    # Process simple drawings as vector shapes
    for drawing in simple_drawings:
        draw_width = drawing.get("width")
        if draw_width is None:
            draw_width = 0.5
        draw_color = drawing.get("color")
        draw_fill = drawing.get("fill")
        draw_rect_region = drawing.get("rect")
        
        # Skip if this drawing is in a rasterized region
        if draw_rect_region:
            drawing_rect = fitz.Rect(draw_rect_region)
            skip = False
            for rast_region in rasterized_regions:
                if rast_region.contains(drawing_rect):
                    skip = True
                    break
            if skip:
                continue

        for item in drawing.get("items", []):
            kind = item[0]
            if kind == "re":  # rectangle
                draw_rect = item[1]
                _draw_rect_shape(
                    anchor_para,
                    _pt2emu(draw_rect.x0),
                    _pt2emu(draw_rect.y0),
                    _pt2emu(draw_rect.width),
                    _pt2emu(draw_rect.height),
                    color_hex=_color_to_hex(draw_color),
                    fill_hex=_color_to_hex(draw_fill) if draw_fill is not None else None,
                    stroke_width_emu=max(_pt2emu(draw_width), 6350),
                    shape_id=_next_shape_id(),
                )
            elif kind == "l":  # line
                p1, p2 = item[1], item[2]
                x0, y0 = p1.x, p1.y
                x1, y1 = p2.x, p2.y
                _draw_line_shape(
                    anchor_para,
                    _pt2emu(x0),
                    _pt2emu(y0),
                    _pt2emu(x1),
                    _pt2emu(y1),
                    color_hex=_color_to_hex(draw_color),
                    stroke_width_emu=max(_pt2emu(draw_width), 6350),
                    shape_id=_next_shape_id(),
                )
    
    # For complex drawings that weren't in merged regions, render them individually
    for drawing in complex_drawings:
        draw_rect_region = drawing.get("rect")
        if not draw_rect_region:
            continue
        
        drawing_rect = fitz.Rect(draw_rect_region)
        
        # Skip if already rasterized
        skip = False
        for rast_region in rasterized_regions:
            if rast_region.contains(drawing_rect) or rast_region.intersects(drawing_rect):
                skip = True
                break
        if skip:
            continue
        
        # Skip very small regions
        if drawing_rect.width < 5 or drawing_rect.height < 5:
            continue
        
        try:
            img_bytes, w_emu, h_emu = _render_region_as_image(page, drawing_rect, dpi=min(dpi, 200))
            
            _add_floating_image(
                word_doc,
                anchor_para,
                img_bytes,
                _pt2emu(drawing_rect.x0),
                _pt2emu(drawing_rect.y0),
                w_emu,
                h_emu,
                _next_shape_id(),
            )
            rasterized_regions.append(drawing_rect)
        except Exception:
            pass


# ── Shape helpers (rectangles, lines) ────────────────────────────────────────

def _draw_rect_shape(
    paragraph,
    x_emu: int,
    y_emu: int,
    w_emu: int,
    h_emu: int,
    color_hex: str = "000000",
    fill_hex: str | None = None,
    stroke_width_emu: int = 12700,
    shape_id: int = 1,
) -> None:
    """Draw a rectangle outline (and optional fill) at exact position."""
    if w_emu <= 0 or h_emu <= 0:
        return

    fill_xml = f'<a:solidFill><a:srgbClr val="{fill_hex}"/></a:solidFill>' if fill_hex else "<a:noFill/>"
    xml = (
        '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        '     xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        '     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        '     xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
        '     xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        "<mc:AlternateContent>"
        '<mc:Choice Requires="wps"><w:drawing>'
        '<wp:anchor distT="0" distB="0" distL="0" distR="0"'
        ' simplePos="0" relativeHeight="{z}"'
        ' behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="page"><wp:posOffset>{x}</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="page"><wp:posOffset>{y}</wp:posOffset></wp:positionV>'
        '<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        "<wp:wrapNone/>"
        '<wp:docPr id="{sid}" name="R{sid}"/>'
        "<wp:cNvGraphicFramePr/>"
        "<a:graphic>"
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        "<wps:wsp>"
        "<wps:cNvSpPr/>"
        "<wps:spPr>"
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        "{fill}"
        "<a:ln w=\"{lw}\">"
        '<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        "</a:ln>"
        "</wps:spPr>"
        '<wps:bodyPr/>'
        "</wps:wsp>"
        "</a:graphicData></a:graphic>"
        "</wp:anchor>"
        "</w:drawing></mc:Choice>"
        "<mc:Fallback><w:pict/></mc:Fallback>"
        "</mc:AlternateContent>"
        "</w:r>"
    ).format(
        x=x_emu, y=y_emu, cx=w_emu, cy=h_emu,
        sid=shape_id, z=251640000 + shape_id,
        color=color_hex, fill=fill_xml, lw=stroke_width_emu,
    )
    paragraph._element.append(parse_xml(xml))


def _draw_line_shape(
    paragraph,
    x0_emu: int,
    y0_emu: int,
    x1_emu: int,
    y1_emu: int,
    color_hex: str = "000000",
    stroke_width_emu: int = 12700,
    shape_id: int = 1,
) -> None:
    """Draw a line from (x0,y0) to (x1,y1) at exact position."""
    # Compute bounding box
    bx = min(x0_emu, x1_emu)
    by = min(y0_emu, y1_emu)
    bw = abs(x1_emu - x0_emu) or stroke_width_emu
    bh = abs(y1_emu - y0_emu) or stroke_width_emu

    # Determine flip
    flipH = "1" if x1_emu < x0_emu else "0"
    flipV = "1" if y1_emu < y0_emu else "0"

    xml = (
        '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        '     xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        '     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        '     xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
        '     xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        "<mc:AlternateContent>"
        '<mc:Choice Requires="wps"><w:drawing>'
        '<wp:anchor distT="0" distB="0" distL="0" distR="0"'
        ' simplePos="0" relativeHeight="{z}"'
        ' behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="page"><wp:posOffset>{bx}</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="page"><wp:posOffset>{by}</wp:posOffset></wp:positionV>'
        '<wp:extent cx="{bw}" cy="{bh}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        "<wp:wrapNone/>"
        '<wp:docPr id="{sid}" name="L{sid}"/>'
        "<wp:cNvGraphicFramePr/>"
        "<a:graphic>"
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        "<wps:wsp>"
        '<wps:cNvCnPr/>'
        "<wps:spPr>"
        '<a:xfrm flipH="{fH}" flipV="{fV}">'
        '<a:off x="0" y="0"/><a:ext cx="{bw}" cy="{bh}"/></a:xfrm>'
        '<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
        '<a:ln w="{lw}">'
        '<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        "</a:ln>"
        "</wps:spPr>"
        '<wps:bodyPr/>'
        "</wps:wsp>"
        "</a:graphicData></a:graphic>"
        "</wp:anchor>"
        "</w:drawing></mc:Choice>"
        "<mc:Fallback><w:pict/></mc:Fallback>"
        "</mc:AlternateContent>"
        "</w:r>"
    ).format(
        bx=bx, by=by, bw=bw, bh=bh,
        sid=shape_id, z=251630000 + shape_id,
        color=color_hex, lw=stroke_width_emu,
        fH=flipH, fV=flipV,
    )
    paragraph._element.append(parse_xml(xml))


# ── Public API ───────────────────────────────────────────────────────────────

def convert_pdf_to_docx(
    pdf_path: str | Path,
    docx_path: Optional[str | Path] = None,
    *,
    pages: Optional[Sequence[int]] = None,
    dpi: int = 300,
    verbose: bool = False,
) -> Path:
    """Convert a PDF to an editable DOCX with exact layout preservation.

    Every text span, image, and line/rectangle is placed at its exact PDF
    coordinate using floating elements.  The output is fully editable in
    Word.

    Parameters
    ----------
    pdf_path:
        Path to the source PDF file.
    docx_path:
        Destination path.  Defaults to ``<input>.docx``.
    pages:
        0-based page indices.  ``None`` → all pages.
    dpi:
        Resolution hint (used for image extraction quality).
    verbose:
        Print progress to stderr.
    """
    global _SHAPE_ID_COUNTER
    _SHAPE_ID_COUNTER = 0

    pdf_path = Path(pdf_path).resolve()
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    if docx_path is None:
        docx_path = pdf_path.with_suffix(".docx")
    else:
        docx_path = Path(docx_path).resolve()

    docx_path.parent.mkdir(parents=True, exist_ok=True)

    pdf_doc = fitz.open(str(pdf_path))
    word_doc = Document()

    page_indices = list(pages) if pages is not None else list(range(len(pdf_doc)))
    total = len(page_indices)

    for i, idx in enumerate(page_indices):
        page = pdf_doc[idx]
        if verbose:
            print(
                f"  [{i + 1}/{total}] Processing page {idx + 1} …",
                file=sys.stderr,
            )
        _process_page(pdf_doc, word_doc, page, is_first=(i == 0), verbose=verbose, dpi=dpi)

    word_doc.save(str(docx_path))
    pdf_doc.close()

    if verbose:
        print("Done.", file=sys.stderr)

    return docx_path

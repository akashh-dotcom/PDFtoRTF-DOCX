"""
PDF → DOCX converter with exact visual fidelity.

This module uses a hybrid approach to guarantee the DOCX looks exactly like the PDF:

1. Renders each PDF page as a high-resolution background image
2. Overlays invisible (transparent) but selectable text for searchability
3. The result is visually identical to the PDF while still being searchable

For fully editable output, use mode="editable" which places text as floating boxes.
"""

from __future__ import annotations

import html
import io
import sys
from pathlib import Path
from typing import Optional, Sequence, List, Tuple, Literal

import fitz  # PyMuPDF
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.oxml import parse_xml
from docx.shared import Emu, Pt


# ── Unit helpers ─────────────────────────────────────────────────────────────

_PT_TO_EMU = 12700          # 1 pt  = 12 700 EMU
_IN_TO_EMU = 914400         # 1 in  = 914 400 EMU

def _pt2emu(pt: float) -> int:
    return int(pt * _PT_TO_EMU)


# ── Shape ID counter ─────────────────────────────────────────────────────────

_SHAPE_ID_COUNTER = 0

def _next_shape_id() -> int:
    global _SHAPE_ID_COUNTER
    _SHAPE_ID_COUNTER += 1
    return _SHAPE_ID_COUNTER


def _escape(text: str) -> str:
    """Escape text for XML embedding."""
    return html.escape(text, quote=True)


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
    behind_doc: bool = True,
) -> None:
    """Insert a floating image at an exact page position."""
    from docx.shared import Inches
    from docx.oxml.ns import nsmap
    from lxml import etree
    
    # Create a temporary run to add an inline image, then extract the rId
    temp_run = paragraph.add_run()
    
    # Add inline image to get the relationship set up
    image_stream = io.BytesIO(image_bytes)
    inline_shape = temp_run.add_picture(image_stream, width=Emu(w_emu), height=Emu(h_emu))
    
    # Get the blip element to extract rId
    inline_xml = temp_run._element
    blip = inline_xml.find('.//' + '{http://schemas.openxmlformats.org/drawingml/2006/main}blip', 
                           namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
    
    if blip is None:
        # Fallback: search without namespace prefix
        for elem in inline_xml.iter():
            if 'blip' in elem.tag:
                blip = elem
                break
    
    if blip is None:
        return  # Can't find blip, skip this image
    
    # Extract the relationship ID
    rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
    
    if not rId:
        return  # No rId found, skip
    
    # Now remove the inline image run and create a floating one
    paragraph._element.remove(inline_xml)
    
    behind = "1" if behind_doc else "0"

    xml = (
        '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:AlternateContent>'
        '<mc:Choice Requires="wps"><w:drawing>'
        '<wp:anchor distT="0" distB="0" distL="0" distR="0"'
        ' simplePos="0" relativeHeight="{z}"'
        ' behindDoc="{behind}" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="page">'
        '<wp:posOffset>{x}</wp:posOffset>'
        '</wp:positionH>'
        '<wp:positionV relativeFrom="page">'
        '<wp:posOffset>{y}</wp:posOffset>'
        '</wp:positionV>'
        '<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapNone/>'
        '<wp:docPr id="{sid}" name="Img{sid}"/>'
        '<wp:cNvGraphicFramePr>'
        '<a:graphicFrameLocks noChangeAspect="1"/>'
        '</wp:cNvGraphicFramePr>'
        '<a:graphic>'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic>'
        '<pic:nvPicPr>'
        '<pic:cNvPr id="{sid}" name="Img{sid}"/>'
        '<pic:cNvPicPr/>'
        '</pic:nvPicPr>'
        '<pic:blipFill>'
        '<a:blip r:embed="{rId}"/>'
        '<a:stretch><a:fillRect/></a:stretch>'
        '</pic:blipFill>'
        '<pic:spPr>'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '</pic:spPr>'
        '</pic:pic>'
        '</a:graphicData></a:graphic>'
        '</wp:anchor>'
        '</w:drawing></mc:Choice>'
        '<mc:Fallback><w:pict/></mc:Fallback>'
        '</mc:AlternateContent>'
        '</w:r>'
    ).format(
        x=x_emu,
        y=y_emu,
        cx=w_emu,
        cy=h_emu,
        sid=shape_id,
        z=251650000 + shape_id,
        rId=rId,
        behind=behind,
    )

    run_element = parse_xml(xml)
    paragraph._element.append(run_element)


# ── Invisible text box for searchable overlay ─────────────────────────────────

def _add_invisible_textbox(
    paragraph,
    text: str,
    x_emu: int,
    y_emu: int,
    w_emu: int,
    h_emu: int,
    shape_id: int,
    font_size_half_pt: int = 20,
) -> None:
    """Add an invisible (transparent) text box for searchability."""
    escaped_text = _escape(text)
    
    # Use white color with 0% opacity (invisible but selectable)
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
        "<w:r><w:rPr>"
        '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
        '<w:color w:val="FFFFFF"/>'  
        '<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
        '<w:vanish/>'  # Makes text invisible but still searchable
        "</w:rPr>"
        '<w:t xml:space="preserve">{text}</w:t>'
        "</w:r>"
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
        z=251700000 + shape_id,
        sz=font_size_half_pt,
        text=escaped_text,
    )

    run_element = parse_xml(xml)
    paragraph._element.append(run_element)


# ── Visible text box for editable mode ────────────────────────────────────────

def _add_visible_textbox(
    paragraph,
    text: str,
    x_emu: int,
    y_emu: int,
    w_emu: int,
    h_emu: int,
    shape_id: int,
    font_name: str = "Arial",
    font_size_half_pt: int = 20,
    bold: bool = False,
    italic: bool = False,
    color_hex: str = "000000",
) -> None:
    """Add a visible, editable text box."""
    escaped_text = _escape(text)
    escaped_font = _escape(font_name)
    
    flags = ""
    if bold:
        flags += "<w:b/>"
    if italic:
        flags += "<w:i/>"
    
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
        "<w:r><w:rPr>"
        '<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
        "{flags}"
        '<w:color w:val="{color}"/>'
        '<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
        "</w:rPr>"
        '<w:t xml:space="preserve">{text}</w:t>'
        "</w:r>"
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
        z=251700000 + shape_id,
        font=escaped_font,
        flags=flags,
        color=color_hex,
        sz=font_size_half_pt,
        text=escaped_text,
    )

    run_element = parse_xml(xml)
    paragraph._element.append(run_element)


# ── Color conversion ──────────────────────────────────────────────────────────

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


# ── Page rendering ────────────────────────────────────────────────────────────

def _render_page_as_image(page: fitz.Page, dpi: int = 300) -> bytes:
    """Render entire page as a high-resolution PNG image."""
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    return pix.tobytes("png")


def _process_page_exact(
    pdf_doc: fitz.Document,
    word_doc: Document,
    page: fitz.Page,
    is_first: bool,
    dpi: int = 300,
    include_text_layer: bool = True,
) -> None:
    """
    Convert one PDF page using EXACT mode:
    - Render entire page as background image
    - Overlay invisible text for searchability
    """
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

    # One anchor paragraph per page
    anchor_para = word_doc.add_paragraph()
    anchor_para.paragraph_format.space_before = Pt(0)
    anchor_para.paragraph_format.space_after = Pt(0)

    # ── Render entire page as background image ────────────────────────────
    img_bytes = _render_page_as_image(page, dpi=dpi)
    
    _add_floating_image(
        word_doc,
        anchor_para,
        img_bytes,
        x_emu=0,
        y_emu=0,
        w_emu=w_emu,
        h_emu=h_emu,
        shape_id=_next_shape_id(),
        behind_doc=True,
    )

    # ── Overlay invisible text for searchability ──────────────────────────
    if include_text_layer:
        blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
        
        for block in blocks:
            if block["type"] != 0:  # 0 = text block
                continue
            
            for line in block["lines"]:
                for span in line["spans"]:
                    text = span["text"]
                    if not text or not text.strip():
                        continue
                    
                    bbox = span["bbox"]
                    size = span["size"]
                    
                    x_emu = _pt2emu(bbox[0])
                    y_emu = _pt2emu(bbox[1])
                    box_w = _pt2emu(bbox[2] - bbox[0])
                    box_h = _pt2emu(bbox[3] - bbox[1])
                    
                    # Ensure minimum size
                    box_w = max(box_w, _pt2emu(len(text) * size * 0.5))
                    box_h = max(box_h, _pt2emu(size * 1.2))
                    
                    size_half_pt = max(int(round(size * 2)), 2)
                    
                    _add_invisible_textbox(
                        anchor_para,
                        text,
                        x_emu,
                        y_emu,
                        box_w,
                        box_h,
                        _next_shape_id(),
                        font_size_half_pt=size_half_pt,
                    )


def _process_page_editable(
    pdf_doc: fitz.Document,
    word_doc: Document,
    page: fitz.Page,
    is_first: bool,
    dpi: int = 300,
) -> None:
    """
    Convert one PDF page using EDITABLE mode:
    - Extract and place images at exact positions
    - Place text as visible editable text boxes
    - Render complex graphics as images
    """
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

    anchor_para = word_doc.add_paragraph()
    anchor_para.paragraph_format.space_before = Pt(0)
    anchor_para.paragraph_format.space_after = Pt(0)

    # ── Extract and render graphics/images ────────────────────────────────
    # First, identify all drawable regions
    drawings = page.get_drawings()
    
    # Collect regions that need rasterization
    raster_regions: List[fitz.Rect] = []
    
    for drawing in drawings:
        draw_rect = drawing.get("rect")
        if draw_rect and draw_rect.width > 2 and draw_rect.height > 2:
            raster_regions.append(fitz.Rect(draw_rect))
    
    # Extract embedded images
    image_list = page.get_images(full=True)
    image_rects: List[fitz.Rect] = []
    
    for img_info in image_list:
        xref = img_info[0]
        try:
            img_rects = page.get_image_rects(xref)
            for img_rect in (img_rects or []):
                if img_rect.width > 0 and img_rect.height > 0:
                    image_rects.append(fitz.Rect(img_rect))
                    
                    base_image = pdf_doc.extract_image(xref)
                    if base_image and base_image.get("image"):
                        img_bytes = base_image["image"]
                        _add_floating_image(
                            word_doc,
                            anchor_para,
                            img_bytes,
                            _pt2emu(img_rect.x0),
                            _pt2emu(img_rect.y0),
                            _pt2emu(img_rect.width),
                            _pt2emu(img_rect.height),
                            _next_shape_id(),
                            behind_doc=True,
                        )
        except Exception:
            continue
    
    # Render complex graphics regions
    if raster_regions:
        # Merge overlapping regions
        merged = _merge_rects(raster_regions)
        
        for region in merged:
            # Skip if this region is just an image
            is_image_region = any(
                ir.contains(region) or region.contains(ir) 
                for ir in image_rects
            )
            if is_image_region:
                continue
            
            # Skip tiny regions
            if region.width < 10 or region.height < 10:
                continue
            
            try:
                zoom = min(dpi, 200) / 72.0
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat, clip=region, alpha=False)
                img_bytes = pix.tobytes("png")
                
                _add_floating_image(
                    word_doc,
                    anchor_para,
                    img_bytes,
                    _pt2emu(region.x0),
                    _pt2emu(region.y0),
                    _pt2emu(region.width),
                    _pt2emu(region.height),
                    _next_shape_id(),
                    behind_doc=True,
                )
            except Exception:
                continue

    # ── Extract and place text ────────────────────────────────────────────
    blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE | fitz.TEXT_PRESERVE_LIGATURES)["blocks"]

    for block in blocks:
        if block["type"] != 0:
            continue

        for line in block["lines"]:
            for span in line["spans"]:
                text = span["text"]
                if not text or text.isspace():
                    continue

                bbox = span["bbox"]
                font = span["font"]
                size = span["size"]
                flags = span["flags"]
                color = span["color"]

                is_bold = bool(flags & 2 ** 4)
                is_italic = bool(flags & 2 ** 1)
                
                size_half_pt = max(int(round(size * 2)), 2)
                color_hex = _color_to_hex(color)

                clean_font = font
                if "+" in clean_font:
                    clean_font = clean_font.split("+", 1)[1]

                x_emu = _pt2emu(bbox[0])
                y_emu = _pt2emu(bbox[1])
                
                pdf_width = bbox[2] - bbox[0]
                pdf_height = bbox[3] - bbox[1]
                
                # Add generous padding for text width
                box_w = _pt2emu(max(pdf_width * 1.3, len(text) * size * 0.7))
                box_h = _pt2emu(max(pdf_height, size * 1.6))

                _add_visible_textbox(
                    anchor_para,
                    text,
                    x_emu,
                    y_emu,
                    box_w,
                    box_h,
                    _next_shape_id(),
                    font_name=clean_font,
                    font_size_half_pt=size_half_pt,
                    bold=is_bold,
                    italic=is_italic,
                    color_hex=color_hex,
                )


def _merge_rects(rects: List[fitz.Rect], margin: float = 5) -> List[fitz.Rect]:
    """Merge overlapping or nearby rectangles."""
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
                other_exp = fitz.Rect(other)
                other_exp.x0 -= margin
                other_exp.y0 -= margin
                other_exp.x1 += margin
                other_exp.y1 += margin
                
                if current.intersects(other_exp):
                    current = current | other_exp
                    used[j] = True
                    changed = True
        
        used[i] = True
        merged.append(current)
    
    return merged


# ── Public API ───────────────────────────────────────────────────────────────

def convert_pdf_to_docx(
    pdf_path: str | Path,
    docx_path: Optional[str | Path] = None,
    *,
    pages: Optional[Sequence[int]] = None,
    dpi: int = 300,
    mode: Literal["exact", "editable"] = "exact",
    verbose: bool = False,
) -> Path:
    """Convert a PDF to a DOCX document.

    Parameters
    ----------
    pdf_path:
        Path to the source PDF file.
    docx_path:
        Destination path. Defaults to ``<input>.docx``.
    pages:
        0-based page indices. ``None`` → all pages.
    dpi:
        Resolution for rendering (higher = sharper but larger file).
        Default is 300 for print quality.
    mode:
        - "exact": Renders pages as images for perfect visual match (default)
        - "editable": Extracts text as editable boxes (may have layout issues)
    verbose:
        Print progress to stderr.
    
    Returns
    -------
    Path to the created DOCX file.
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
                f"  [{i + 1}/{total}] Processing page {idx + 1} ({mode} mode)…",
                file=sys.stderr,
            )
        
        if mode == "exact":
            _process_page_exact(
                pdf_doc, word_doc, page, 
                is_first=(i == 0), 
                dpi=dpi,
                include_text_layer=True,
            )
        else:
            _process_page_editable(
                pdf_doc, word_doc, page,
                is_first=(i == 0),
                dpi=dpi,
            )

    word_doc.save(str(docx_path))
    pdf_doc.close()

    if verbose:
        print("Done.", file=sys.stderr)

    return docx_path

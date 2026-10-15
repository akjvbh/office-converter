"""不依赖 Microsoft Office 的 docx → PDF 兜底渲染。

背景：docx→PDF 原本依赖 Word COM（见 converter._run_office_isolated）。在部分环境下
Word 自动化会挂起或不可用（Office 自动更新后行为变化、激活/许可对话框、没有可交互
桌面的会话等），此时转换会长时间无响应，用户拿不到任何结果。

本模块用 python-docx 解析文档结构生成 HTML，再交给 PyMuPDF 的排版引擎分页输出 PDF，
覆盖正文、标题、列表、图片、表格、分页符等常见元素。排版与 Word 不会逐像素一致
（字体、行距、浮动图形位置会退化），但内容完整、可直接阅读，作为 Word 不可用时的保底。
"""
from __future__ import annotations

import io
import os
import re
import sys

from docx import Document
from docx.oxml.ns import nsmap, qn

_EMU_PER_PT = 12700.0
_TWIPS_PER_PT = 20.0
_DEFAULT_SIZE_PT = 10.5
_DEFAULT_LINE_HEIGHT = 1.45

_MONO_HINTS = ("courier", "consolas", "menlo", "mono", "code")
_SANS_HINTS = ("黑体", "simhei", "微软雅黑", "microsoft yahei", "雅黑", "等线", "dengxian",
               "思源黑", "arial", "helvetica", "verdana", "tahoma", "calibri", "segoe",
               "euclid", "fzhei", "sans")
_SERIF_HINTS = ("宋体", "simsun", "songti", "仿宋", "fangsong", "楷体", "kaiti", "times",
                "cambria", "georgia", "garamond", "book antiqua", "bookman", "palatino",
                "constantia", "serif", "symbol")

# 可被 PDF 直接使用的位图格式；WMF/EMF 这类矢量图无法渲染，跳过并计数
_RASTER_MIME = {"image/png", "image/jpeg", "image/gif", "image/bmp", "image/tiff",
                "image/webp"}


def _xp(el, path: str):
    """在 python-docx 元素上执行 xpath，失败时返回空列表。"""
    if el is None:
        return []
    try:
        return el.xpath(path)
    except Exception:
        try:
            return el.xpath(path, namespaces=nsmap)
        except Exception:
            return []


def _attr(el, name: str, default=None):
    if el is None:
        return default
    return el.get(qn("w:" + name), default)


def _pt_from_twips(value) -> float:
    return float(value) / _TWIPS_PER_PT


def _pt_from_emu(value) -> float:
    return float(value) / _EMU_PER_PT


def _css_family(names) -> str:
    for name in names:
        if not name:
            continue
        low = name.lower()
        if any(hint in low for hint in _MONO_HINTS):
            return "monospace"
        if any(hint in low for hint in _SANS_HINTS):
            return "sans-serif"
        if any(hint in low for hint in _SERIF_HINTS):
            return "serif"
    return "serif"


_CN_DIGITS = "零一二三四五六七八九"


def _chinese_numeral(value: int) -> str:
    """中文数字（支持到 9999，足够覆盖常见标题编号）。"""
    if value <= 0:
        return str(value)
    if value < 10:
        return _CN_DIGITS[value]
    if value < 20:
        return "十" + (_CN_DIGITS[value % 10] if value % 10 else "")
    if value < 100:
        out = _CN_DIGITS[value // 10] + "十"
        return out + (_CN_DIGITS[value % 10] if value % 10 else "")
    if value < 1000:
        out = _CN_DIGITS[value // 100] + "百"
        rem = value % 100
        if rem == 0:
            return out
        if rem < 10:
            return out + "零" + _CN_DIGITS[rem]
        return out + _chinese_numeral(rem)
    out = _CN_DIGITS[value // 1000] + "千"
    rem = value % 1000
    if rem == 0:
        return out
    if rem < 100:
        return out + "零" + _chinese_numeral(rem)
    return out + _chinese_numeral(rem)


def _roman_numeral(value: int) -> str:
    table = ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
             (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"))
    out = []
    for num, sym in table:
        while value >= num:
            out.append(sym)
            value -= num
    return "".join(out) or "I"


def _letters(value: int, upper: bool) -> str:
    out = []
    while value > 0:
        value, rem = divmod(value - 1, 26)
        out.append(chr((65 if upper else 97) + rem))
    return "".join(reversed(out))


def _style_chain(style):
    seen = set()
    while style is not None and id(style) not in seen:
        seen.add(id(style))
        yield style
        try:
            style = style.base_style
        except Exception:
            style = None


def _style_prop(style, path: str):
    """沿样式继承链取第一个命中的属性节点。"""
    for item in _style_chain(style):
        found = _xp(getattr(item, "element", None), path)
        if found:
            return found[0]
    return None


class _Context:
    """一次渲染过程中的共享状态。"""

    def __init__(self, doc):
        self.doc = doc
        self.archive = None
        self.image_index = 0
        self.counters: dict = {}
        self.numbering: dict = self._numbering_formats()
        self.default_size_pt = self._default_size()
        self.skipped_images: list = []
        self.missing_images = 0
        self.content_width_pt = 468.0

    def _default_size(self) -> float:
        try:
            found = _xp(self.doc.styles.element,
                        "./w:docDefaults/w:rPrDefault/w:rPr/w:sz")
            if found:
                return int(_attr(found[0], "val", 21)) / 2.0
        except Exception:
            pass
        return _DEFAULT_SIZE_PT

    def _numbering_formats(self) -> dict:
        """解析 numbering.xml：numId → {ilvl: {fmt, text, left, hanging, start}}"""
        result: dict = {}
        try:
            part = getattr(self.doc.part, "numbering_part", None)
        except Exception:
            part = None
        if part is None:
            return result
        try:
            root = part.element
            abstracts = {}
            for abstract in _xp(root, "./w:abstractNum"):
                aid = _attr(abstract, "abstractNumId")
                levels = {}
                for lvl in _xp(abstract, "./w:lvl"):
                    ilvl = int(_attr(lvl, "ilvl", 0))
                    fmt_el = _xp(lvl, "./w:numFmt")
                    text_el = _xp(lvl, "./w:lvlText")
                    ind_el = _xp(lvl, "./w:ind")
                    start_el = _xp(lvl, "./w:start")
                    levels[ilvl] = {
                        "fmt": _attr(fmt_el[0], "val", "decimal") if fmt_el else "decimal",
                        "text": _attr(text_el[0], "val", "%1.") if text_el else "%1.",
                        "left": _attr(ind_el[0], "left") if ind_el else None,
                        "hanging": _attr(ind_el[0], "hanging") if ind_el else None,
                        "start": int(_attr(start_el[0], "val", 1)) if start_el else 1,
                    }
                abstracts[aid] = levels
            for num in _xp(root, "./w:num"):
                num_id = _attr(num, "numId")
                ref = _xp(num, "./w:abstractNumId")
                if not ref:
                    continue
                levels = dict(abstracts.get(_attr(ref[0], "val"), {}))
                for override in _xp(num, "./w:lvlOverride"):
                    ilvl = int(_attr(override, "ilvl", 0))
                    start = _xp(override, "./w:startOverride")
                    if start and ilvl in levels:
                        levels[ilvl] = dict(levels[ilvl],
                                            start=int(_attr(start[0], "val", 1)))
                result[num_id] = levels
        except Exception:
            return result
        return result

    def add_image(self, blob: bytes, content_type: str, ext_hint: str = "png") -> str:
        """把图片写进 PyMuPDF 的 Archive，返回 HTML 里可引用的名字。"""
        import pymupdf

        suffix = {
            "image/png": "png", "image/jpeg": "jpg", "image/gif": "gif",
            "image/bmp": "bmp", "image/tiff": "tif", "image/webp": "webp",
        }.get(content_type, ext_hint)
        self.image_index += 1
        name = "img%03d.%s" % (self.image_index, suffix)
        if self.archive is None:
            self.archive = pymupdf.Archive()
        self.archive.add(blob, name)
        return name


def _escape(text: str) -> str:
    return (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace("\u00a0", "&nbsp;"))


def _font_names(r_el, style) -> list:
    names = []
    sources = list(_xp(r_el, "./w:rPr/w:rFonts"))
    style_fonts = _style_prop(style, "./w:rPr/w:rFonts")
    if style_fonts is not None:
        sources.append(style_fonts)
    for fonts_el in sources:
        for key in ("eastAsia", "ascii", "hAnsi", "cs"):
            value = _attr(fonts_el, key)
            if value:
                names.append(value)
    return names


def _run_size_pt(r_el, style, ctx: _Context) -> float:
    own = _xp(r_el, "./w:rPr/w:sz")
    if own:
        try:
            return int(_attr(own[0], "val")) / 2.0
        except Exception:
            pass
    for item in _style_chain(style):
        found = _xp(getattr(item, "element", None), "./w:rPr/w:sz")
        if found:
            try:
                return int(_attr(found[0], "val")) / 2.0
            except Exception:
                break
    return ctx.default_size_pt


def _run_flags(r_el, style) -> dict:
    def flag(tag: str) -> bool:
        own = _xp(r_el, "./w:rPr/w:" + tag)
        if own:
            return _attr(own[0], "val", "1") not in ("0", "false", "off")
        style_flag = _style_prop(style, "./w:rPr/w:" + tag)
        if style_flag is not None:
            return _attr(style_flag, "val", "1") not in ("0", "false", "off")
        return False

    vert = _xp(r_el, "./w:rPr/w:vertAlign")
    if not vert:
        style_vert = _style_prop(style, "./w:rPr/w:vertAlign")
        vert = [style_vert] if style_vert is not None else []
    vert_val = _attr(vert[0], "val") if vert else None

    color = None
    color_el = _xp(r_el, "./w:rPr/w:color")
    if color_el:
        color = _attr(color_el[0], "val")
    if color in ("auto", None):
        style_color = _style_prop(style, "./w:rPr/w:color")
        color = _attr(style_color, "val") if style_color is not None else None
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", color or ""):
        color = None

    return {
        "bold": flag("b"),
        "italic": flag("i"),
        "underline": flag("u"),
        "strike": flag("strike"),
        "vert": vert_val if vert_val in ("superscript", "subscript") else None,
        "color": color,
    }


def _image_html(element, ctx: _Context, doc_part) -> str:
    """把 w:drawing / w:pict 里的图片渲染成 <img>，失败返回空串。"""
    blips = _xp(element, ".//a:blip") or _xp(element, ".//v:imagedata")
    if not blips:
        return ""
    rid = None
    for blip in blips:
        rid = (blip.get(qn("r:embed")) or blip.get(qn("r:id")))
        if rid:
            break
    if not rid:
        ctx.missing_images += 1
        return ""
    try:
        part = doc_part.related_parts[rid]
    except Exception:
        ctx.missing_images += 1
        return ""
    content_type = getattr(part, "content_type", "") or ""
    if content_type not in _RASTER_MIME:
        ctx.skipped_images.append(content_type or "unknown")
        return ""

    width_pt = height_pt = None
    extents = _xp(element, ".//wp:extent")
    if extents:
        try:
            width_pt = _pt_from_emu(int(extents[0].get("cx")))
            height_pt = _pt_from_emu(int(extents[0].get("cy")))
        except Exception:
            width_pt = height_pt = None
    if width_pt is None:
        style_attr = ""
        shapes = _xp(element, ".//v:shape")
        if shapes:
            style_attr = shapes[0].get("style") or ""
        if style_attr:
            match = re.search(r"width:\s*([\d.]+)pt", style_attr)
            if match:
                width_pt = float(match.group(1))
            match = re.search(r"height:\s*([\d.]+)pt", style_attr)
            if match:
                height_pt = float(match.group(1))
    limit = ctx.content_width_pt
    if width_pt and limit > 0 and width_pt > limit:
        ratio = limit / width_pt
        width_pt = limit
        if height_pt:
            height_pt = height_pt * ratio

    name = ctx.add_image(part.blob, content_type)
    style = ["max-width:100%"]
    if width_pt:
        style.append("width:%.1fpt" % width_pt)
    if height_pt:
        style.append("height:%.1fpt" % height_pt)
    return '<img src="%s" style="%s"/>' % (name, ";".join(style))


def _run_html(r_el, style, ctx: _Context, doc_part) -> str:
    flags = _run_flags(r_el, style)
    size = _run_size_pt(r_el, style, ctx)
    css = ["font-size:%.2fpt" % size, "font-family:%s" % _css_family(_font_names(r_el, style))]
    if flags["bold"]:
        css.append("font-weight:bold")
    if flags["italic"]:
        css.append("font-style:italic")
    if flags["underline"]:
        css.append("text-decoration:underline")
    if flags["strike"]:
        css.append("text-decoration:line-through")
    if flags["color"]:
        css.append("color:#%s" % flags["color"])
    if flags["vert"] == "superscript":
        css.append("vertical-align:super")
        css.append("font-size:%.2fpt" % max(4.0, size * 0.65))
    elif flags["vert"] == "subscript":
        css.append("vertical-align:sub")
        css.append("font-size:%.2fpt" % max(4.0, size * 0.65))

    pieces = []
    for child in r_el.iterchildren():
        tag = child.tag
        if tag == qn("w:t"):
            pieces.append(_escape(child.text or ""))
        elif tag == qn("w:tab"):
            pieces.append('<span style="white-space:pre">\t</span>')
        elif tag == qn("w:br"):
            if child.get(qn("w:type")) == "page":
                pieces.append('<span style="page-break-after:always"></span>')
            else:
                pieces.append("<br/>")
        elif tag == qn("w:noBreakHyphen"):
            pieces.append("-")
        elif tag in (qn("w:drawing"), qn("w:pict"), qn("w:object")):
            pieces.append(_image_html(child, ctx, doc_part))
        elif tag in (qn("w:instrText"), qn("w:fldChar"), qn("w:sym")):
            continue
    text = "".join(pieces)
    if not text:
        return ""
    return '<span style="%s">%s</span>' % (";".join(css), text)


def _runs_html(el, style, ctx: _Context, doc_part, in_link: bool = False) -> str:
    """把段落内的 run/hyperlink 按顺序转成 HTML。"""
    out = []
    for child in el.iterchildren():
        tag = child.tag
        if tag == qn("w:r"):
            html = _run_html(child, style, ctx, doc_part)
            if html and in_link:
                out.append('<a style="color:#0000ee;text-decoration:underline">%s</a>' % html)
            else:
                out.append(html)
        elif tag == qn("w:hyperlink"):
            out.append(_runs_html(child, style, ctx, doc_part, in_link=True))
        elif tag == qn("w:sdt"):
            for content in _xp(child, "./w:sdtContent"):
                out.append(_runs_html(content, style, ctx, doc_part, in_link))
        elif tag in (qn("w:smartTag"), qn("w:dir"), qn("w:bdo"), qn("w:ins")):
            out.append(_runs_html(child, style, ctx, doc_part, in_link))
    return "".join(part for part in out if part)


_HEADING_RE = re.compile(r"(heading\s*([1-6])|标题\s*([一二三四五六1-6])|title)", re.I)
_HEADING_CN = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6}


def _heading_level(p, pPr):
    for source in (pPr, getattr(p.style, "element", None)):
        found = _xp(source, "./w:outlineLvl")
        if found:
            try:
                return min(6, max(1, int(_attr(found[0], "val", 0)) + 1))
            except Exception:
                pass
    name = (p.style.name or "") if p.style is not None else ""
    match = _HEADING_RE.search(name)
    if not match:
        return None
    if match.group(2):
        return int(match.group(2))
    if match.group(3):
        token = match.group(3)
        return int(token) if token.isdigit() else _HEADING_CN.get(token, 2)
    return 1


def _paragraph_css(p, pPr, style, ctx: _Context):
    """返回 (CSS 片段, 段落信息)。"""
    css = []
    info = {"align": None, "first_line_pt": None, "left_pt": None}

    own_jc = _xp(pPr, "./w:jc")
    if own_jc:
        jc = _attr(own_jc[0], "val")
    else:
        style_jc = _style_prop(style, "./w:pPr/w:jc")
        jc = _attr(style_jc, "val") if style_jc is not None else None
    align_map = {"center": "center", "right": "right", "both": "justify",
                 "distribute": "justify", "left": "left", "justifyLow": "justify"}
    if jc in align_map:
        info["align"] = align_map[jc]
        css.append("text-align:%s" % align_map[jc])

    ind = _xp(pPr, "./w:ind")
    if not ind:
        style_ind = _style_prop(style, "./w:pPr/w:ind")
        ind = [style_ind] if style_ind is not None else []
    if ind:
        el = ind[0]
        first = _attr(el, "firstLine")
        hanging = _attr(el, "hanging")
        left = _attr(el, "left") or _attr(el, "start")
        if first:
            info["first_line_pt"] = _pt_from_twips(first)
            css.append("text-indent:%.2fpt" % info["first_line_pt"])
        elif hanging:
            info["first_line_pt"] = -_pt_from_twips(hanging)
            css.append("text-indent:%.2fpt" % info["first_line_pt"])
        if left:
            info["left_pt"] = _pt_from_twips(left)
            css.append("margin-left:%.2fpt" % info["left_pt"])

    def apply_spacing(el) -> bool:
        if el is None:
            return False
        applied = False
        line = _attr(el, "line")
        rule = _attr(el, "lineRule", "auto")
        if line and line.isdigit():
            try:
                if rule in ("exact", "atLeast"):
                    css.append("line-height:%.2fpt" % (float(line) / _TWIPS_PER_PT))
                else:
                    css.append("line-height:%.3f" % (float(line) / 240.0))
                applied = True
            except Exception:
                pass
        before = _attr(el, "before")
        if before and before.isdigit() and int(before) > 0:
            css.append("margin-top:%.2fpt" % _pt_from_twips(before))
        after = _attr(el, "after")
        if after and after.isdigit() and int(after) > 0:
            css.append("margin-bottom:%.2fpt" % _pt_from_twips(after))
        return applied

    own_spacing = _xp(pPr, "./w:spacing")
    if not apply_spacing(own_spacing[0] if own_spacing else None):
        apply_spacing(_style_prop(style, "./w:pPr/w:spacing"))
    return ";".join(css), info


def _list_marker(pPr, ctx: _Context, default_indent: bool):
    """返回 (标记 HTML, 额外 CSS)。"""
    num_pr = _xp(pPr, "./w:numPr")
    if not num_pr:
        return "", ""
    ilvl_el = _xp(num_pr[0], "./w:ilvl")
    num_id_el = _xp(num_pr[0], "./w:numId")
    ilvl = int(_attr(ilvl_el[0], "val", 0)) if ilvl_el else 0
    num_id = _attr(num_id_el[0], "val") if num_id_el else None
    if num_id in (None, "0"):
        return "", ""
    levels = ctx.numbering.get(num_id, {})
    spec = levels.get(ilvl) or levels.get(0) or {"fmt": "bullet", "text": "•", "start": 1}
    fmt = spec.get("fmt", "bullet")

    if fmt == "bullet":
        marker = spec.get("text") or "•"
        if marker in ("\uf0b7", "\uf0a7", "\uf0d8", "o", ""):
            marker = "•"
        marker = _escape(marker[:2])
    else:
        key = (num_id, ilvl)
        ctx.counters[key] = ctx.counters.get(key, int(spec.get("start", 1)) - 1) + 1
        for other_key in list(ctx.counters):
            if other_key[0] == num_id and other_key[1] > ilvl:
                del ctx.counters[other_key]
        value = ctx.counters[key]
        if fmt == "decimalZero":
            text = "%02d" % value
        elif fmt == "lowerLetter":
            text = _letters(value, upper=False)
        elif fmt == "upperLetter":
            text = _letters(value, upper=True)
        elif fmt == "lowerRoman":
            text = _roman_numeral(value).lower()
        elif fmt == "upperRoman":
            text = _roman_numeral(value)
        elif fmt in ("chineseCounting", "chineseCountingThousand", "chineseLegalSimplified",
                     "ideographDigital", "ideographTraditional"):
            text = _chinese_numeral(value)
        else:
            text = str(value)
        pattern = spec.get("text") or "%1."
        marker = _escape(re.sub(r"%(\d)", lambda m: text if m.group(1) == "1" else "",
                                pattern))

    css = []
    left = spec.get("left")
    hanging = spec.get("hanging")
    if left:
        css.append("margin-left:%.2fpt" % _pt_from_twips(left))
    elif default_indent:
        css.append("margin-left:%.2fpt" % (18.0 * (ilvl + 1)))
    if hanging:
        css.append("text-indent:%.2fpt" % (-_pt_from_twips(hanging)))
    elif default_indent:
        css.append("text-indent:-18.00pt")
    width = _pt_from_twips(hanging) if hanging else 18.0
    return ('<span style="display:inline-block;min-width:%.2fpt">%s</span>'
            % (width, marker)), ";".join(css)


def _paragraph_html(p, ctx: _Context, doc_part) -> str:
    pPr = p._element.pPr
    style = p.style
    css, info = _paragraph_css(p, pPr, style, ctx)
    level = _heading_level(p, pPr)
    if level:
        css = ";".join(filter(None, [css, "margin:8pt 0 4pt 0"]))
        tag = "h%d" % level
    else:
        tag = "p"

    marker_html, marker_css = _list_marker(pPr, ctx, default_indent=info["left_pt"] is None)
    if marker_css:
        css = ";".join(filter(None, [css, marker_css]))

    inner = _runs_html(p._element, style, ctx, doc_part)
    if not inner:
        if not marker_html:
            return '<%s style="%s">&nbsp;</%s>' % (tag, css, tag)
        inner = marker_html
    elif marker_html:
        inner = marker_html + inner

    if not css:
        return "<%s>%s</%s>" % (tag, inner, tag)
    return '<%s style="%s">%s</%s>' % (tag, css, inner, tag)


def _table_pr(tc):
    parent = tc.getparent()
    while parent is not None and parent.tag != qn("w:tbl"):
        parent = parent.getparent()
    if parent is None:
        return None
    found = _xp(parent, "./w:tblPr")
    return found[0] if found else None


def _cell_border_css(tc_pr, tbl_pr, default_visible: bool) -> list:
    css = []
    tc_borders = _xp(tc_pr, "./w:tcBorders")
    source = tc_borders[0] if tc_borders else None
    if source is None and tbl_pr is not None:
        table_borders = _xp(tbl_pr, "./w:tblBorders")
        source = table_borders[0] if table_borders else None
    for side in ("top", "left", "bottom", "right"):
        el = _xp(source, "./w:" + side) if source is not None else []
        if el:
            val = _attr(el[0], "val", "single")
            if val in ("nil", "none"):
                css.append("border-%s:none" % side)
            else:
                size = _attr(el[0], "sz")
                width = max(0.5, float(size) / 8.0) if size and size.isdigit() else 0.75
                css.append("border-%s:%.2fpt solid #000" % (side, width))
        elif default_visible:
            css.append("border-%s:0.75pt solid #000" % side)
    return css


def _cell_html(tc, ctx: _Context, doc_part, default_visible: bool) -> str:
    tc_prs = _xp(tc, "./w:tcPr")
    tc_pr = tc_prs[0] if tc_prs else None
    colspan = 1
    grid_span = _xp(tc_pr, "./w:gridSpan")
    if grid_span:
        try:
            colspan = max(1, int(_attr(grid_span[0], "val", 1)))
        except Exception:
            colspan = 1

    css = ["vertical-align:top", "padding:2pt 4pt"]
    shd = _xp(tc_pr, "./w:shd")
    if shd:
        fill = _attr(shd[0], "fill")
        if fill and re.fullmatch(r"[0-9A-Fa-f]{6}", fill) and fill.lower() != "ffffff":
            css.append("background-color:#%s" % fill)
    css.extend(_cell_border_css(tc_pr, _table_pr(tc), default_visible))

    va = _xp(tc_pr, "./w:vAlign")
    if va:
        val = _attr(va[0], "val")
        if val == "center":
            css.append("vertical-align:middle")
        elif val == "bottom":
            css.append("vertical-align:bottom")

    inner = _container_html(tc, ctx, doc_part)
    attrs = ' colspan="%d"' % colspan if colspan > 1 else ""
    return '<td%s style="%s">%s</td>' % (attrs, ";".join(css), inner or "&nbsp;")


def _table_html(tbl, ctx: _Context, doc_part) -> str:
    tbl_prs = _xp(tbl, "./w:tblPr")
    tbl_pr = tbl_prs[0] if tbl_prs else None
    default_visible = True
    if tbl_pr is not None:
        borders = _xp(tbl_pr, "./w:tblBorders")
        if borders:
            default_visible = any(
                _attr(_xp(borders[0], "./w:" + side)[0], "val", "single")
                not in ("nil", "none")
                for side in ("top", "left", "bottom", "right", "insideH", "insideV")
                if _xp(borders[0], "./w:" + side))

    widths = []
    for col in _xp(tbl, "./w:tblGrid/w:gridCol"):
        try:
            widths.append(float(_attr(col, "w", 0)))
        except Exception:
            widths.append(0.0)
    total = sum(widths)
    colgroup = ""
    if total > 0:
        colgroup = "<colgroup>" + "".join(
            '<col style="width:%.2f%%"/>' % (w / total * 100.0) for w in widths
        ) + "</colgroup>"

    rows_html = []
    for row in _xp(tbl, "./w:tr"):
        rendered = []
        for cell in _xp(row, "./w:tc"):
            tc_prs = _xp(cell, "./w:tcPr")
            tc_pr = tc_prs[0] if tc_prs else None
            vmerge = _xp(tc_pr, "./w:vMerge")
            if vmerge and _attr(vmerge[0], "val", "continue") in ("continue", None):
                continue
            rendered.append(_cell_html(cell, ctx, doc_part, default_visible))
        rows_html.append("<tr>%s</tr>" % "".join(rendered))

    return '<table style="border-collapse:collapse;width:100%%;margin:6pt 0">%s%s</table>' % (
        colgroup, "".join(rows_html))


def _container_html(container, ctx: _Context, doc_part) -> str:
    """按文档顺序渲染段落/表格（正文与表格单元格共用）。"""
    from docx.text.paragraph import Paragraph

    parts = []
    for child in container.iterchildren():
        tag = child.tag
        if tag == qn("w:p"):
            parts.append(_paragraph_html(Paragraph(child, ctx.doc), ctx, doc_part))
        elif tag == qn("w:tbl"):
            parts.append(_table_html(child, ctx, doc_part))
        elif tag == qn("w:sdt"):
            for content in _xp(child, "./w:sdtContent"):
                parts.append(_container_html(content, ctx, doc_part))
    return "".join(part for part in parts if part)


_BASE_CSS = (
    "* {font-family: serif; font-size: %.2fpt; line-height: %.2f; color: #000000;}"
    "p {margin: 0 0 2pt 0;}"
    "table {border-collapse: collapse;}"
    "td {vertical-align: top;}"
    "img {max-width: 100%%;}"
    "a {color: #0000ee;}"
) % (_DEFAULT_SIZE_PT, _DEFAULT_LINE_HEIGHT)


def _page_geometry(doc):
    """返回 (页宽pt, 页高pt, 上边距, 右边距, 下边距, 左边距)。"""
    try:
        section = doc.sections[0]
        width = _pt_from_emu(section.page_width)
        height = _pt_from_emu(section.page_height)
        top = _pt_from_emu(section.top_margin)
        right = _pt_from_emu(section.right_margin)
        bottom = _pt_from_emu(section.bottom_margin)
        left = _pt_from_emu(section.left_margin)
    except Exception:
        width, height, top, right, bottom, left = 595.0, 842.0, 56.0, 56.0, 56.0, 56.0
    if not (100 < width < 3000) or not (100 < height < 3000):
        width, height = 595.0, 842.0
    if not (0 < left < width / 2) or not (0 < right < width / 2):
        left = right = min(56.0, width / 8.0)
    if not (0 < top < height / 2) or not (0 < bottom < height / 2):
        top = bottom = min(56.0, height / 8.0)
    return width, height, top, right, bottom, left


def build_html(docx_path: str):
    """把 docx 解析为 HTML；返回 (html, 页面几何, 上下文)。"""
    doc = Document(docx_path)
    width, height, top, right, bottom, left = _page_geometry(doc)
    ctx = _Context(doc)
    ctx.content_width_pt = max(120.0, width - left - right)
    body = _container_html(doc.element.body, ctx, doc.part)
    html = ("<html><head><meta charset='utf-8'/></head><body>%s</body></html>" % body)
    return html, (width, height, top, right, bottom, left), ctx


def render_to_pdf(docx_path: str) -> bytes:
    """docx → PDF 字节流。"""
    import pymupdf

    html, (width, height, top, right, bottom, left), ctx = build_html(docx_path)
    mediabox = pymupdf.Rect(0, 0, width, height)
    content = pymupdf.Rect(left, top, width - right, height - bottom)

    buffer = io.BytesIO()
    writer = pymupdf.DocumentWriter(buffer)
    story = pymupdf.Story(html=html, user_css=_BASE_CSS, archive=ctx.archive)
    story.write(writer, lambda rect_num, filled: (mediabox, content, None))
    writer.close()

    data = buffer.getvalue()
    doc = pymupdf.open(stream=data, filetype="pdf")
    try:
        if doc.page_count == 0:
            raise RuntimeError("排版引擎未生成任何页面")
        text_len = sum(len(page.get_text().strip()) for page in doc)
        if text_len == 0 and ctx.image_index == 0:
            raise RuntimeError("排版引擎未输出任何内容")
        try:
            doc.subset_fonts()
        except Exception:
            pass
        out = io.BytesIO()
        doc.save(out, garbage=3, deflate=True)
        return out.getvalue()
    finally:
        doc.close()


def convert(src: str, dst: str) -> None:
    """docx → PDF 写入文件（供 converter 调用）。"""
    data = render_to_pdf(src)
    with open(dst, "wb") as fh:
        fh.write(data)


def text_to_pdf_bytes(text: str) -> bytes:
    """纯文本兜底：解析失败时至少把文字落到 PDF 上。"""
    import pymupdf

    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)
    box = pymupdf.Rect(56, 56, 539, 786)
    remaining = text or "（文档内容为空）"
    while remaining:
        result = page.insert_textbox(box, remaining, fontsize=10.5, fontname="china-s",
                                     align=0, lineheight=1.4)
        if result >= 0:
            break
        per_page = int(box.height / (10.5 * 1.5))
        lines = remaining.split("\n")
        page.insert_textbox(box, "\n".join(lines[:per_page]), fontsize=10.5,
                            fontname="china-s", align=0, lineheight=1.4)
        remaining = "\n".join(lines[per_page:])
        if remaining:
            page = doc.new_page(width=595, height=842)
    out = io.BytesIO()
    doc.save(out)
    doc.close()
    return out.getvalue()


def _main() -> int:
    if len(sys.argv) < 3:
        print("用法: python docx_pdf.py <输入.docx> <输出.pdf>")
        return 2
    src, dst = sys.argv[1], sys.argv[2]
    _html, geometry, ctx = build_html(src)
    print("页面: %.1f x %.1f pt，包含图片 %d 张，跳过 %d 张" % (
        geometry[0], geometry[1], ctx.image_index, len(ctx.skipped_images)))
    convert(src, dst)
    print("已输出 %s (%d 字节)" % (dst, os.path.getsize(dst)))
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())

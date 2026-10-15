"""英文文献翻译：段落抽取 + LLM 翻译 + 中英对照 / 纯译文 Word 生成。

设计要点
- 只用项目已安装的库（PyMuPDF、python-docx、httpx），不新增依赖。
- 走 OpenAI 兼容的 /chat/completions 接口，因此 DeepSeek、OpenAI、通义千问等都能直接换用。
- 分段批量翻译并用分隔符对齐；某批解析失败时自动降级为逐段翻译，保证原文与译文段落一一对应。
"""
from __future__ import annotations

import os
import re
import time
import html as html_module
from dataclasses import dataclass, field, replace

# 段落分隔符：选用正文几乎不可能出现的字符串，降低串段风险
DELIM = "<<<PARA>>>"

# 每页需要保留的正文块最小长度，用于过滤页码、页眉等噪声
_MIN_PARA_CHARS = 2

# 原位翻译时允许的最小字号；低于它就宁可保留原文，也不让译文压住别的文字
# （用户反馈"保留英文太多"后调低：小框宁可把字缩小，也尽量给出译文）
_MIN_FIT_SIZE = 4.0

# 译文行高。中文方块字比英文更"满"，行距太紧就会看着堆成一团。
_TRANSLATE_LINE_HEIGHT = 1.25

# 适配字号时要求的最小余量：按字号比例另加固定下限，
# 让译文四周留出空白，而不是顶满整个框。
_FIT_SPARE_RATIO = 0.15
_FIT_SPARE_MIN = 0.6

# 是否把"横向几乎不重叠的相邻行"当作两段（并行分栏、表格相邻单元格）。
# 关掉就退回到旧行为：整页粘成一段，译文会压到旁边那一栏的文字上。
_SPLIT_DISJOINT_LINES = True

# 译文与"未翻译内容"（页眉、公式行、表格数字）的重叠容忍度：超过就放弃这一段，
# 保留原文英文，避免压字。
_PROTECTED_OVERLAP_RATIO = 0.50

# 译文比原文长时可以向下借用的空白上限（相对原框高度的倍数）。
# 借得越多，能翻译的段落越多，但越容易撞到图或别的文字（撞到会被放弃）。
_GROW_FACTOR = 2.2

_REF_HEADING = re.compile(
    r"^\s*(references|reference|bibliography|literature\s+cited|works\s+cited|参考文献)\s*[:：]?\s*$",
    re.IGNORECASE,
)

# 形如 "12" / "Page 3 of 10" / "12 | Journal of ..." 的页眉页脚噪声
_PAGE_NOISE = re.compile(r"^(page\s*)?\d{1,4}(\s*(of|/)\s*\d{1,4})?$", re.IGNORECASE)

_SENTENCE_SPLIT = re.compile(r"(?<=[.!?;:])\s+")

# 行尾的句末标点（中英文）
_ENDS_SENTENCE = re.compile(r"[.。！？!?;；:：]\s*$")

_CJK = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")

# 公式特征字符，用于识别不该被翻译/覆盖的数学排版
_MATH_CHARS = set("=+−-×÷±≈≤≥≠∑∏∫√∞∂∇∈∉⊂⊆∪∩→←↔∀∃"
                  "αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ")


def _median(values: list[float]) -> float:
    if not values:
        return 0.0
    ordered = sorted(values)
    mid = len(ordered) // 2
    return ordered[mid] if len(ordered) % 2 else (ordered[mid - 1] + ordered[mid]) / 2


def _join_lines(parts: list[str]) -> str:
    """拼接同一段内的行。

    中文换行不该补空格，英文换行必须补空格；英文行尾的连字符断词要还原。
    """
    result = ""
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if not result:
            result = part
            continue
        if result.endswith("-"):
            result = result[:-1] + part
        elif _CJK.search(result[-1]) and _CJK.search(part[0]):
            result += part
        else:
            result += " " + part
    return result


# --------------------------------------------------------------------------
# 配置
# --------------------------------------------------------------------------
@dataclass
class TranslationConfig:
    """翻译配置；api_key 为空表示尚未配置，接口应给出明确提示。"""

    api_key: str = ""
    base_url: str = "https://api.deepseek.com/v1"
    model: str = "deepseek-chat"
    temperature: float = 1.0
    timeout: float = 180.0
    target_lang: str = "简体中文"
    domain: str = ""          # 学科/领域提示，例如「水产养殖、鱼类行为学」
    glossary: str = ""        # 术语表，每行「英文=中文」或自由文本
    translate_names: bool = False   # 人名是否译成中文；默认保留原文拼音（猜汉字容易出错）
    skip_references: bool = True
    max_chars: int = 3500     # 单次请求携带的原文上限
    retries: int = 3
    workers: int = 1          # 同一文档内并发请求数；默认串行以规避限流

    @property
    def configured(self) -> bool:
        return bool(self.api_key.strip())

    def endpoint(self) -> str:
        return self.base_url.rstrip("/") + "/chat/completions"

    def with_overrides(self, **kw) -> "TranslationConfig":
        clean = {k: v for k, v in kw.items() if v not in (None, "")}
        return replace(self, **clean) if clean else self


def config_from_env(env: dict | None = None) -> TranslationConfig:
    """从环境变量构造配置；命名兼容常见的几种写法。"""
    env = env if env is not None else os.environ

    api_key = (
        env.get("TRANSLATE_API_KEY")
        or env.get("DEEPSEEK_API_KEY")
        or env.get("OPENAI_API_KEY")
        or ""
    ).strip()

    base_url = (env.get("TRANSLATE_BASE_URL") or "").strip()
    if not base_url:
        # 未显式指定时，按密钥来源推断默认端点
        base_url = (
            "https://api.openai.com/v1"
            if env.get("OPENAI_API_KEY") and not env.get("DEEPSEEK_API_KEY")
            else "https://api.deepseek.com/v1"
        )

    model = (env.get("TRANSLATE_MODEL") or "").strip()
    if not model:
        model = "gpt-4o-mini" if "api.openai.com" in base_url else "deepseek-chat"

    def _num(name: str, default):
        raw = env.get(name)
        if raw in (None, ""):
            return default
        try:
            return type(default)(raw)
        except (TypeError, ValueError):
            return default

    return TranslationConfig(
        api_key=api_key,
        base_url=base_url,
        model=model,
        temperature=_num("TRANSLATE_TEMPERATURE", 1.0),
        timeout=_num("TRANSLATE_TIMEOUT", 180.0),
        target_lang=(env.get("TRANSLATE_TARGET_LANG") or "简体中文").strip(),
        skip_references=str(env.get("TRANSLATE_SKIP_REFERENCES", "1")).lower()
        not in ("0", "false", "no"),
        max_chars=_num("TRANSLATE_MAX_CHARS", 3500),
        workers=max(1, _num("TRANSLATE_WORKERS", 1)),
    )


# --------------------------------------------------------------------------
# 段落抽取
# --------------------------------------------------------------------------
def _clean_text(text: str) -> str:
    """合并断行与多余空白；保留段落内正常空格。"""
    text = text.replace("\u00ad", "")            # 软连字符
    text = re.sub(r"-\s*\n\s*", "", text)        # 行尾连字符断词
    text = re.sub(r"\s*\n\s*", " ", text)        # 段内换行合成空格
    text = re.sub(r"[ \t\u3000]+", " ", text)
    return text.strip()


def _looks_like_body(text: str) -> bool:
    """过滤页码、页眉页脚等噪声段落。"""
    if len(text) < _MIN_PARA_CHARS:
        return False
    if _PAGE_NOISE.match(text):
        return False
    # 纯数字/纯符号段落没有翻译价值
    if not re.search(r"[A-Za-z]{2}", text):
        return False
    return True


def _is_translatable(text: str) -> bool:
    """原位替换时是否应该翻译这一段。

    只翻译含成串英文的段落：已经是中文的内容、公式行、纯数字/单位保持原样，
    避免把公式或本就来是中文的内容改坏。
    """
    stripped = text.strip()
    if len(stripped) < 2:
        return False
    if not re.search(r"[A-Za-z]{2,}", stripped):
        return False

    letters = sum(1 for ch in stripped if ch.isascii() and ch.isalpha())
    digits = sum(1 for ch in stripped if ch.isdigit())
    math = sum(1 for ch in stripped if ch in _MATH_CHARS)
    total = len(stripped)

    if letters + digits == 0:
        return False
    # 公式行：数学符号多且字母占比低（如 "L softmax = − 1/n Σ ln p"）
    if math >= 3 and letters < total * 0.45:
        return False
    if math >= 2 and letters < total * 0.25:
        return False
    return True


def _order_lines(lines: list[dict], page_width: float) -> list[dict]:
    """按阅读顺序排列行，处理学术论文常见的双栏排版。

    横跨页面大部分宽度的行（标题、跨栏图表说明）作为分节边界；每个分节内部
    按「先左栏后右栏」排序。单栏页面自然退化为按纵坐标排序。
    """
    if not lines:
        return []

    wide_ratio = 0.62
    sections: list[list[dict]] = []
    current: list[dict] = []

    for line in sorted(lines, key=lambda x: x["bbox"][1]):
        width = line["bbox"][2] - line["bbox"][0]
        if page_width and width > page_width * wide_ratio:
            if current:
                sections.append(current)
                current = []
            sections.append([line])
        else:
            current.append(line)
    if current:
        sections.append(current)

    ordered: list[dict] = []
    mid = page_width / 2 if page_width else 0
    two_col = _is_two_column(lines, page_width)
    for section in sections:
        if len(section) == 1 or not two_col:
            ordered.extend(sorted(section, key=lambda x: x["bbox"][1]))
            continue
        # 双栏页面：整栏读完左栏再读右栏，否则左右栏会按 y 交错串行
        left = [x for x in section if (x["bbox"][0] + x["bbox"][2]) / 2 < mid]
        right = [x for x in section if (x["bbox"][0] + x["bbox"][2]) / 2 >= mid]
        ordered.extend(sorted(left, key=lambda x: x["bbox"][1]))
        ordered.extend(sorted(right, key=lambda x: x["bbox"][1]))
    return ordered


def _is_two_column(lines: list[dict], page_width: float) -> bool:
    """判断整页是否双栏。

    逐节判断会被跨栏标题切成的小节误导，回退成按 y 排序，使左右栏交错；
    这里改为整页统计：两半都够多行，且大多数行明显窄于页宽。
    """
    if not page_width or len(lines) < 6:
        return False
    mid = page_width / 2
    left = sum(1 for l in lines if (l["bbox"][0] + l["bbox"][2]) / 2 < mid)
    right = len(lines) - left
    if min(left, right) < 3:
        return False
    wide = sum(1 for l in lines if (l["bbox"][2] - l["bbox"][0]) > page_width * 0.62)
    return wide <= len(lines) * 0.4


def _page_lines(page) -> list[dict]:
    """把页面拆成视觉行。

    PyMuPDF 会把同一条视觉行拆成多个 line，甚至另起一个 block——上标引用（如 [12]）
    就会造成这种断裂。若直接按纵坐标排序，片段会被交错拼接，出现
    「具有较 [1]。…高的经济价值」这类错乱。这里按「纵向明显重叠 + 横向紧邻」把
    同一条视觉行的片段拼回，再按横坐标排列片段文字，恢复正确语序。
    """
    raw = []
    for block in page.get_text("dict").get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            text = "".join(span.get("text", "") for span in spans)
            if not text.strip() or not spans:
                continue
            # 取字符最多的 span 作为该行的主字号/颜色，用于原位重排时还原外观
            main = max(spans, key=lambda s: len(s.get("text", "")))
            raw.append({
                "bbox": list(line["bbox"]),
                "text": text,
                "size": float(main.get("size", 10.0)),
                "color": int(main.get("color", 0)),
            })
    if not raw:
        return []

    heights = [l["bbox"][3] - l["bbox"][1] for l in raw]
    char_w = max(2.0, (_median(heights) or 10.0) * 0.5)

    merged: list[dict] = []
    for line in sorted(raw, key=lambda l: (l["bbox"][1], l["bbox"][0])):
        target = None
        for candidate in merged:
            top = max(candidate["bbox"][1], line["bbox"][1])
            bottom = min(candidate["bbox"][3], line["bbox"][3])
            smaller = min(candidate["bbox"][3] - candidate["bbox"][1],
                          line["bbox"][3] - line["bbox"][1])
            if smaller <= 0 or (bottom - top) / smaller < 0.6:
                continue
            # 横向必须紧邻；跨栏或跨表格的间隔远大于一个字符宽
            if line["bbox"][0] > candidate["bbox"][2] + char_w * 1.2:
                continue
            if line["bbox"][2] < candidate["bbox"][0] - char_w * 1.2:
                continue
            target = candidate
            break

        if target is None:
            merged.append({"bbox": list(line["bbox"]), "fragments": [line]})
        else:
            target["fragments"].append(line)
            target["bbox"] = [
                min(target["bbox"][0], line["bbox"][0]),
                min(target["bbox"][1], line["bbox"][1]),
                max(target["bbox"][2], line["bbox"][2]),
                max(target["bbox"][3], line["bbox"][3]),
            ]

    result = []
    for item in merged:
        text = ""
        previous_x1 = None
        for fragment in sorted(item["fragments"], key=lambda f: f["bbox"][0]):
            if previous_x1 is not None:
                gap = fragment["bbox"][0] - previous_x1
                # 只在明显断开且两侧都不是中文时补空格，避免破坏中文排版
                if (gap > char_w * 0.6 and not text.endswith(" ")
                        and not fragment["text"].startswith(" ")
                        and not _CJK.search(text[-1:])
                        and not _CJK.search(fragment["text"][:1])):
                    text += " "
            text += fragment["text"]
            previous_x1 = fragment["bbox"][2]
        if text.strip():
            result.append({"bbox": item["bbox"], "text": text})
    return result


def _group_line_boxes(lines: list[dict]) -> list[dict]:
    """把行重组成自然段，并保留每段的位置信息。

    返回的每一项包含 text / bbox（整段外框）/ rects（各原始行外框）/ size / color。
    原位翻译需要 rects 来精确擦除原文、bbox 来放置译文，因此不能只返回字符串。

    PyMuPDF 的 block 并不等于自然段，中文文献里上标引用还会被切成独立行。
    这里按行间「中心距」判断换段：明显大于常规行距、或上一行已以句末标点结束
    且本行首行缩进，才另起一段；其余一律合并，避免把句子切碎影响翻译质量。
    """
    if not lines:
        return []

    heights = [l["bbox"][3] - l["bbox"][1] for l in lines]
    median_h = _median(heights) or 10.0

    centers = [(l["bbox"][1] + l["bbox"][3]) / 2 for l in lines]
    leads = [centers[i] - centers[i - 1] for i in range(1, len(centers))]
    median_lead = _median([d for d in leads if d > 0]) or median_h * 1.4
    gap_limit = max(median_lead * 1.35, median_h * 0.9)

    base_left = _median([l["bbox"][0] for l in lines])
    indent_limit = base_left + max(6.0, median_h * 0.8)

    groups: list[dict] = []
    buffer: list[dict] = []
    buffer_box: list[float] | None = None
    previous_center = None

    def extend_box(box, bbox) -> list[float]:
        if box is None:
            return [bbox[0], bbox[1], bbox[2], bbox[3]]
        return [min(box[0], bbox[0]), min(box[1], bbox[1]),
                max(box[2], bbox[2]), max(box[3], bbox[3])]

    def flush() -> None:
        nonlocal buffer_box
        if not buffer:
            return
        text = _clean_text(_join_lines([l["text"] for l in buffer]))
        if _looks_like_body(text):
            rects = [list(l["bbox"]) for l in buffer]
            groups.append({
                "text": text,
                "bbox": [
                    min(r[0] for r in rects), min(r[1] for r in rects),
                    max(r[2] for r in rects), max(r[3] for r in rects),
                ],
                "rects": rects,
                "size": _median([l.get("size", median_h) for l in buffer]) or median_h,
                "color": buffer[0].get("color", 0),
            })
        buffer.clear()
        buffer_box = None

    for line in lines:
        if not line["text"].strip():
            continue
        center = (line["bbox"][1] + line["bbox"][3]) / 2

        if buffer and previous_center is not None:
            delta = center - previous_center
            previous = buffer[-1]
            # 两个外框在纵横两个方向都重叠，说明它们属于同一块文字区域（大字号标题、
            # 作者行这类紧密排版很常见）。此时若还按行距/回跳规则断开，就会被拆成
            # 两段并互相覆盖，因此直接合并。左右两栏之间横向不重叠，不会误并。
            # 注意必须限定在"同一视觉行"内：否则阅读顺序稍有回跳，就会把相隔几百点、
            # 完全不同的两块（例如左栏关键词与右栏摘要）粘成一个横跨整页的巨型段落，
            # 之后的译文要么放不下、要么被判定压字，整块都翻译不出来。
            same_row = abs(center - previous_center) < median_h * 1.2
            same_region = (same_row
                           and line["bbox"][1] < previous["bbox"][3] - 0.5
                           and line["bbox"][0] < previous["bbox"][2]
                           and previous["bbox"][0] < line["bbox"][2])
            # delta 明显为负（或几乎为 0）说明阅读顺序回跳了——典型情况是左栏读完
            # 跳到右栏开头。必须在此另起一段，否则会把两栏粘成横跨整页的巨型段落。
            starts_new = delta > gap_limit or delta < median_lead * 0.2
            # 两行横向几乎不重叠（并行两栏、表格相邻单元格、图注与相邻图表）说明它们
            # 不是同一段文字，必须断开。否则会被粘成横跨整页的一段，译成中文后长度变化，
            # 就会压到旁边那一栏/单元格的文字上。用较窄的一行做分母，避免把"段末短行"
            # 误判成新段。
            overlap = (min(line["bbox"][2], previous["bbox"][2])
                       - max(line["bbox"][0], previous["bbox"][0]))
            narrower = max(1.0, min(line["bbox"][2] - line["bbox"][0],
                                    previous["bbox"][2] - previous["bbox"][0]))
            far_apart = _SPLIT_DISJOINT_LINES and (overlap / narrower) < 0.25
            if same_region:
                starts_new = False
            elif far_apart:
                starts_new = True
            elif (not starts_new and line["bbox"][0] > indent_limit
                    and _ENDS_SENTENCE.search(previous["text"])):
                starts_new = True
            # 与"整段已累积的外框"也要横向重叠，避免接力式串栏
            if buffer_box is not None:
                box_overlap = (min(line["bbox"][2], buffer_box[2])
                               - max(line["bbox"][0], buffer_box[0]))
                box_narrow = max(1.0, min(line["bbox"][2] - line["bbox"][0],
                                          buffer_box[2] - buffer_box[0]))
                if (box_overlap / box_narrow) < 0.25:
                    starts_new = True
            if starts_new:
                flush()

        buffer.append(line)
        buffer_box = extend_box(buffer_box, line["bbox"])
        previous_center = center

    flush()
    return groups


def _group_lines(lines: list[dict]) -> list[str]:
    """仅取段落文本（供 Word 输出使用）。"""
    return [group["text"] for group in _group_line_boxes(lines)]


def _pdf_paragraphs(path: str) -> list[str]:
    import fitz

    doc = fitz.open(path)
    try:
        paragraphs: list[str] = []
        for page in doc:
            ordered = _order_lines(_page_lines(page), page.rect.width)
            paragraphs.extend(_group_lines(ordered))
        return paragraphs
    finally:
        doc.close()


def _docx_paragraphs(path: str) -> list[str]:
    from docx import Document

    doc = Document(path)
    paragraphs = []
    for para in doc.paragraphs:
        text = _clean_text(para.text)
        if _looks_like_body(text):
            paragraphs.append(text)
    return paragraphs


def _text_paragraphs(path: str) -> list[str]:
    with open(path, "r", encoding="utf-8", errors="ignore") as handle:
        content = handle.read()
    # Markdown/TXT：空行分段
    chunks = re.split(r"\n\s*\n", content)
    paragraphs = []
    for chunk in chunks:
        text = _clean_text(chunk)
        if _looks_like_body(text):
            paragraphs.append(text)
    return paragraphs


def extract_paragraphs(src: str) -> list[str]:
    """按扩展名抽取文献正文段落。"""
    ext = os.path.splitext(src)[1].lower().lstrip(".")
    if ext == "pdf":
        paragraphs = _pdf_paragraphs(src)
    elif ext == "docx":
        paragraphs = _docx_paragraphs(src)
    elif ext in ("txt", "md", "markdown"):
        paragraphs = _text_paragraphs(src)
    else:
        raise ValueError(f"文献翻译暂不支持 .{ext} 格式，请上传 PDF / DOCX / TXT")
    if not paragraphs:
        raise ValueError("没有从文件中提取到可翻译的文字；扫描版 PDF 请先用「格式转换」转为可编辑 Word")
    return paragraphs


def drop_references(paragraphs: list[str]) -> list[str]:
    """截断参考文献部分：通常无需翻译，可显著减少请求量与费用。"""
    for index, para in enumerate(paragraphs):
        if _REF_HEADING.match(para.strip()) and index > 0:
            return paragraphs[:index]
    return paragraphs


# --------------------------------------------------------------------------
# LLM 调用
# --------------------------------------------------------------------------
class TranslationError(RuntimeError):
    pass


def _system_prompt(cfg: TranslationConfig) -> str:
    parts = [
        f"你是专业的学术文献翻译引擎，擅长把英文学术论文准确、通顺地翻译成{cfg.target_lang}。",
        "",
        "必须遵守：",
        "1. 保持学术语体与逻辑严谨，术语准确且全文前后一致。",
        f"2. 严格保持段落数量与顺序：输入由 {DELIM} 分隔为若干段，输出必须用同样数量的 {DELIM} 分隔，不得合并、拆分或增删段落。",
        "3. 只输出译文本身，不要输出原文、编号、解释、Markdown 代码块或任何额外说明。",
        "4. 公式、变量名、基因/物种拉丁名、化学式、单位、图表编号（如 Fig. 1、Table 2）与文献引用（如 [12]、(Smith et al., 2020)）保持原样不译。",
        "5. 若某段是标题或图表说明，同样翻译，并保持简洁。",
        "6. " + (
            "人名（作者署名、通讯作者、致谢中的人名）一律保留原文拉丁字母写法，不要按拼音"
            "猜汉字——猜错比不译更糟（例如「Jun-Neng Liang」照写即可，不要写成「梁俊能」）。"
            if not cfg.translate_names else
            "人名（作者署名、通讯作者、致谢中的人名）可以译为中文；但若不确定对应的汉字，"
            "就保留原文拉丁字母写法，不要写成可能错误的汉字。"
        ) + "只有术语表明确给出人名对照时才使用其中的中文写法；机构名、期刊名照常翻译。",
    ]
    if cfg.domain.strip():
        parts.append(f"7. 该文献所属领域：{cfg.domain.strip()}。请使用该领域的规范译名。")
    if cfg.glossary.strip():
        parts.append("")
        parts.append("术语对照（必须优先采用以下译法）：")
        parts.append(cfg.glossary.strip())
    return "\n".join(parts)


class _LlmClient:
    """最小化的 OpenAI 兼容客户端，带重试与连接复用。"""

    def __init__(self, cfg: TranslationConfig):
        self.cfg = cfg
        try:
            import httpx
        except ImportError as exc:  # pragma: no cover - 依赖缺失时给出明确指引
            raise TranslationError("缺少 httpx 依赖，请先执行 pip install httpx") from exc
        self._httpx = httpx
        self._client = httpx.Client(timeout=cfg.timeout)

    def close(self) -> None:
        try:
            self._client.close()
        except Exception:
            pass

    def chat(self, user_content: str) -> str:
        payload = {
            "model": self.cfg.model,
            "messages": [
                {"role": "system", "content": _system_prompt(self.cfg)},
                {"role": "user", "content": user_content},
            ],
            "temperature": self.cfg.temperature,
            "stream": False,
        }
        headers = {
            "Authorization": f"Bearer {self.cfg.api_key}",
            "Content-Type": "application/json",
        }
        last_error = ""
        for attempt in range(self.cfg.retries + 1):
            try:
                resp = self._client.post(self.cfg.endpoint(), json=payload, headers=headers)
            except Exception as exc:  # 网络层异常
                last_error = f"网络错误：{exc}"
                if attempt < self.cfg.retries:
                    time.sleep(min(2 ** attempt * 2, 20))
                    continue
                raise TranslationError(last_error) from exc

            if resp.status_code == 200:
                try:
                    data = resp.json()
                    return data["choices"][0]["message"]["content"] or ""
                except (KeyError, IndexError, ValueError) as exc:
                    raise TranslationError(f"接口返回格式异常：{resp.text[:300]}") from exc

            if resp.status_code in (401, 403):
                raise TranslationError("API 密钥无效或没有权限，请检查 TRANSLATE_API_KEY 配置")
            if resp.status_code == 402:
                raise TranslationError("API 账户余额不足，请先充值")
            if resp.status_code in (429, 500, 502, 503, 504) and attempt < self.cfg.retries:
                last_error = f"接口返回 {resp.status_code}"
                time.sleep(min(2 ** attempt * 2, 20))
                continue
            raise TranslationError(f"接口返回 {resp.status_code}：{resp.text[:300]}")

        raise TranslationError(last_error or "翻译接口调用失败")


def _split_for_limit(text: str, limit: int) -> list[str]:
    """把超长段落切成不超过 limit 的块，优先按句子边界。

    注意：没有句末标点的长文本（例如整段公式或连续的英文变量名）不会被句子切分，
    必须再按字符硬切，否则会超出模型上下文。
    """
    if len(text) <= limit:
        return [text]

    chunks, current = [], ""
    for sentence in _SENTENCE_SPLIT.split(text):
        while len(sentence) > limit:
            if current:
                chunks.append(current.strip())
                current = ""
            chunks.append(sentence[:limit])
            sentence = sentence[limit:]
        if current and len(current) + len(sentence) + 1 > limit:
            chunks.append(current.strip())
            current = sentence
        else:
            current = f"{current} {sentence}".strip()
    if current.strip():
        chunks.append(current.strip())
    return [c for c in chunks if c]


def _build_batches(paragraphs: list[str], limit: int) -> list[list[int]]:
    """按字符预算把段落编号分组，保证一批请求不会过长。"""
    batches: list[list[int]] = []
    current: list[int] = []
    size = 0
    for index, para in enumerate(paragraphs):
        para_len = len(para) + len(DELIM) + 2
        if current and size + para_len > limit:
            batches.append(current)
            current, size = [], 0
        current.append(index)
        size += para_len
    if current:
        batches.append(current)
    return batches


def _parse_batch(reply: str, expected: int) -> list[str] | None:
    """按分隔符解析批次译文；数量不符返回 None 触发降级。"""
    if DELIM in reply:
        pieces = [p.strip() for p in reply.split(DELIM)]
    else:
        # 模型偶尔用编号列表作答，退化尝试按行解析
        pieces = [p.strip() for p in re.split(r"\n{2,}", reply.strip()) if p.strip()]
    pieces = [p for p in pieces if p]
    if len(pieces) == expected:
        return pieces
    return None


def translate_paragraphs(paragraphs: list[str], cfg: TranslationConfig, progress=None) -> list[str]:
    """翻译全部段落，返回与输入等长的译文列表。

    progress: 可选回调 progress(done, total, message)
    """
    if not cfg.configured:
        raise TranslationError(
            "尚未配置翻译密钥。请在项目根目录创建 .env 并填写 TRANSLATE_API_KEY=sk-xxx"
        )

    total = len(paragraphs)
    results: list[str] = [""] * total
    client = _LlmClient(cfg)
    done = 0

    def report(message: str) -> None:
        if progress:
            progress(done, total, message)

    try:
        batches = _build_batches(paragraphs, max(500, cfg.max_chars))
        report(f"共 {total} 段，分 {len(batches)} 批翻译")

        for batch in batches:
            texts = [paragraphs[i] for i in batch]

            # 超长段落先按句子切分，逐块翻译后再合并回该段
            oversized = [i for i in batch if len(paragraphs[i]) > cfg.max_chars * 2]
            if oversized:
                for index in oversized:
                    pieces = _split_for_limit(paragraphs[index], cfg.max_chars)
                    merged = []
                    for piece in pieces:
                        merged.append(client.chat(piece).strip())
                        done += 1
                        report(f"正在翻译长段落 {done}/{total}")
                    results[index] = " ".join(m for m in merged if m)
                texts = [paragraphs[i] for i in batch if i not in oversized]
                if not texts:
                    continue
                batch = [i for i in batch if i not in oversized]

            reply = client.chat(f"\n{DELIM}\n".join(texts))
            pieces = _parse_batch(reply, len(batch))

            if pieces is None:
                # 降级：逐段翻译，牺牲速度换取段落严格对齐
                report(f"第 {len(batch)} 段批次格式异常，改为逐段翻译")
                for index in batch:
                    results[index] = client.chat(paragraphs[index]).strip()
                    done += 1
                    report(f"正在逐段翻译 {done}/{total}")
            else:
                for index, piece in zip(batch, pieces):
                    results[index] = piece
                done += len(batch)
                report(f"已完成 {done}/{total} 段")
    finally:
        client.close()

    # 理论上不会发生；真出现空译文时保留原文，便于人工定位
    return [text if text.strip() else paragraphs[i] for i, text in enumerate(results)]


# --------------------------------------------------------------------------
# Word 生成
# --------------------------------------------------------------------------
def _style_run(run, size: int, color: tuple[int, int, int], bold: bool = False) -> None:
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor

    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*color)
    # 中文字体需要单独设置 eastAsia，否则 Word 里会回退成默认字体
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")


def _add_para(doc, text: str, size: int, color: tuple[int, int, int], bold: bool = False):
    para = doc.add_paragraph()
    para.paragraph_format.space_after = None
    run = para.add_run(text)
    _style_run(run, size, color, bold)
    return para


def build_bilingual_docx(pairs: list[tuple[str, str]], dst: str, title: str = "") -> str:
    """逐段「原文 + 译文」对照的 Word。"""
    from docx import Document
    from docx.shared import Pt

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(10.5)

    if title:
        _add_para(doc, title, 14, (0x1F, 0x33, 0x64), bold=True)
    _add_para(doc, "格式：每段先英文原文，后中文译文。", 9, (0x80, 0x80, 0x80))

    for original, translated in pairs:
        para = doc.add_paragraph()
        para.paragraph_format.space_before = Pt(8)
        _style_run(para.add_run(original), 9.5, (0x70, 0x70, 0x70))
        _add_para(doc, translated, 10.5, (0x11, 0x11, 0x11))

    doc.save(dst)
    return dst


def build_translated_docx(pairs: list[tuple[str, str]], dst: str, title: str = "") -> str:
    """仅含译文的 Word。"""
    from docx import Document
    from docx.shared import Pt

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(10.5)

    if title:
        _add_para(doc, title, 14, (0x1F, 0x33, 0x64), bold=True)

    for _, translated in pairs:
        _add_para(doc, translated, 10.5, (0x11, 0x11, 0x11))

    doc.save(dst)
    return dst


def _boxes_conflict(a, b, ratio: float = 0.25) -> bool:
    """两个版块是否明显重叠（重叠面积超过较小的一块）。"""
    ix = min(a[2], b[2]) - max(a[0], b[0])
    iy = min(a[3], b[3]) - max(a[1], b[1])
    if ix <= 0 or iy <= 0:
        return False
    area_a = max(1e-6, (a[2] - a[0]) * (a[3] - a[1]))
    area_b = max(1e-6, (b[2] - b[0]) * (b[3] - b[1]))
    return (ix * iy) / min(area_a, area_b) > ratio


def _free_below(groups: list[dict], index: int, page_height: float, margin: float = 46.0) -> float:
    """本段下方到「同栏下一段」之间还有多少空白，可用于容纳变长的译文。"""
    group = groups[index]
    limit = page_height - margin
    for other in groups[index + 1:]:
        if other is group:
            continue
        # 只统计与当前段横向重叠（即同一栏）且位于其下方的块
        if other["bbox"][0] < group["bbox"][2] - 5 and other["bbox"][2] > group["bbox"][0] + 5:
            if other["bbox"][1] >= group["bbox"][3] - 1:
                limit = min(limit, other["bbox"][1])
    return max(0.0, limit - group["bbox"][3])


def _fit_size(body_html: str, width: float, height: float,
              start_size: float, min_size: float = _MIN_FIT_SIZE):
    """实测出能把译文真正放下的最大字号；返回 (字号, 剩余高度)；放不下返回 None。

    早先按字符宽度估算字号并不可靠：CSS 行高、行尾折行浪费都会让实际占用比估算高，
    遇到页脚这类又长又窄的块就会溢出压住下一段。这里在临时页上真实排版一次，
    用 insert_htmlbox 返回的 spare_height 判断是否放得下，从原始字号逐步降级。

    判定要求留有明确的余量（_FIT_SPARE_*）：只要"刚好放下"就仍然会显得密不透风，
    中文尤其明显，所以宁可再小一号，也不要顶满框。
    """
    import fitz

    if width <= 4 or height <= 4:
        return None
    doc = fitz.open()
    try:
        page = doc.new_page(width=width + 6, height=height + 6)
        box = fitz.Rect(0, 0, width, height)
        size = start_size
        while size >= min_size:
            try:
                spare, _ = page.insert_htmlbox(box, body_html, css=_translation_css(size),
                                               scale_low=1.0)
            except Exception:
                return None
            if spare >= max(_FIT_SPARE_MIN, size * _FIT_SPARE_RATIO):
                return size, spare
            size -= max(0.5, size * 0.10)
        return None
    finally:
        doc.close()


def _translation_css(size: float, color: int = 0x000000) -> str:
    """原位译文的统一样式；适配字号与实际落位必须用同一份，否则测算会失真。"""
    return ("* {font-family: 'cjk'; font-size: %gpx; line-height: %g; color: #%06x;}"
            % (max(4.0, size), _TRANSLATE_LINE_HEIGHT, color & 0xFFFFFF))


def _place_html(page, box, body_html: str, size: float, color: int) -> None:
    """用指定字号把译文写进 box；写不进去时必须报错（原文已被擦除，不能静默留白）。"""
    spare, _ = page.insert_htmlbox(box, body_html, css=_translation_css(size, color),
                                   scale_low=0.88)
    if spare < -0.5:
        raise RuntimeError("译文未能放入原区域（余量 %.1f）" % spare)


# --------------------------------------------------------------------------
# 对外入口
# --------------------------------------------------------------------------
def translate_pdf_layout(src: str, out_dir: str, cfg: TranslationConfig, progress=None) -> dict:
    """原版式翻译：擦除原文，在原坐标写入译文并保留图片、表格与公式。

    分两遍：第一遍只读取每页段落与坐标、收集待译文本，一次性送翻译；
    第二遍才逐页擦除这些段落的原始行并把译文写回同一区域。
    译文放不下时由 insert_htmlbox 自动缩小字号，而不是撑破版面。
    """
    import fitz

    if os.path.splitext(src)[1].lower() != ".pdf":
        raise TranslationError("「原版式 PDF」只支持 PDF 输入；其它格式请选择 Word 输出")

    doc = fitz.open(src)
    try:
        pages: list[list[dict]] = []
        for page in doc:
            ordered = _order_lines(_page_lines(page), page.rect.width)
            pages.append(_group_line_boxes(ordered))

        # 记录每个待译段落位于「第几页的第几段」，便于翻译后回填。
        # 版面重叠保护：万一分组仍不理想（例如复杂表头），同一区域只保留面积较大的
        # 一段，放弃与之重叠的段落——宁可保留原文，也不要把译文叠成一团。
        jobs: list[tuple[int, int]] = []
        texts: list[str] = []
        dropped = 0
        unfitted = 0
        failed = 0
        failure_notes: list[str] = []
        for pno, groups in enumerate(pages):
            candidates = [gno for gno, g in enumerate(groups)
                          if _is_translatable(g["text"])]
            accepted: list[int] = []
            for gno in sorted(candidates, key=lambda i: -(
                    (groups[i]["bbox"][2] - groups[i]["bbox"][0])
                    * (groups[i]["bbox"][3] - groups[i]["bbox"][1]))):
                if any(_boxes_conflict(groups[gno]["bbox"], groups[other]["bbox"])
                       for other in accepted):
                    dropped += 1
                    continue
                accepted.append(gno)
            for gno in sorted(accepted):
                jobs.append((pno, gno))
                texts.append(groups[gno]["text"])

        if not texts:
            raise TranslationError(
                "没有检测到可翻译的英文段落。扫描版 PDF 没有文字层，请先转为可编辑 Word"
            )

        skipped_refs = False
        if cfg.skip_references:
            kept = drop_references(texts)
            skipped_refs = len(kept) < len(texts)
            jobs, texts = jobs[:len(kept)], kept
        if not texts:
            raise TranslationError("跳过参考文献后没有剩余正文，请关闭「跳过参考文献」后重试")

        translations = translate_paragraphs(texts, cfg, progress=progress)

        per_page: dict[int, list[tuple[int, str]]] = {}
        for (pno, gno), translated in zip(jobs, translations):
            per_page.setdefault(pno, []).append((gno, translated))

        for pno, items in per_page.items():
            page = doc[pno]
            page_groups = pages[pno]

            # 先据实测算每段的落位与字号；放不下的直接放弃（保留原文），
            # 这样绝不会出现译文溢出压住下一段的情况。
            planned = []
            for gno, translated in items:
                group = page_groups[gno]
                # 译文当作文本插入前必须转义，否则 & < > 会被当作 HTML
                body_html = html_module.escape(translated).replace("\n", "<br>")
                box = fitz.Rect(group["bbox"])
                box.x0 -= 1.0
                box.y0 -= 1.0
                box.x1 += 1.5
                box.y1 += 2.0

                start = max(6.0, group["size"])
                fitted = _fit_size(body_html, box.width, box.height, start)
                if fitted is None:
                    # 原框放不下时借下方空白，而不是缩到看不清
                    room = _free_below(page_groups, gno, page.rect.height)
                    if room > 4:
                        grown = fitz.Rect(box)
                        grown.y1 += min(room, box.height * _GROW_FACTOR + 24)
                        grown_fit = _fit_size(body_html, grown.width, grown.height, start)
                        if grown_fit is not None:
                            box, fitted = grown, grown_fit
                if fitted is None:
                    unfitted += 1
                    continue
                size, spare = fitted
                # 译文实际绘制的范围（自框顶向下）：保护检查只看这块，
                # 避免"框里恰好嵌着别的英文"就把整段判为压字而放弃翻译。
                drawn_bottom = box.y0 + max(1.0, box.height - spare)
                planned.append({"gno": gno, "box": box, "size": size, "html": body_html,
                                "drawn": [box.x0, box.y0, box.x1, drawn_bottom]})

            # 落位后再查一次重叠（借下方空白可能让两段相撞），同样宁缺勿叠
            accepted: list[dict] = []
            for item in sorted(planned, key=lambda it: -it["box"].get_area()):
                if any(_boxes_conflict(list(item["box"]), list(other["box"]))
                       for other in accepted):
                    unfitted += 1
                    continue
                accepted.append(item)

            # 未翻译的内容（页眉页脚、公式行、表格里的数字等）会原样留在页面上，
            # 译文同样不能压到它们身上——这正是"堆在一起"的主要来源。
            translated_ids = {item["gno"] for item in accepted}
            protected: list[list[float]] = []
            for gno, other in enumerate(page_groups):
                if gno not in translated_ids:
                    protected.extend(other["rects"])
            # 图表区域同样要保护：向下借空白时不要把译文写到图上。
            # 整页级的背景图（面积过大）跳过，否则整页译文都会被判为冲突。
            try:
                page_area = max(1.0, page.rect.width * page.rect.height)
                for info in page.get_image_info():
                    bbox = info.get("bbox")
                    if not bbox:
                        continue
                    area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
                    if area < 400 or area > page_area * 0.6:
                        continue
                    protected.append([float(v) for v in bbox])
            except Exception:
                pass
            if protected:
                kept: list[dict] = []
                for item in accepted:
                    drawn = list(item["drawn"])
                    hits = [rect for rect in protected
                            if _boxes_conflict(drawn, rect, _PROTECTED_OVERLAP_RATIO)]
                    if hits:
                        # 译文会压到尚未翻译的内容（常见：段落框底部叠着图注、表格行）。
                        # 先尝试把框裁到冲突内容上方并缩小字号重排；实在放不下才放弃。
                        trimmed = fitz.Rect(item["box"])
                        trimmed.y1 = min(trimmed.y1, min(rect[1] for rect in hits) - 1.0)
                        refit = None
                        if trimmed.height > 8:
                            refit = _fit_size(item["html"], trimmed.width, trimmed.height,
                                              item["size"])
                        if refit is not None:
                            size, spare = refit
                            item["box"] = trimmed
                            item["size"] = size
                            item["drawn"] = [trimmed.x0, trimmed.y0, trimmed.x1,
                                             trimmed.y0 + max(1.0, trimmed.height - spare)]
                        else:
                            unfitted += 1
                            continue
                    kept.append(item)
                accepted = kept

            accepted.sort(key=lambda it: it["gno"])
            if not accepted:
                continue

            for item in accepted:
                for rect in page_groups[item["gno"]]["rects"]:
                    # fill=False：不画白底块，避免遮住带底色的区域
                    page.add_redact_annot(fitz.Rect(rect), fill=False)
            # 只移除文字，保留图片与矢量图形（图表、公式线条）
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                                  graphics=fitz.PDF_REDACT_LINE_ART_NONE)

            for item in accepted:
                try:
                    _place_html(page, item["box"], item["html"], item["size"],
                                page_groups[item["gno"]]["color"])
                except Exception as exc:
                    # 写入失败时原文已被擦除，必须记账而不是静默吞掉
                    failed += 1
                    failure_notes.append(f"p{pno + 1} 段{item['gno']}: {exc}")

        dst = os.path.join(
            out_dir, f"{os.path.splitext(os.path.basename(src))[0]}_原版式译文.pdf")
        doc.save(dst, garbage=4, deflate=True)
    finally:
        doc.close()

    return {
        "path": dst,
        "paragraphs": sum(len(g) for g in pages),
        "translated": len(jobs),
        "skipped_references": skipped_refs,
        "dropped": dropped,
        "unfitted": unfitted,
        "failed": failed,
        "failure_notes": failure_notes[:5],
    }


def translate_document(
    src: str,
    out_dir: str,
    cfg: TranslationConfig,
    output_mode: str = "bilingual",
    progress=None,
) -> dict:
    """翻译一篇文献并产出 Word / 原版式 PDF。

    output_mode:
      bilingual — 中英对照 Word
      chinese   — 仅中文译文 Word
      both      — 两者打包 zip
      layout_pdf — 原版式 PDF：擦除原文，在同一坐标写入中文，保留图表与公式
    返回 {"path": 输出文件, "paragraphs": 段数, "translated": 段数, "skipped_references": bool}
    """
    base = os.path.splitext(os.path.basename(src))[0]

    if output_mode == "layout_pdf":
        return translate_pdf_layout(src, out_dir, cfg, progress=progress)

    paragraphs = extract_paragraphs(src)
    extracted = len(paragraphs)

    skipped = False
    if cfg.skip_references:
        trimmed = drop_references(paragraphs)
        skipped = len(trimmed) < len(paragraphs)
        paragraphs = trimmed

    if not paragraphs:
        raise TranslationError("截断参考文献后没有剩余正文，请关闭「跳过参考文献」后重试")

    translations = translate_paragraphs(paragraphs, cfg, progress=progress)
    pairs = list(zip(paragraphs, translations))

    outputs = []
    if output_mode in ("bilingual", "both"):
        dst = os.path.join(out_dir, f"{base}_中英对照.docx")
        build_bilingual_docx(pairs, dst, title=base)
        outputs.append(dst)
    if output_mode in ("chinese", "both"):
        dst = os.path.join(out_dir, f"{base}_译文.docx")
        build_translated_docx(pairs, dst, title=base)
        outputs.append(dst)

    if not outputs:
        raise TranslationError(f"未知的输出形式：{output_mode}")

    if len(outputs) == 1:
        final = outputs[0]
    else:
        import converter

        final = os.path.join(out_dir, f"{base}_翻译结果.zip")
        converter.zip_files(outputs, final)

    return {
        "path": final,
        "paragraphs": extracted,
        "translated": len(pairs),
        "skipped_references": skipped,
    }

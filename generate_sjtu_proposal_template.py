from __future__ import annotations

import html
import shutil
import struct
import zipfile
from pathlib import Path


OUT = Path("上海交通大学_开题答辩PPT模板.pptx")
ASSET_DIR = Path("assets")
LOGO_SEAL = ASSET_DIR / "sjtu_official_seal.png"
LOGO_H = ASSET_DIR / "sjtu_official_horizontal_logo.png"

# PowerPoint EMU units
EMU_PER_IN = 914400
SLIDE_W = 13_333_333
SLIDE_H = 7_500_000

SJTU_RED = "B01F24"
SJTU_DARK = "2B2B2B"
SJTU_GRAY = "5F6368"
SJTU_LIGHT_BG = "F7F5F2"
SJTU_GOLD = "C8A45D"
WHITE = "FFFFFF"


def emu(inches: float) -> int:
    return int(inches * EMU_PER_IN)


def x(y: float) -> int:
    return emu(y)


def esc(value: str) -> str:
    return html.escape(value, quote=True)


def rgb(hex_color: str) -> str:
    return hex_color.strip("#").upper()


def read_png_size(path: Path) -> tuple[int, int]:
    data = path.read_bytes()
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError(f"{path} is not a PNG file")
    return struct.unpack(">II", data[16:24])


def rels_xml(rels: list[tuple[str, str, str, str | None]]) -> str:
    items = []
    for rid, typ, target, mode in rels:
        mode_attr = f' TargetMode="{mode}"' if mode else ""
        items.append(
            f'<Relationship Id="{rid}" Type="{typ}" Target="{target}"{mode_attr}/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(items)
        + "</Relationships>"
    )


def shape_id_gen():
    n = 2
    while True:
        yield n
        n += 1


def solid_fill(color: str, alpha: int | None = None) -> str:
    alpha_xml = f'<a:alpha val="{alpha}"/>' if alpha is not None else ""
    return f'<a:solidFill><a:srgbClr val="{rgb(color)}">{alpha_xml}</a:srgbClr></a:solidFill>'


def no_line() -> str:
    return "<a:ln><a:noFill/></a:ln>"


def rect_shape(
    sid: int,
    name: str,
    left: int,
    top: int,
    width: int,
    height: int,
    fill: str,
    line: str | None = None,
    radius: str = "rect",
) -> str:
    line = line if line is not None else no_line()
    return f"""
<p:sp>
  <p:nvSpPr><p:cNvPr id="{sid}" name="{esc(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{left}" y="{top}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm>
    <a:prstGeom prst="{radius}"><a:avLst/></a:prstGeom>
    {fill}
    {line}
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>"""


def text_run(
    text: str,
    size: int = 2400,
    color: str = SJTU_DARK,
    bold: bool = False,
    latin: str = "Aptos",
    ea: str = "Microsoft YaHei",
) -> str:
    b = ' b="1"' if bold else ""
    return (
        f'<a:r><a:rPr lang="zh-CN" sz="{size}"{b}>'
        f"{solid_fill(color)}"
        f'<a:latin typeface="{esc(latin)}"/><a:ea typeface="{esc(ea)}"/>'
        f'</a:rPr><a:t>{esc(text)}</a:t></a:r>'
    )


def paragraph(
    text: str = "",
    *,
    size: int = 2400,
    color: str = SJTU_DARK,
    bold: bool = False,
    align: str = "l",
    bullet: bool = False,
    level: int = 0,
    before: int = 0,
    after: int = 0,
) -> str:
    mar_l = 457200 + level * 274320 if bullet else 0
    indent = -228600 if bullet else 0
    bullet_xml = "<a:buChar char=\"•\"/>" if bullet else "<a:buNone/>"
    return (
        f'<a:p><a:pPr algn="{align}" marL="{mar_l}" indent="{indent}" '
        f'spcBef="{before}" spcAft="{after}">{bullet_xml}</a:pPr>'
        f'{text_run(text, size=size, color=color, bold=bold)}</a:p>'
    )


def text_box(
    sid: int,
    name: str,
    left: int,
    top: int,
    width: int,
    height: int,
    paragraphs_xml: str,
    *,
    fill: str = '<a:noFill/>',
    line: str | None = None,
    margin_l: int = 0,
    margin_r: int = 0,
    margin_t: int = 0,
    margin_b: int = 0,
    vertical_anchor: str = "top",
) -> str:
    line = line if line is not None else no_line()
    anchor = "ctr" if vertical_anchor == "middle" else "t"
    return f"""
<p:sp>
  <p:nvSpPr><p:cNvPr id="{sid}" name="{esc(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{left}" y="{top}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    {fill}
    {line}
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="{anchor}" lIns="{margin_l}" rIns="{margin_r}" tIns="{margin_t}" bIns="{margin_b}"/>
    <a:lstStyle/>
    {paragraphs_xml}
  </p:txBody>
</p:sp>"""


def picture(
    sid: int,
    name: str,
    rid: str,
    left: int,
    top: int,
    width: int,
    height: int,
    *,
    alpha_mod_fix: int | None = None,
) -> str:
    alpha_xml = ""
    if alpha_mod_fix is not None:
        alpha_xml = f'<a:alphaModFix amt="{alpha_mod_fix}"/>'
    return f"""
<p:pic>
  <p:nvPicPr><p:cNvPr id="{sid}" name="{esc(name)}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="{rid}">{alpha_xml}</a:blip>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="{left}" y="{top}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>"""


def line_shape(
    sid: int,
    name: str,
    x1: int,
    y1: int,
    x2: int,
    y2: int,
    color: str = SJTU_RED,
    width: int = 19050,
) -> str:
    left, top = min(x1, x2), min(y1, y2)
    cx, cy = abs(x2 - x1), abs(y2 - y1)
    return f"""
<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="{sid}" name="{esc(name)}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{left}" y="{top}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
    <a:ln w="{width}"><a:solidFill><a:srgbClr val="{rgb(color)}"/></a:solidFill></a:ln>
  </p:spPr>
</p:cxnSp>"""


class Slide:
    def __init__(self, title: str):
        self.title = title
        self.parts: list[str] = []
        self.rels: list[tuple[str, str, str, str | None]] = [
            (
                "rId1",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
                "../slideLayouts/slideLayout1.xml",
                None,
            )
        ]
        self._ids = shape_id_gen()
        self._rel_num = 2

    def sid(self) -> int:
        return next(self._ids)

    def add(self, xml: str) -> None:
        self.parts.append(xml)

    def add_picture(self, path: Path, name: str, left: int, top: int, width: int, height: int, *, alpha: int | None = None) -> None:
        rid = f"rId{self._rel_num}"
        self._rel_num += 1
        target = "../media/" + path.name
        self.rels.append(
            (
                rid,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                target,
                None,
            )
        )
        self.add(picture(self.sid(), name, rid, left, top, width, height, alpha_mod_fix=alpha))

    def common_background(self, *, with_header_logo: bool = True, section: str = "") -> None:
        self.add(rect_shape(self.sid(), "background", 0, 0, SLIDE_W, SLIDE_H, solid_fill(SJTU_LIGHT_BG)))
        self.add(rect_shape(self.sid(), "top red rule", 0, 0, SLIDE_W, emu(0.08), solid_fill(SJTU_RED)))
        self.add(rect_shape(self.sid(), "bottom rule", 0, SLIDE_H - emu(0.05), SLIDE_W, emu(0.05), solid_fill(SJTU_RED)))
        self.add_picture(LOGO_SEAL, "seal watermark", emu(9.7), emu(2.15), emu(3.2), emu(3.2), alpha=8500)
        if with_header_logo:
            self.add_picture(LOGO_H, "SJTU horizontal logo", emu(0.52), emu(0.25), emu(2.55), emu(0.84))
        if section:
            self.add(
                text_box(
                    self.sid(),
                    "section label",
                    emu(10.05),
                    emu(0.36),
                    emu(2.72),
                    emu(0.32),
                    paragraph(section, size=1100, color=SJTU_GRAY, align="r"),
                )
            )

    def add_title(self, text: str, subtitle: str | None = None) -> None:
        self.add(
            text_box(
                self.sid(),
                "slide title",
                emu(0.76),
                emu(0.98),
                emu(9.85),
                emu(0.72),
                paragraph(text, size=2500, color=SJTU_DARK, bold=True),
            )
        )
        self.add(rect_shape(self.sid(), "title accent", emu(0.76), emu(1.74), emu(0.92), emu(0.05), solid_fill(SJTU_RED)))
        if subtitle:
            self.add(
                text_box(
                    self.sid(),
                    "slide subtitle",
                    emu(1.85),
                    emu(1.63),
                    emu(8.5),
                    emu(0.28),
                    paragraph(subtitle, size=1050, color=SJTU_GRAY),
                )
            )

    def add_footer(self, page: int) -> None:
        self.add(
            text_box(
                self.sid(),
                "footer text",
                emu(0.74),
                emu(7.12),
                emu(6.0),
                emu(0.22),
                paragraph("上海交通大学 | 开题答辩", size=800, color=SJTU_GRAY),
            )
        )
        self.add(
            text_box(
                self.sid(),
                "page number",
                emu(12.2),
                emu(7.12),
                emu(0.42),
                emu(0.22),
                paragraph(f"{page:02d}", size=800, color=SJTU_GRAY, align="r"),
            )
        )

    def xml(self) -> str:
        body = "\n".join(self.parts)
        return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {body}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""

    def rels_xml(self) -> str:
        return rels_xml(self.rels)


def rounded_card(slide: Slide, left: float, top: float, width: float, height: float, title: str, body: list[str], accent: str = SJTU_RED) -> None:
    slide.add(
        rect_shape(
            slide.sid(),
            f"{title} card",
            emu(left),
            emu(top),
            emu(width),
            emu(height),
            solid_fill(WHITE),
            '<a:ln w="6350"><a:solidFill><a:srgbClr val="E6E1DA"/></a:solidFill></a:ln>',
            radius="roundRect",
        )
    )
    slide.add(rect_shape(slide.sid(), f"{title} accent", emu(left), emu(top), emu(0.08), emu(height), solid_fill(accent)))
    paras = paragraph(title, size=1450, color=SJTU_DARK, bold=True, after=80000)
    for item in body:
        paras += paragraph(item, size=1050, color=SJTU_GRAY, bullet=True, after=30000)
    slide.add(
        text_box(
            slide.sid(),
            f"{title} text",
            emu(left + 0.24),
            emu(top + 0.18),
            emu(width - 0.42),
            emu(height - 0.3),
            paras,
            margin_l=emu(0.02),
            margin_r=emu(0.02),
            margin_t=emu(0.02),
        )
    )


def build_slides() -> list[Slide]:
    slides: list[Slide] = []

    # 1 Cover
    s = Slide("封面")
    s.add(rect_shape(s.sid(), "background", 0, 0, SLIDE_W, SLIDE_H, solid_fill(SJTU_LIGHT_BG)))
    s.add(rect_shape(s.sid(), "left red panel", 0, 0, emu(1.05), SLIDE_H, solid_fill(SJTU_RED)))
    s.add(rect_shape(s.sid(), "gold pin", emu(0.92), emu(0.78), emu(0.13), emu(5.9), solid_fill(SJTU_GOLD)))
    s.add_picture(LOGO_H, "SJTU official horizontal logo", emu(1.45), emu(0.58), emu(3.2), emu(1.05))
    s.add_picture(LOGO_SEAL, "SJTU seal watermark", emu(8.55), emu(0.8), emu(3.95), emu(3.95), alpha=13000)
    s.add(
        text_box(
            s.sid(),
            "defense label",
            emu(1.48),
            emu(2.02),
            emu(2.1),
            emu(0.36),
            paragraph("开题答辩", size=1450, color=SJTU_RED, bold=True),
        )
    )
    s.add(
        text_box(
            s.sid(),
            "main title",
            emu(1.45),
            emu(2.55),
            emu(8.55),
            emu(1.3),
            paragraph("论文题目：在此填写研究课题名称", size=3000, color=SJTU_DARK, bold=True),
            margin_l=0,
            margin_r=0,
        )
    )
    s.add(rect_shape(s.sid(), "title rule", emu(1.48), emu(4.02), emu(2.05), emu(0.07), solid_fill(SJTU_RED)))
    meta = (
        paragraph("学院 / 系：在此填写", size=1250, color=SJTU_GRAY, after=50000)
        + paragraph("答 辩 人：姓名", size=1250, color=SJTU_GRAY, after=50000)
        + paragraph("指导教师：姓名 职称", size=1250, color=SJTU_GRAY, after=50000)
        + paragraph("日    期：2026年 月 日", size=1250, color=SJTU_GRAY)
    )
    s.add(text_box(s.sid(), "meta", emu(1.48), emu(4.55), emu(5.1), emu(1.5), meta))
    s.add(text_box(s.sid(), "source note", emu(1.48), emu(6.82), emu(9.6), emu(0.28), paragraph("标识素材来源：上海交通大学视觉形象识别系统（vi.sjtu.edu.cn）", size=850, color=SJTU_GRAY)))
    slides.append(s)

    # 2 Agenda
    s = Slide("目录")
    s.common_background(section="CONTENTS")
    s.add_title("汇报提纲", "建议控制在 15-20 分钟，可按学院要求增删")
    items = [
        ("01", "研究背景与问题提出", "研究对象、现实需求、理论意义"),
        ("02", "文献综述与研究现状", "关键脉络、争议焦点、研究空白"),
        ("03", "研究目标、内容与创新点", "核心问题、假设或命题、预期贡献"),
        ("04", "研究方法与技术路线", "数据、模型、实验或案例设计"),
        ("05", "可行性分析与进度安排", "基础条件、风险控制、阶段计划"),
    ]
    top = 2.18
    for i, (num, title, desc) in enumerate(items):
        y0 = top + i * 0.78
        s.add(rect_shape(s.sid(), f"agenda {num}", emu(0.98), emu(y0), emu(0.54), emu(0.44), solid_fill(SJTU_RED), radius="roundRect"))
        s.add(text_box(s.sid(), f"agenda num {num}", emu(0.98), emu(y0 + 0.07), emu(0.54), emu(0.18), paragraph(num, size=850, color=WHITE, bold=True, align="ctr")))
        s.add(text_box(s.sid(), f"agenda title {num}", emu(1.78), emu(y0 - 0.01), emu(3.0), emu(0.28), paragraph(title, size=1350, color=SJTU_DARK, bold=True)))
        s.add(text_box(s.sid(), f"agenda desc {num}", emu(5.05), emu(y0 + 0.03), emu(4.4), emu(0.26), paragraph(desc, size=1000, color=SJTU_GRAY)))
        s.add(line_shape(s.sid(), f"agenda line {num}", emu(1.78), emu(y0 + 0.56), emu(9.55), emu(y0 + 0.56), color="DED8D0", width=6350))
    s.add_footer(2)
    slides.append(s)

    # 3 Background
    s = Slide("研究背景")
    s.common_background(section="01 / BACKGROUND")
    s.add_title("研究背景与问题提出", "把“大背景”收束为一个清晰、可研究的问题")
    rounded_card(s, 0.92, 2.05, 3.55, 3.7, "现实背景", ["行业/领域正在发生的关键变化", "已有实践中的痛点或未满足需求", "研究问题的重要性与紧迫性"], SJTU_RED)
    rounded_card(s, 4.9, 2.05, 3.55, 3.7, "理论背景", ["相关理论或学科脉络", "已有研究的解释边界", "需要进一步澄清的概念或机制"], SJTU_GOLD)
    rounded_card(s, 8.88, 2.05, 3.25, 3.7, "问题提出", ["本文聚焦的核心问题", "研究对象与范围限定", "预期回答的关键问题"], "4D5B50")
    s.add_footer(3)
    slides.append(s)

    # 4 Literature
    s = Slide("文献综述")
    s.common_background(section="02 / LITERATURE")
    s.add_title("文献综述与研究现状", "建议按主题、方法或争议线索组织，而不只按作者堆叠")
    # timeline
    s.add(line_shape(s.sid(), "timeline", emu(1.25), emu(3.42), emu(11.45), emu(3.42), color="CFC7BA", width=12700))
    blocks = [
        ("主题一", "代表文献与主要结论\n可填 2-3 条要点"),
        ("主题二", "方法或数据演进\n说明适用边界"),
        ("主题三", "尚未解决的问题\n形成研究空白"),
    ]
    for i, (t, b) in enumerate(blocks):
        lx = 1.55 + i * 3.65
        s.add(rect_shape(s.sid(), f"literature dot {i}", emu(lx), emu(3.24), emu(0.36), emu(0.36), solid_fill(SJTU_RED), radius="ellipse"))
        s.add(text_box(s.sid(), f"literature title {i}", emu(lx - 0.35), emu(2.15), emu(2.05), emu(0.36), paragraph(t, size=1500, color=SJTU_DARK, bold=True, align="ctr")))
        body_xml = "".join(paragraph(line, size=1000, color=SJTU_GRAY, align="ctr", after=25000) for line in b.split("\n"))
        s.add(text_box(s.sid(), f"literature body {i}", emu(lx - 0.75), emu(3.85), emu(2.75), emu(0.95), body_xml))
    s.add(text_box(s.sid(), "gap label", emu(1.1), emu(5.78), emu(10.7), emu(0.55), paragraph("研究空白：请用一句话概括“已有研究尚未充分解释/验证/解决什么”。", size=1400, color=SJTU_RED, bold=True, align="ctr")))
    s.add_footer(4)
    slides.append(s)

    # 5 Objectives
    s = Slide("研究目标")
    s.common_background(section="03 / OBJECTIVES")
    s.add_title("研究目标、内容与创新点", "把目标写成可检验、可完成的表达")
    rounded_card(s, 0.92, 2.0, 3.4, 3.88, "研究目标", ["目标 1：描述要解释或解决的问题", "目标 2：描述拟建立的模型/框架", "目标 3：描述预期验证或应用结果"], SJTU_RED)
    rounded_card(s, 4.72, 2.0, 3.4, 3.88, "研究内容", ["内容 1：概念界定与理论分析", "内容 2：数据收集、实验或案例研究", "内容 3：结果分析与机制讨论"], SJTU_GOLD)
    rounded_card(s, 8.52, 2.0, 3.4, 3.88, "创新点", ["视角创新：新的研究切入点", "方法创新：新的模型/数据/识别策略", "应用创新：可落地的解释或工具"], "2F6F73")
    s.add_footer(5)
    slides.append(s)

    # 6 Technical route
    s = Slide("技术路线")
    s.common_background(section="04 / METHOD")
    s.add_title("研究方法与技术路线", "用流程图说明“从问题到结论”的路径")
    steps = [
        ("问题界定", "研究问题\n变量/概念"),
        ("数据与材料", "数据来源\n样本范围"),
        ("方法设计", "模型/实验\n案例框架"),
        ("分析验证", "稳健性\n对比分析"),
        ("结论产出", "理论贡献\n实践建议"),
    ]
    y0 = 3.0
    for i, (t, b) in enumerate(steps):
        lx = 0.78 + i * 2.48
        fill = SJTU_RED if i in (0, 4) else WHITE
        color = WHITE if i in (0, 4) else SJTU_DARK
        line = no_line() if i in (0, 4) else '<a:ln w="9525"><a:solidFill><a:srgbClr val="D9D2C6"/></a:solidFill></a:ln>'
        s.add(rect_shape(s.sid(), f"route step {i}", emu(lx), emu(y0), emu(1.82), emu(1.08), solid_fill(fill), line, radius="roundRect"))
        paras = paragraph(t, size=1150, color=color, bold=True, align="ctr", after=20000)
        for line_text in b.split("\n"):
            paras += paragraph(line_text, size=850, color=color if i in (0, 4) else SJTU_GRAY, align="ctr")
        s.add(text_box(s.sid(), f"route text {i}", emu(lx + 0.1), emu(y0 + 0.14), emu(1.62), emu(0.72), paras, vertical_anchor="middle"))
        if i < len(steps) - 1:
            s.add(line_shape(s.sid(), f"route arrow line {i}", emu(lx + 1.86), emu(y0 + 0.54), emu(lx + 2.35), emu(y0 + 0.54), color=SJTU_GOLD, width=19050))
            s.add(rect_shape(s.sid(), f"route arrow head {i}", emu(lx + 2.27), emu(y0 + 0.45), emu(0.16), emu(0.16), solid_fill(SJTU_GOLD), radius="triangle"))
    s.add(text_box(s.sid(), "method note", emu(1.08), emu(5.25), emu(10.95), emu(0.5), paragraph("可在下方补充：核心模型、实验设计、样本选择、变量定义、评价指标等。", size=1150, color=SJTU_GRAY, align="ctr")))
    s.add_footer(6)
    slides.append(s)

    # 7 Feasibility
    s = Slide("可行性")
    s.common_background(section="05 / FEASIBILITY")
    s.add_title("可行性分析与风险控制", "让评委看到“能做完、做得成、风险可控”")
    rows = [
        ("研究基础", "已有课程、论文、前期调研、数据积累"),
        ("资料/数据", "数据来源、权限、样本规模、获取周期"),
        ("方法条件", "软件工具、实验平台、计算资源、导师支持"),
        ("风险控制", "数据不足、模型失效、进度延误的替代方案"),
    ]
    s.add(rect_shape(s.sid(), "table bg", emu(0.98), emu(2.08), emu(10.9), emu(3.75), solid_fill(WHITE), '<a:ln w="6350"><a:solidFill><a:srgbClr val="E6E1DA"/></a:solidFill></a:ln>', radius="roundRect"))
    for i, (left_text, right_text) in enumerate(rows):
        yy = 2.18 + i * 0.86
        s.add(rect_shape(s.sid(), f"table tag {i}", emu(1.18), emu(yy), emu(1.62), emu(0.46), solid_fill(SJTU_RED if i == 0 else "EFE8DF"), no_line(), radius="roundRect"))
        s.add(text_box(s.sid(), f"table left {i}", emu(1.28), emu(yy + 0.1), emu(1.42), emu(0.18), paragraph(left_text, size=900, color=WHITE if i == 0 else SJTU_DARK, bold=True, align="ctr")))
        s.add(text_box(s.sid(), f"table right {i}", emu(3.18), emu(yy + 0.03), emu(7.75), emu(0.28), paragraph(right_text, size=1120, color=SJTU_GRAY)))
        if i < len(rows) - 1:
            s.add(line_shape(s.sid(), f"row line {i}", emu(1.18), emu(yy + 0.68), emu(11.58), emu(yy + 0.68), color="E9E3DA", width=6350))
    s.add_footer(7)
    slides.append(s)

    # 8 Schedule
    s = Slide("进度安排")
    s.common_background(section="06 / SCHEDULE")
    s.add_title("研究计划与进度安排", "可按学期、月份或学院节点调整")
    months = ["第1阶段", "第2阶段", "第3阶段", "第4阶段", "第5阶段"]
    tasks = [
        ("文献整理", 0, 1),
        ("研究设计", 1, 2),
        ("数据/材料收集", 2, 3),
        ("分析与验证", 3, 4),
        ("论文撰写与修改", 4, 5),
    ]
    left = 2.65
    top = 2.2
    cell_w = 1.65
    for i, m in enumerate(months):
        s.add(text_box(s.sid(), f"month {i}", emu(left + i * cell_w), emu(top), emu(cell_w), emu(0.28), paragraph(m, size=900, color=SJTU_DARK, bold=True, align="ctr")))
        s.add(line_shape(s.sid(), f"month grid {i}", emu(left + i * cell_w), emu(top + 0.44), emu(left + i * cell_w), emu(5.9), color="E1DBD3", width=6350))
    s.add(line_shape(s.sid(), "month grid last", emu(left + len(months) * cell_w), emu(top + 0.44), emu(left + len(months) * cell_w), emu(5.9), color="E1DBD3", width=6350))
    for i, (task, start, end) in enumerate(tasks):
        yy = top + 0.66 + i * 0.62
        s.add(text_box(s.sid(), f"task {i}", emu(0.95), emu(yy + 0.05), emu(1.42), emu(0.22), paragraph(task, size=900, color=SJTU_GRAY, align="r")))
        s.add(rect_shape(s.sid(), f"bar {i}", emu(left + start * cell_w + 0.1), emu(yy), emu((end - start) * cell_w - 0.2), emu(0.28), solid_fill(SJTU_RED if i in (0, 4) else SJTU_GOLD), no_line(), radius="roundRect"))
        s.add(line_shape(s.sid(), f"task grid {i}", emu(left), emu(yy + 0.44), emu(left + len(months) * cell_w), emu(yy + 0.44), color="EDE7DF", width=6350))
    s.add_footer(8)
    slides.append(s)

    # 9 Expected outcomes
    s = Slide("预期成果")
    s.common_background(section="07 / OUTCOMES")
    s.add_title("预期成果与可能贡献", "把成果类型和贡献边界说清楚")
    rounded_card(s, 1.05, 2.1, 3.25, 3.7, "论文成果", ["完成开题报告与学位论文", "形成清晰的问题、方法与结论链条", "按学院规范完成查重与格式要求"], SJTU_RED)
    rounded_card(s, 5.0, 2.1, 3.25, 3.7, "学术贡献", ["补充已有理论解释", "提供新的数据、模型或经验证据", "为后续研究留下可复用框架"], SJTU_GOLD)
    rounded_card(s, 8.95, 2.1, 3.25, 3.7, "实践价值", ["对行业/组织/政策形成建议", "提炼可操作方案或评估指标", "明确适用场景与局限"], "4D5B50")
    s.add_footer(9)
    slides.append(s)

    # 10 Thanks
    s = Slide("致谢")
    s.add(rect_shape(s.sid(), "background", 0, 0, SLIDE_W, SLIDE_H, solid_fill(SJTU_RED)))
    s.add_picture(LOGO_H, "SJTU official horizontal logo", emu(0.7), emu(0.55), emu(3.0), emu(0.98))
    s.add_picture(LOGO_SEAL, "SJTU seal watermark", emu(8.55), emu(1.1), emu(3.65), emu(3.65), alpha=16000)
    s.add(text_box(s.sid(), "thanks", emu(1.05), emu(2.6), emu(7.0), emu(0.84), paragraph("敬请各位老师批评指正", size=3100, color=WHITE, bold=True)))
    s.add(rect_shape(s.sid(), "thanks rule", emu(1.08), emu(3.75), emu(2.1), emu(0.06), solid_fill(SJTU_GOLD)))
    s.add(text_box(s.sid(), "contact", emu(1.08), emu(4.28), emu(6.8), emu(0.72), paragraph("答辩人：姓名    邮箱：name@sjtu.edu.cn", size=1250, color=WHITE)))
    s.add(text_box(s.sid(), "source note", emu(1.08), emu(6.82), emu(10.8), emu(0.24), paragraph("本模板使用的校徽及中英文横版标志来自上海交通大学视觉形象识别系统官方附件。", size=850, color=WHITE)))
    slides.append(s)

    # 11 Source and usage note
    s = Slide("标识来源说明")
    s.common_background(section="APPENDIX")
    s.add_title("标识来源与使用说明", "保留此页便于答辩材料归档，也可按需要删除")
    notes = [
        "校徽图片：上海交通大学视觉形象识别系统官方附件“校标-校徽.png”。",
        "横版标志：上海交通大学视觉形象识别系统官方附件“校标-标志中英文横版.png”。",
        "模板主色：依据官方 VI 管理办法中标准色 CMYK（C20 M100 Y100 K0）近似换算为屏幕红色使用。",
        "建议不要拉伸、重绘、改色或拼接校徽/校名标志；如需更多格式，请从官方 VI 页面重新下载。",
    ]
    body = ""
    for item in notes:
        body += paragraph(item, size=1250, color=SJTU_GRAY, bullet=True, after=55000)
    s.add(text_box(s.sid(), "usage notes", emu(1.05), emu(2.1), emu(10.9), emu(2.5), body))
    s.add(text_box(s.sid(), "url", emu(1.05), emu(5.28), emu(10.9), emu(0.65), paragraph("官方页面：https://vi.sjtu.edu.cn/index.php/articles/bulletin/13", size=1200, color=SJTU_RED, bold=True)))
    s.add_footer(11)
    slides.append(s)

    return slides


def content_types_xml(slide_count: int) -> str:
    slide_overrides = "".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, slide_count + 1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
  <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
  <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  {slide_overrides}
</Types>"""


def presentation_xml(slide_count: int) -> str:
    sld_ids = "".join(
        f'<p:sldId id="{255 + i}" r:id="rId{i}"/>' for i in range(1, slide_count + 1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                saveSubsetFonts="1">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId{slide_count + 1}"/></p:sldMasterIdLst>
  <p:sldIdLst>{sld_ids}</p:sldIdLst>
  <p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}" type="wide"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle>
    <a:defPPr><a:defRPr lang="zh-CN"><a:latin typeface="Aptos"/><a:ea typeface="Microsoft YaHei"/></a:defRPr></a:defPPr>
  </p:defaultTextStyle>
</p:presentation>"""


def presentation_rels_xml(slide_count: int) -> str:
    rels = []
    for i in range(1, slide_count + 1):
        rels.append(
            (
                f"rId{i}",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                f"slides/slide{i}.xml",
                None,
            )
        )
    rels.extend(
        [
            (
                f"rId{slide_count + 1}",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
                "slideMasters/slideMaster1.xml",
                None,
            ),
            (
                f"rId{slide_count + 2}",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                "theme/theme1.xml",
                None,
            ),
            (
                f"rId{slide_count + 3}",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps",
                "presProps.xml",
                None,
            ),
            (
                f"rId{slide_count + 4}",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps",
                "viewProps.xml",
                None,
            ),
            (
                f"rId{slide_count + 5}",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles",
                "tableStyles.xml",
                None,
            ),
        ]
    )
    return rels_xml(rels)


ROOT_RELS = rels_xml(
    [
        (
            "rId1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "ppt/presentation.xml",
            None,
        ),
        (
            "rId2",
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
            "docProps/core.xml",
            None,
        ),
        (
            "rId3",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
            "docProps/app.xml",
            None,
        ),
    ]
)


CORE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>上海交通大学开题答辩PPT模板</dc:title>
  <dc:subject>开题答辩模板</dc:subject>
  <dc:creator>Codex</dc:creator>
  <cp:keywords>上海交通大学; SJTU; 开题答辩; PPT模板</cp:keywords>
  <dc:description>使用上海交通大学视觉形象识别系统官方附件生成的开题答辩模板。</dc:description>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-07-08T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-08T00:00:00Z</dcterms:modified>
</cp:coreProperties>"""


def app_xml(slide_count: int) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft PowerPoint</Application>
  <PresentationFormat>On-screen Show (16:9)</PresentationFormat>
  <Slides>{slide_count}</Slides>
  <Company>上海交通大学</Company>
  <AppVersion>16.0000</AppVersion>
</Properties>"""


THEME_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="SJTU Proposal Theme">
  <a:themeElements>
    <a:clrScheme name="SJTU">
      <a:dk1><a:srgbClr val="{SJTU_DARK}"/></a:dk1>
      <a:lt1><a:srgbClr val="{WHITE}"/></a:lt1>
      <a:dk2><a:srgbClr val="{SJTU_RED}"/></a:dk2>
      <a:lt2><a:srgbClr val="{SJTU_LIGHT_BG}"/></a:lt2>
      <a:accent1><a:srgbClr val="{SJTU_RED}"/></a:accent1>
      <a:accent2><a:srgbClr val="{SJTU_GOLD}"/></a:accent2>
      <a:accent3><a:srgbClr val="2F6F73"/></a:accent3>
      <a:accent4><a:srgbClr val="4D5B50"/></a:accent4>
      <a:accent5><a:srgbClr val="8A6D3B"/></a:accent5>
      <a:accent6><a:srgbClr val="{SJTU_GRAY}"/></a:accent6>
      <a:hlink><a:srgbClr val="{SJTU_RED}"/></a:hlink>
      <a:folHlink><a:srgbClr val="7A1B1D"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="SJTU Fonts">
      <a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface="Microsoft YaHei"/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/><a:ea typeface="Microsoft YaHei"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="SJTU Format">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>"""


SLIDE_MASTER_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>"""

SLIDE_MASTER_RELS = rels_xml(
    [
        (
            "rId1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
            "../slideLayouts/slideLayout1.xml",
            None,
        ),
        (
            "rId2",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
            "../theme/theme1.xml",
            None,
        ),
    ]
)

SLIDE_LAYOUT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"""

SLIDE_LAYOUT_RELS = rels_xml(
    [
        (
            "rId1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
            "../slideMasters/slideMaster1.xml",
            None,
        )
    ]
)

PRES_PROPS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"""

VIEW_PROPS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>
  <p:slideViewPr><p:cSldViewPr><p:cViewPr varScale="1"><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr>
</p:viewPr>"""

TABLE_STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>"""


def write_pptx(slides: list[Slide]) -> None:
    if not LOGO_SEAL.exists() or not LOGO_H.exists():
        raise FileNotFoundError("Missing official SJTU logo assets in assets/")

    # Validate PNGs before embedding.
    read_png_size(LOGO_SEAL)
    read_png_size(LOGO_H)

    if OUT.exists():
        OUT.unlink()

    with zipfile.ZipFile(OUT, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types_xml(len(slides)))
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("docProps/core.xml", CORE_XML)
        z.writestr("docProps/app.xml", app_xml(len(slides)))
        z.writestr("ppt/presentation.xml", presentation_xml(len(slides)))
        z.writestr("ppt/_rels/presentation.xml.rels", presentation_rels_xml(len(slides)))
        z.writestr("ppt/theme/theme1.xml", THEME_XML)
        z.writestr("ppt/slideMasters/slideMaster1.xml", SLIDE_MASTER_XML)
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", SLIDE_MASTER_RELS)
        z.writestr("ppt/slideLayouts/slideLayout1.xml", SLIDE_LAYOUT_XML)
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", SLIDE_LAYOUT_RELS)
        z.writestr("ppt/presProps.xml", PRES_PROPS_XML)
        z.writestr("ppt/viewProps.xml", VIEW_PROPS_XML)
        z.writestr("ppt/tableStyles.xml", TABLE_STYLES_XML)
        z.write(LOGO_SEAL, "ppt/media/" + LOGO_SEAL.name)
        z.write(LOGO_H, "ppt/media/" + LOGO_H.name)
        for i, slide in enumerate(slides, 1):
            z.writestr(f"ppt/slides/slide{i}.xml", slide.xml())
            z.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", slide.rels_xml())


def main() -> None:
    slides = build_slides()
    write_pptx(slides)
    print(f"Created {OUT.resolve()} ({OUT.stat().st_size:,} bytes)")
    print(f"Embedded official assets: {LOGO_SEAL.name}, {LOGO_H.name}")


if __name__ == "__main__":
    main()

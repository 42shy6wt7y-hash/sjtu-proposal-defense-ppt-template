from __future__ import annotations

from pathlib import Path

import generate_sjtu_proposal_template as base


OUT = Path("上海交通大学_开题答辩PPT模板_深蓝版.pptx")

NAVY = "08111F"
NAVY_2 = "0B1B30"
NAVY_3 = "10263F"
INK = "EAF2FF"
MUTED = "A9B8CC"
SOFT = "DDEBFF"
BLUE = "4EA1FF"
CYAN = "77E2FF"
ICE = "F7FBFF"
GOLD = "C8A45D"
RED = "B01F24"
PANEL = "122A45"
PANEL_2 = "16375A"
LINE = "34516F"


def e(value: float) -> int:
    return base.emu(value)


def p(
    text: str,
    *,
    size: int = 1200,
    color: str = INK,
    bold: bool = False,
    align: str = "l",
    bullet: bool = False,
    after: int = 0,
) -> str:
    return base.paragraph(
        text,
        size=size,
        color=color,
        bold=bold,
        align=align,
        bullet=bullet,
        after=after,
    )


def add_bg(slide: base.Slide, section: str = "", *, logo: bool = True) -> None:
    slide.add(base.rect_shape(slide.sid(), "deep blue background", 0, 0, base.SLIDE_W, base.SLIDE_H, base.solid_fill(NAVY)))
    slide.add(base.rect_shape(slide.sid(), "soft top wash", 0, 0, base.SLIDE_W, e(1.25), base.solid_fill(NAVY_2)))
    slide.add(base.rect_shape(slide.sid(), "bottom depth", 0, e(6.25), base.SLIDE_W, e(1.25), base.solid_fill("050B14")))
    slide.add(base.rect_shape(slide.sid(), "left hairline", e(0.54), e(0.55), e(0.015), e(6.15), base.solid_fill(BLUE, alpha=65000)))
    slide.add(base.rect_shape(slide.sid(), "top hairline", e(0.54), e(0.55), e(11.95), e(0.012), base.solid_fill(CYAN, alpha=55000)))
    slide.add_picture(base.LOGO_SEAL, "SJTU seal watermark", e(8.72), e(1.05), e(3.85), e(3.85), alpha=9000)
    if logo:
        slide.add(
            base.rect_shape(
                slide.sid(),
                "official logo capsule",
                e(0.78),
                e(0.72),
                e(2.55),
                e(0.72),
                base.solid_fill(ICE, alpha=93000),
                '<a:ln w="6350"><a:solidFill><a:srgbClr val="D9E8F8"><a:alpha val="65000"/></a:srgbClr></a:solidFill></a:ln>',
                radius="roundRect",
            )
        )
        slide.add_picture(base.LOGO_H, "SJTU official horizontal logo", e(0.96), e(0.85), e(2.1), e(0.47))
    if section:
        slide.add(
            base.text_box(
                slide.sid(),
                "section label",
                e(9.6),
                e(0.82),
                e(2.9),
                e(0.28),
                p(section, size=850, color=MUTED, align="r"),
            )
        )


def title(slide: base.Slide, text: str, kicker: str = "") -> None:
    if kicker:
        slide.add(base.text_box(slide.sid(), "kicker", e(0.84), e(1.62), e(5.0), e(0.32), p(kicker, size=900, color=CYAN, bold=True)))
    slide.add(base.text_box(slide.sid(), "slide title", e(0.82), e(1.9), e(8.6), e(0.72), p(text, size=2700, color=INK, bold=True)))
    slide.add(base.rect_shape(slide.sid(), "title rule", e(0.84), e(2.76), e(1.15), e(0.035), base.solid_fill(CYAN)))


def footer(slide: base.Slide, page: int) -> None:
    slide.add(base.text_box(slide.sid(), "footer", e(0.84), e(7.05), e(4.8), e(0.22), p("上海交通大学 | 开题答辩深蓝模板", size=720, color="7891AD")))
    slide.add(base.text_box(slide.sid(), "page", e(12.0), e(7.04), e(0.55), e(0.24), p(f"{page:02d}", size=780, color="7891AD", align="r")))


def glass_card(
    slide: base.Slide,
    left: float,
    top: float,
    width: float,
    height: float,
    heading: str,
    body: list[str],
    *,
    accent: str = BLUE,
) -> None:
    slide.add(
        base.rect_shape(
            slide.sid(),
            f"{heading} glass",
            e(left),
            e(top),
            e(width),
            e(height),
            base.solid_fill(PANEL, alpha=82000),
            '<a:ln w="6350"><a:solidFill><a:srgbClr val="40627F"><a:alpha val="68000"/></a:srgbClr></a:solidFill></a:ln>',
            radius="roundRect",
        )
    )
    slide.add(base.rect_shape(slide.sid(), f"{heading} accent", e(left + 0.22), e(top + 0.22), e(0.36), e(0.045), base.solid_fill(accent)))
    text = p(heading, size=1280, color=INK, bold=True, after=65000)
    for item in body:
        text += p(item, size=920, color=MUTED, bullet=True, after=26000)
    slide.add(base.text_box(slide.sid(), f"{heading} text", e(left + 0.28), e(top + 0.36), e(width - 0.5), e(height - 0.58), text))


def stat_card(slide: base.Slide, left: float, top: float, num: str, label: str, note: str, accent: str = CYAN) -> None:
    slide.add(
        base.rect_shape(
            slide.sid(),
            f"stat {num}",
            e(left),
            e(top),
            e(2.52),
            e(1.32),
            base.solid_fill(PANEL_2, alpha=76000),
            '<a:ln w="6350"><a:solidFill><a:srgbClr val="3B6388"><a:alpha val="62000"/></a:srgbClr></a:solidFill></a:ln>',
            radius="roundRect",
        )
    )
    slide.add(base.text_box(slide.sid(), f"stat num {num}", e(left + 0.22), e(top + 0.2), e(1.0), e(0.44), p(num, size=2300, color=accent, bold=True)))
    slide.add(base.text_box(slide.sid(), f"stat label {num}", e(left + 1.12), e(top + 0.23), e(1.15), e(0.24), p(label, size=950, color=INK, bold=True)))
    slide.add(base.text_box(slide.sid(), f"stat note {num}", e(left + 0.24), e(top + 0.82), e(2.0), e(0.24), p(note, size=760, color=MUTED)))


def build_slides() -> list[base.Slide]:
    slides: list[base.Slide] = []

    s = base.Slide("封面")
    add_bg(s, logo=False)
    s.add(base.rect_shape(s.sid(), "logo plate", e(0.92), e(0.78), e(3.25), e(0.86), base.solid_fill(ICE, alpha=94000), '<a:ln w="6350"><a:solidFill><a:srgbClr val="D9E8F8"/></a:solidFill></a:ln>', radius="roundRect"))
    s.add_picture(base.LOGO_H, "SJTU official horizontal logo", e(1.12), e(0.94), e(2.62), e(0.58))
    s.add(base.text_box(s.sid(), "small label", e(0.94), e(2.15), e(2.6), e(0.28), p("PROPOSAL DEFENSE", size=920, color=CYAN, bold=True)))
    s.add(base.text_box(s.sid(), "cover title", e(0.88), e(2.55), e(8.9), e(1.25), p("论文题目：在此填写研究课题名称", size=3300, color=INK, bold=True)))
    s.add(base.text_box(s.sid(), "cover subtitle", e(0.94), e(4.1), e(7.6), e(0.42), p("深蓝开题答辩模板", size=1420, color=SOFT)))
    meta = p("学院 / 系：在此填写", size=1050, color=MUTED, after=45000) + p("答辩人：姓名    指导教师：姓名 职称", size=1050, color=MUTED, after=45000) + p("日期：2026年 月 日", size=1050, color=MUTED)
    s.add(base.text_box(s.sid(), "cover meta", e(0.96), e(5.18), e(5.6), e(0.92), meta))
    s.add(base.rect_shape(s.sid(), "cover glass line", e(0.96), e(6.47), e(10.8), e(0.018), base.solid_fill(LINE)))
    s.add(base.text_box(s.sid(), "source", e(0.96), e(6.75), e(9.7), e(0.22), p("标识素材来源：上海交通大学视觉形象识别系统（vi.sjtu.edu.cn）", size=760, color="7E92AA")))
    slides.append(s)

    s = base.Slide("目录")
    add_bg(s, "CONTENTS")
    title(s, "汇报提纲", "一屏看清逻辑路径")
    agenda = [
        ("01", "背景", "研究场景与问题提出"),
        ("02", "综述", "现有研究与不足"),
        ("03", "目标", "研究内容与创新点"),
        ("04", "方法", "技术路线与数据设计"),
        ("05", "可行", "基础条件与风险控制"),
        ("06", "计划", "进度安排与成果预期"),
    ]
    for idx, (num, head, desc) in enumerate(agenda):
        col = idx % 3
        row = idx // 3
        lx = 0.9 + col * 3.85
        ty = 3.05 + row * 1.55
        glass_card(s, lx, ty, 3.25, 1.15, f"{num}  {head}", [desc], accent=CYAN if idx % 2 == 0 else BLUE)
    footer(s, 2)
    slides.append(s)

    s = base.Slide("研究背景")
    add_bg(s, "01 / BACKGROUND")
    title(s, "研究背景与问题提出", "从大叙事收束到可研究的问题")
    s.add(base.text_box(s.sid(), "hero statement", e(0.92), e(3.0), e(6.6), e(0.82), p("一句话说明：为什么这个问题现在值得被研究。", size=2150, color=INK, bold=True)))
    s.add(base.text_box(s.sid(), "hero note", e(0.96), e(4.02), e(5.75), e(0.45), p("建议把背景写成“变化 - 冲突 - 空白 - 本文切入”的链条。", size=1080, color=MUTED)))
    glass_card(s, 7.7, 2.28, 4.2, 1.1, "现实驱动", ["行业变化、政策需求、实践痛点"], accent=CYAN)
    glass_card(s, 7.7, 3.72, 4.2, 1.1, "理论驱动", ["已有理论解释边界、争议焦点"], accent=BLUE)
    glass_card(s, 7.7, 5.16, 4.2, 1.1, "研究问题", ["本文聚焦的对象、范围与核心问题"], accent=GOLD)
    footer(s, 3)
    slides.append(s)

    s = base.Slide("研究空白")
    add_bg(s, "02 / GAP")
    title(s, "文献综述与研究空白", "少堆作者，多画结构")
    s.add(base.line_shape(s.sid(), "literature axis", e(1.08), e(3.58), e(11.7), e(3.58), color=LINE, width=12700))
    topics = [
        ("主题一", "核心结论", "可填代表文献"),
        ("主题二", "方法演进", "可填数据/模型"),
        ("主题三", "争议焦点", "可填相反观点"),
        ("空白", "本文切入", "明确不足"),
    ]
    for idx, (a, b, c) in enumerate(topics):
        lx = 1.08 + idx * 2.75
        s.add(base.rect_shape(s.sid(), f"timeline point {idx}", e(lx), e(3.43), e(0.3), e(0.3), base.solid_fill(CYAN if idx == 3 else BLUE), radius="ellipse"))
        s.add(base.text_box(s.sid(), f"timeline head {idx}", e(lx - 0.15), e(2.5), e(1.85), e(0.28), p(a, size=1050, color=INK, bold=True)))
        s.add(base.text_box(s.sid(), f"timeline body {idx}", e(lx - 0.15), e(4.0), e(1.95), e(0.52), p(b, size=900, color=MUTED) + p(c, size=820, color="7E92AA")))
    s.add(base.rect_shape(s.sid(), "gap panel", e(1.0), e(5.45), e(10.95), e(0.72), base.solid_fill(PANEL_2, alpha=78000), '<a:ln w="6350"><a:solidFill><a:srgbClr val="3B6388"/></a:solidFill></a:ln>', radius="roundRect"))
    s.add(base.text_box(s.sid(), "gap text", e(1.24), e(5.64), e(10.2), e(0.22), p("研究空白：请用一句话写清“已有研究尚未解释/验证/解决什么”。", size=1080, color=SOFT, bold=True, align="ctr")))
    footer(s, 4)
    slides.append(s)

    s = base.Slide("目标创新")
    add_bg(s, "03 / OBJECTIVES")
    title(s, "研究目标、内容与创新点", "把目标写得可检验、可完成")
    stat_card(s, 0.95, 3.0, "01", "研究目标", "回答核心问题", CYAN)
    stat_card(s, 3.95, 3.0, "02", "研究内容", "拆成三项任务", BLUE)
    stat_card(s, 6.95, 3.0, "03", "创新点", "说明新增贡献", GOLD)
    body = p("示例写法", size=1150, color=INK, bold=True, after=50000)
    for item in ["构建/验证一个可解释的分析框架", "基于数据、实验或案例形成证据链", "给出理论贡献与实践启示的边界"]:
        body += p(item, size=980, color=MUTED, bullet=True, after=26000)
    s.add(base.text_box(s.sid(), "objective body", e(1.05), e(5.12), e(7.6), e(0.9), body))
    footer(s, 5)
    slides.append(s)

    s = base.Slide("技术路线")
    add_bg(s, "04 / METHOD")
    title(s, "研究方法与技术路线", "从输入到输出保持一条线")
    steps = [
        ("问题界定", "概念/变量"),
        ("数据材料", "样本/来源"),
        ("方法模型", "识别/实验"),
        ("分析验证", "稳健/对照"),
        ("结论产出", "贡献/建议"),
    ]
    for idx, (head, desc) in enumerate(steps):
        lx = 0.92 + idx * 2.32
        s.add(base.rect_shape(s.sid(), f"flow {idx}", e(lx), e(3.18), e(1.78), e(1.04), base.solid_fill(PANEL, alpha=84000), '<a:ln w="6350"><a:solidFill><a:srgbClr val="426D91"/></a:solidFill></a:ln>', radius="roundRect"))
        s.add(base.text_box(s.sid(), f"flow head {idx}", e(lx + 0.14), e(3.42), e(1.5), e(0.26), p(head, size=970, color=INK, bold=True, align="ctr")))
        s.add(base.text_box(s.sid(), f"flow desc {idx}", e(lx + 0.18), e(3.78), e(1.42), e(0.22), p(desc, size=760, color=MUTED, align="ctr")))
        if idx < len(steps) - 1:
            base.line_shape(s.sid(), f"flow line {idx}", e(lx + 1.82), e(3.7), e(lx + 2.22), e(3.7), color=CYAN, width=12700)
            s.add(base.line_shape(s.sid(), f"flow connector {idx}", e(lx + 1.84), e(3.7), e(lx + 2.18), e(3.7), color=CYAN, width=12700))
    s.add(base.text_box(s.sid(), "method note", e(1.08), e(5.25), e(10.2), e(0.42), p("可在这里补充：核心模型、实验设置、变量定义、评价指标或案例选择标准。", size=1050, color=MUTED, align="ctr")))
    footer(s, 6)
    slides.append(s)

    s = base.Slide("数据设计")
    add_bg(s, "05 / DESIGN")
    title(s, "数据、实验或案例设计", "把可行性落到材料和指标")
    glass_card(s, 0.95, 2.85, 3.35, 2.55, "数据来源", ["公开数据库 / 调研 / 实验记录", "样本范围与时间窗口", "数据清洗与缺失处理"], accent=CYAN)
    glass_card(s, 4.7, 2.85, 3.35, 2.55, "研究变量", ["解释变量与被解释变量", "控制变量与中介机制", "指标定义与量化方式"], accent=BLUE)
    glass_card(s, 8.45, 2.85, 3.35, 2.55, "验证策略", ["基准模型或实验流程", "稳健性检验", "替代解释排除"], accent=GOLD)
    footer(s, 7)
    slides.append(s)

    s = base.Slide("可行性")
    add_bg(s, "06 / FEASIBILITY")
    title(s, "可行性分析与风险控制", "评委关心的不是宏大，而是能落地")
    risks = [
        ("研究基础", "已有课程、阅读、预研或数据积累"),
        ("资源条件", "软件、平台、导师支持、数据权限"),
        ("潜在风险", "数据不足、模型不稳定、进度延迟"),
        ("备选方案", "替代样本、替代模型、降级范围"),
    ]
    for idx, (head, desc) in enumerate(risks):
        yy = 2.78 + idx * 0.84
        s.add(base.rect_shape(s.sid(), f"risk row {idx}", e(1.0), e(yy), e(10.7), e(0.55), base.solid_fill(PANEL_2 if idx % 2 == 0 else PANEL, alpha=76000), '<a:ln w="6350"><a:solidFill><a:srgbClr val="34516F"/></a:solidFill></a:ln>', radius="roundRect"))
        s.add(base.text_box(s.sid(), f"risk head {idx}", e(1.28), e(yy + 0.15), e(1.35), e(0.18), p(head, size=850, color=CYAN if idx < 2 else GOLD, bold=True)))
        s.add(base.text_box(s.sid(), f"risk desc {idx}", e(3.0), e(yy + 0.14), e(7.9), e(0.2), p(desc, size=880, color=MUTED)))
    footer(s, 8)
    slides.append(s)

    s = base.Slide("进度计划")
    add_bg(s, "07 / SCHEDULE")
    title(s, "研究计划与阶段成果", "时间表保持克制、清楚、可执行")
    phases = ["阶段一", "阶段二", "阶段三", "阶段四", "阶段五"]
    tasks = [
        ("文献综述", 0, 1),
        ("方案设计", 1, 2),
        ("数据/实验", 2, 3),
        ("分析验证", 3, 4),
        ("论文撰写", 4, 5),
    ]
    left = 2.55
    top = 2.85
    cell = 1.72
    for idx, phase in enumerate(phases):
        s.add(base.text_box(s.sid(), f"phase {idx}", e(left + idx * cell), e(top), e(cell), e(0.22), p(phase, size=780, color=MUTED, align="ctr")))
        s.add(base.line_shape(s.sid(), f"grid {idx}", e(left + idx * cell), e(top + 0.42), e(left + idx * cell), e(6.0), color=LINE, width=6350))
    s.add(base.line_shape(s.sid(), "grid last", e(left + len(phases) * cell), e(top + 0.42), e(left + len(phases) * cell), e(6.0), color=LINE, width=6350))
    for idx, (task, start, end) in enumerate(tasks):
        yy = top + 0.62 + idx * 0.55
        s.add(base.text_box(s.sid(), f"task {idx}", e(0.96), e(yy + 0.05), e(1.3), e(0.18), p(task, size=780, color=MUTED, align="r")))
        s.add(base.rect_shape(s.sid(), f"bar {idx}", e(left + start * cell + 0.12), e(yy), e((end - start) * cell - 0.24), e(0.24), base.solid_fill(CYAN if idx in (0, 4) else BLUE), None, radius="roundRect"))
    footer(s, 9)
    slides.append(s)

    s = base.Slide("预期成果")
    add_bg(s, "08 / OUTCOMES")
    title(s, "预期成果与贡献边界", "说明产出，也说明适用范围")
    glass_card(s, 1.0, 2.9, 3.15, 2.35, "论文成果", ["完成开题报告与学位论文", "形成完整问题 - 方法 - 结论链条"], accent=CYAN)
    glass_card(s, 5.05, 2.9, 3.15, 2.35, "学术贡献", ["补充理论解释或经验证据", "提供可复用研究框架"], accent=BLUE)
    glass_card(s, 9.1, 2.9, 3.15, 2.35, "实践价值", ["形成建议、指标或方案", "明确应用场景与局限"], accent=GOLD)
    footer(s, 10)
    slides.append(s)

    s = base.Slide("致谢")
    add_bg(s, logo=False)
    s.add(base.rect_shape(s.sid(), "thanks logo plate", e(0.92), e(0.82), e(3.25), e(0.86), base.solid_fill(ICE, alpha=94000), '<a:ln w="6350"><a:solidFill><a:srgbClr val="D9E8F8"/></a:solidFill></a:ln>', radius="roundRect"))
    s.add_picture(base.LOGO_H, "SJTU official horizontal logo", e(1.12), e(0.98), e(2.62), e(0.58))
    s.add(base.text_box(s.sid(), "thanks text", e(1.0), e(2.88), e(7.6), e(0.82), p("敬请各位老师批评指正", size=3300, color=INK, bold=True)))
    s.add(base.rect_shape(s.sid(), "thanks accent", e(1.02), e(4.0), e(1.7), e(0.04), base.solid_fill(CYAN)))
    s.add(base.text_box(s.sid(), "thanks contact", e(1.02), e(4.58), e(7.6), e(0.32), p("答辩人：姓名    邮箱：name@sjtu.edu.cn", size=1120, color=MUTED)))
    s.add(base.text_box(s.sid(), "usage", e(1.02), e(6.45), e(10.5), e(0.48), p("标识来源：上海交通大学视觉形象识别系统官方附件。请勿拉伸、重绘、改色或拼接校徽/校名标志。", size=790, color="7E92AA")))
    slides.append(s)

    s = base.Slide("标识来源")
    add_bg(s, "APPENDIX")
    title(s, "标识来源与规范说明", "保留此页便于归档")
    notes = [
        "校徽图片：上海交通大学视觉形象识别系统官方附件“校标-校徽.png”。",
        "横版标志：上海交通大学视觉形象识别系统官方附件“校标-标志中英文横版.png”。",
        "本深蓝模板未重绘、改色或拆分标志；深色页面中用浅色底板承载官方彩色标志。",
        "如需其他格式或最新规范，请以学校官方 VI 页面发布内容为准。",
    ]
    body = ""
    for item in notes:
        body += p(item, size=1120, color=MUTED, bullet=True, after=52000)
    s.add(base.text_box(s.sid(), "source body", e(1.02), e(2.85), e(10.3), e(2.2), body))
    s.add(base.text_box(s.sid(), "source url", e(1.02), e(5.72), e(10.4), e(0.32), p("官方页面：https://vi.sjtu.edu.cn/index.php/articles/bulletin/13", size=1020, color=CYAN, bold=True)))
    footer(s, 12)
    slides.append(s)

    return slides


def main() -> None:
    base.OUT = OUT
    base.CORE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>上海交通大学开题答辩PPT模板_深蓝版</dc:title>
  <dc:subject>开题答辩模板</dc:subject>
  <dc:creator>Codex</dc:creator>
  <cp:keywords>上海交通大学; SJTU; 开题答辩; PPT模板; 深蓝</cp:keywords>
  <dc:description>使用上海交通大学视觉形象识别系统官方附件生成的深蓝开题答辩模板。</dc:description>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-07-10T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-10T00:00:00Z</dcterms:modified>
</cp:coreProperties>"""
    slides = build_slides()
    base.write_pptx(slides)
    print(f"Created {OUT.resolve()} ({OUT.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()

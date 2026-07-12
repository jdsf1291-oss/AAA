# -*- coding: utf-8 -*-
"""
Generate a PowerPoint deck:
TAVR 围术期心电评估与永久起搏器风险预测及管理
Sources are annotated on each slide and compiled on the final references slide.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ---------- Theme ----------
NAVY = RGBColor(0x0B, 0x2E, 0x59)
BLUE = RGBColor(0x1F, 0x5C, 0x99)
TEAL = RGBColor(0x1A, 0x9E, 0x9E)
LIGHT = RGBColor(0xEA, 0xF1, 0xF8)
LIGHT2 = RGBColor(0xF5, 0xF9, 0xFC)
GREY = RGBColor(0x55, 0x5F, 0x6B)
DARK = RGBColor(0x1B, 0x22, 0x2B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ORANGE = RGBColor(0xE1, 0x7A, 0x1E)
RED = RGBColor(0xC0, 0x39, 0x2B)
GREEN = RGBColor(0x2E, 0x8B, 0x57)

FONT = "Microsoft YaHei"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def add_slide():
    return prs.slides.add_slide(BLANK)


def rect(slide, x, y, w, h, fill, line=None):
    from pptx.enum.shapes import MSO_SHAPE
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.75)
    shp.shadow.inherit = False
    return shp


def txt(slide, x, y, w, h, text, size=18, color=DARK, bold=False,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font=FONT, italic=False,
        line_spacing=1.0):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Pt(4); tf.margin_right = Pt(4)
    tf.margin_top = Pt(2); tf.margin_bottom = Pt(2)
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
        r.font.name = font
    return tb


def header(slide, title, kicker=None):
    rect(slide, 0, 0, SW, Inches(1.15), NAVY)
    rect(slide, 0, Inches(1.15), SW, Inches(0.06), TEAL)
    txt(slide, Inches(0.55), Inches(0.18), Inches(11.8), Inches(0.8), title,
        size=26, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    if kicker:
        txt(slide, Inches(0.58), Inches(0.02), Inches(11), Inches(0.3), kicker,
            size=11, color=TEAL, bold=True)


def footer(slide, idx, src=None):
    txt(slide, Inches(0.55), Inches(7.05), Inches(9.5), Inches(0.35),
        "TAVR 围术期心电评估与起搏器管理", size=9, color=GREY)
    txt(slide, Inches(12.2), Inches(7.05), Inches(0.9), Inches(0.35),
        str(idx), size=10, color=GREY, align=PP_ALIGN.RIGHT)
    if src:
        txt(slide, Inches(0.55), Inches(6.78), Inches(12.2), Inches(0.3),
            "来源: " + src, size=9, color=BLUE, italic=True)


def bullets(slide, x, y, w, h, items, size=15, gap=6, color=DARK):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, it in enumerate(items):
        if isinstance(it, tuple):
            level, text = it
        else:
            level, text = 0, it
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.level = level
        p.space_after = Pt(gap)
        p.line_spacing = 1.05
        bullet = "•  " if level == 0 else "–  "
        r = p.add_run()
        r.text = bullet + text
        r.font.size = Pt(size - level * 1)
        r.font.name = FONT
        r.font.color.rgb = color if level == 0 else GREY
        r.font.bold = False
    return tb


def table(slide, x, y, w, rows, col_widths, header_fill=BLUE,
          header_color=WHITE, fsize=12, row_h=Inches(0.42), header_h=Inches(0.5)):
    nrows = len(rows)
    ncols = len(rows[0])
    total_h = header_h + row_h * (nrows - 1)
    gt = slide.shapes.add_table(nrows, ncols, x, y, w, total_h).table
    gt.first_row = False
    gt.horz_banding = False
    # column widths
    tw = sum(col_widths)
    for j, cw in enumerate(col_widths):
        gt.columns[j].width = Emu(int(w * cw / tw))
    gt.rows[0].height = header_h
    for i in range(1, nrows):
        gt.rows[i].height = row_h
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = gt.cell(i, j)
            c.margin_left = Pt(5); c.margin_right = Pt(5)
            c.margin_top = Pt(2); c.margin_bottom = Pt(2)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.fill.solid()
            if i == 0:
                c.fill.fore_color.rgb = header_fill
            else:
                c.fill.fore_color.rgb = LIGHT if i % 2 else WHITE
            tf = c.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if j > 0 else PP_ALIGN.LEFT
            r = p.add_run()
            r.text = val
            r.font.size = Pt(fsize)
            r.font.name = FONT
            r.font.bold = (i == 0)
            r.font.color.rgb = header_color if i == 0 else DARK
    return gt


def chip(slide, x, y, w, h, text, fill, tcolor=WHITE, size=13, bold=True):
    from pptx.enum.shapes import MSO_SHAPE
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.fill.solid(); shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    tf = shp.text_frame; tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold; r.font.name = FONT
    r.font.color.rgb = tcolor
    return shp


slide_no = 0

# ============ SLIDE 1: TITLE ============
s = add_slide()
rect(s, 0, 0, SW, SH, NAVY)
rect(s, 0, Inches(4.55), SW, Inches(0.08), TEAL)
rect(s, 0, Inches(2.0), Inches(0.18), Inches(2.4), TEAL)
txt(s, Inches(0.9), Inches(2.05), Inches(11.5), Inches(1.4),
    "TAVR 围术期心电评估与\n永久起搏器风险预测及管理", size=40, color=WHITE, bold=True,
    line_spacing=1.05)
txt(s, Inches(0.95), Inches(4.75), Inches(11), Inches(0.6),
    "术前 · 术中 · 术后全流程  |  传导阻滞 · 风险评分 · 拔管与起搏决策",
    size=18, color=TEAL, bold=True)
txt(s, Inches(0.95), Inches(5.5), Inches(11.5), Inches(1.0),
    "基于 2020 ACC 专家共识、2019 JACC 专家声明、2021 ESC 起搏指南\n"
    "及 D-PACE / Emory 评分等最新证据整理  ·  2026",
    size=13, color=RGBColor(0xB8,0xC6,0xD8), line_spacing=1.3)
footer(s, "")  # no number on title
slide_no += 1

# ============ SLIDE 2: AGENDA ============
s = add_slide(); slide_no += 1
header(s, "目录", "Agenda")
items_left = [
    ("1", "背景：解剖毗邻与流行病学"),
    ("2", "术前心电评估与危险因素"),
    ("3", "术中评估与监测"),
    ("4", "术后即刻/次日心电评估"),
]
items_right = [
    ("5", "起搏器概率预测（分层数据）"),
    ("6", "风险评分：Emory 与 D-PACE"),
    ("7", "起搏器植入时机分布"),
    ("8", "处理流程与拔管/出院决策"),
]
def agenda_col(x, arr):
    y = Inches(1.7)
    for num, t in arr:
        chip(s, x, y, Inches(0.7), Inches(0.7), num, TEAL, size=22)
        txt(s, x + Inches(0.95), y, Inches(4.6), Inches(0.7), t, size=17,
            color=DARK, bold=True, anchor=MSO_ANCHOR.MIDDLE)
        y += Inches(1.15)
agenda_col(Inches(0.8), items_left)
agenda_col(Inches(7.0), items_right)
footer(s, slide_no)

# ============ SLIDE 3: BACKGROUND ============
s = add_slide(); slide_no += 1
header(s, "背景：解剖毗邻与流行病学", "Anatomy & Epidemiology")
bullets(s, Inches(0.55), Inches(1.5), Inches(6.4), Inches(4.8), [
    "主动脉瓣环与传导系统（希氏束、左束支）解剖上紧密毗邻",
    (1, "瓣膜支架、输送系统、硬导丝对室间隔膜部及传导束产生机械压迫"),
    (1, "损伤多表现为房室阻滞或左束支阻滞，可为一过性或持续性"),
    "术后新发 LBBB 发生率约 27%（4–57%）",
    "永久起搏器（PPM）总体植入率约 11–17%（各研究 2–51%）",
    (1, "自膨胀瓣 (CoreValve/Evolut) 约 28% > 球扩瓣 (SAPIEN) 约 6%"),
    "高度房室阻滞 (HAVB) 发生率 9–26%，是 PPM 的主要原因",
], size=15, gap=9)
# right stat panel
rect(s, Inches(7.3), Inches(1.5), Inches(5.4), Inches(4.7), LIGHT2)
rect(s, Inches(7.3), Inches(1.5), Inches(5.4), Inches(0.55), BLUE)
txt(s, Inches(7.3), Inches(1.5), Inches(5.4), Inches(0.55), "关键数字", size=15,
    color=WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
stats = [("11–17%", "总体 PPM 植入率"), ("9–26%", "高度房室阻滞发生率"),
         ("~27%", "新发 LBBB 发生率"), ("0.5–0.75%", "新发 RBBB（罕见但高危）")]
y = Inches(2.25)
for big, lab in stats:
    txt(s, Inches(7.5), y, Inches(2.4), Inches(0.9), big, size=26, color=TEAL,
        bold=True, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(9.9), y, Inches(2.7), Inches(0.9), lab, size=13, color=DARK,
        anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.95)
footer(s, slide_no, "Circ Rev/CFR 2021; EuroIntervention 2023; JACC 2020")

# ============ SLIDE 4: PRE-OP ============
s = add_slide(); slide_no += 1
header(s, "① 术前心电评估与危险因素", "Pre-procedure")
txt(s, Inches(0.55), Inches(1.35), Inches(12), Inches(0.4),
    "术前 12 导联 ECG 是最重要的风险评估工具；建议术前至少 24h 心电监测以发现隐匿异常",
    size=14, color=BLUE, bold=True)
bullets(s, Inches(0.55), Inches(1.95), Inches(6.3), Inches(4.3), [
    "预示 PPM 的 ECG 危险因素：",
    (1, "术前 RBBB —— 最强预测因子（住院期 HAVB 风险可达 24%）"),
    (1, "一度房室阻滞、PR 延长"),
    (1, "左前分支阻滞 (LAFB)、双分支阻滞"),
    (1, "QRS 增宽（≥140 ms 进入 Emory 评分）"),
    "临床/影像危险因素：",
    (1, "男性、房颤史"),
    (1, "室间隔膜部长度短、LVOT/瓣环钙化（尤其 NCC 下方）"),
], size=14, gap=7)
rect(s, Inches(7.15), Inches(1.95), Inches(5.6), Inches(4.25), LIGHT)
txt(s, Inches(7.35), Inches(2.1), Inches(5.2), Inches(0.5),
    "要点：术前 RBBB", size=16, color=NAVY, bold=True)
bullets(s, Inches(7.35), Inches(2.65), Inches(5.2), Inches(3.4), [
    "TAVR 后传导恶化风险最高的基线特征",
    "新发 RBBB 罕见 (0.5–0.75%)，但一旦出现 46–67% 需 PPM",
    "术前 RBBB 患者应预案：保留临时起搏能力、延长监测",
    "高度 AVB 风险可持续至术后 7 天（自膨胀瓣更甚）",
], size=13, gap=9, color=DARK)
footer(s, slide_no, "2020 ACC 共识 (JACC 2020;76:2391); Kiani JACC Interv 2019")

# ============ SLIDE 5: INTRA-OP ============
s = add_slide(); slide_no += 1
header(s, "② 术中评估与监测", "Intra-procedure")
bullets(s, Inches(0.55), Inches(1.5), Inches(6.3), Inches(4.7), [
    "全程连续心电/血流动力学监测 + 临时经静脉起搏",
    (1, "高危患者建议经右颈内静脉保留临时起搏电极"),
    (1, "若股静脉用于快速起搏，可另置颈内静脉电极"),
    "术中一过性高度 AVB —— 迟发 AVB 的独立预测因子 (OR≈3.5)",
    "可改良的操作因素：",
    (1, "植入深度（越深风险越高，每深 1mm OR≈1.46）"),
    (1, "瓣膜类型（自膨胀 > 球扩）、避免过度预扩张"),
    "术后即刻快速心房起搏测 Wenckebach 点，辅助判断需否 PPM",
], size=14, gap=8)
rect(s, Inches(7.15), Inches(1.5), Inches(5.6), Inches(4.7), LIGHT2)
txt(s, Inches(7.35), Inches(1.65), Inches(5.2), Inches(0.5),
    "术中→术后即刻 12 导联 ECG", size=15, color=NAVY, bold=True)
bullets(s, Inches(7.35), Inches(2.2), Inches(5.2), Inches(3.9), [
    "记录 3 种典型情形，决定后续路径：",
    (1, "无传导紊乱 → 低危"),
    (1, "新发传导异常（LBBB / PR、QRS 延长）→ 监测"),
    (1, "术中出现高度 AVB/完全阻滞 → 高危，多需 PPM"),
    "即刻 ECG 完全无变化：迟发 AVB 风险 <1%",
], size=13, gap=9)
footer(s, slide_no, "2019 Rodés-Cabau (JACC 2019;74:1086); Krishnaswamy JACC Interv 2020")

# ============ SLIDE 6: POST-OP ECG ============
s = add_slide(); slide_no += 1
header(s, "③ 术后即刻/次日心电评估", "Post-procedure ECG")
txt(s, Inches(0.55), Inches(1.35), Inches(12.2), Inches(0.5),
    "核心原则：区分「一过性」与「持续性」——次日 (12–24h) ECG 仍存在的异常才具预测价值",
    size=14, color=BLUE, bold=True)
table(s, Inches(0.55), Inches(2.0), Inches(12.2), [
    ["评估项", "阈值 / 意义", "临床提示"],
    ["术后 PR 间期", "≥230 ms（特异度~95%）", "绝对起搏指征高风险"],
    ["PR 变化 ΔPR", "≥24 ms（敏感度~83%）", "传导恶化风险↑"],
    ["新发束支阻滞", "次日仍持续 (LBBB/RBBB)", "迟发 AVB 独立预测因子"],
    ["一过性异常", "术后出现、次日消退", "不增加迟发 AVB 风险"],
    ["新发一度 AVB", "术前/术后新发", "临时起搏电极保留至 24h"],
    ["即刻 ECG 无变化", "无新发/无延长", "迟发 AVB <1%，可早出院"],
], [1.4, 2.2, 2.2], fsize=13, row_h=Inches(0.52))
footer(s, slide_no, "Jorgensen JACC Interv 2018; 以色列 RBBB 队列 (PMID 40767801); D-PACE 2024")

# ============ SLIDE 7: PPM PROBABILITY ============
s = add_slide(); slide_no += 1
header(s, "④ 起搏器概率预测（分层数据）", "PPM Probability")
table(s, Inches(0.55), Inches(1.5), Inches(12.2), [
    ["人群 / 情形", "需要 PPM 的概率", "备注"],
    ["TAVR 总体（30 天内）", "≈ 11%（均值）", "自 2012 稳定；范围 2–51%"],
    ["自膨胀瓣 vs 球扩瓣", "≈ 28% vs 6%", "瓣膜类型差异显著"],
    ["高度房室阻滞 (HAVB)", "发生率 9–26%", "PPM 主要原因；部分可恢复"],
    ["术前 RBBB", "住院期 HAVB 达 ~24%", "最强预测因子之一"],
    ["新发 RBBB（术后）", "46–67% 需 PPM", "HR≈8.4；中位植入 1 天"],
    ["ECG 无变化 / 低危评分", "< 1–2%", "适合早期拔管与出院"],
], [3.0, 2.2, 3.0], fsize=13, row_h=Inches(0.5))
txt(s, Inches(0.55), Inches(5.7), Inches(12.2), Inches(1.0),
    "要点：概率高度依赖基线传导状态与瓣膜类型。术前 RBBB 合并 PR 延长/LAFB/交替束支阻滞时，"
    "进展至需 PPM 的比例最高；而术后 ECG 始终干净者绝对风险 <2%。",
    size=13, color=DARK, line_spacing=1.2)
footer(s, slide_no, "STS/ACC TVT Registry; CIRCEP 2024 (RBBB); EuroIntervention 2020")

# ============ SLIDE 8: EMORY SCORE ============
s = add_slide(); slide_no += 1
header(s, "⑤ 风险评分（1）：Emory 评分", "Emory Risk Score — 预测是否需 PPM")
table(s, Inches(0.55), Inches(1.5), Inches(6.6), [
    ["因素", "分值"],
    ["术前 RBBB", "2"],
    ["晕厥史", "1"],
    ["QRS ≥ 140 ms", "1"],
    ["瓣膜 oversizing ≥ 16%", "1"],
    ["总分范围", "0 – 5"],
], [3.0, 1.0], fsize=14, row_h=Inches(0.55), header_h=Inches(0.55))
bullets(s, Inches(7.4), Inches(1.55), Inches(5.4), Inches(4.6), [
    "基于 Edwards SAPIEN 3 球扩瓣队列 (Kiani 等)",
    "每增加 1 分，PPM 风险 OR ≈ 2.2",
    "推导队列 AUC 0.778",
    "外部验证表现中等（AUC ≈ 0.61–0.66）",
    "多项验证：单用「术前 RBBB」预测力与整分相当",
    "→ 作术前初筛参考，不宜单独依赖",
], size=14, gap=10)
footer(s, slide_no, "Kiani et al. JACC Cardiovasc Interv 2019;12:2133; 验证: CJC Open 2022; J Interv Cardiol 2020")

# ============ SLIDE 9: D-PACE SCORE ============
s = add_slide(); slide_no += 1
header(s, "⑤ 风险评分（2）：D-PACE 评分", "预测「迟发」高度 AVB（术后 24h–30d）")
txt(s, Inches(0.55), Inches(1.3), Inches(12.2), Inches(0.4),
    "在术后次日 (12–24h) ECG 上计算；用于筛选可 24h 后早出院的低危患者。AUC 0.879 / 0.799",
    size=13, color=BLUE, bold=True)
table(s, Inches(0.55), Inches(1.85), Inches(6.7), [
    ["独立预测因子", "校正 OR"],
    ["次日持续性新发 RBBB", "9.28"],
    ["术前 RBBB", "5.57"],
    ["次日持续性新发 LBBB", "4.49"],
    ["自膨胀瓣", "2.17"],
    ["植入深度（每 mm）", "1.46"],
    ["次日 PR 增加（每 ms）", "1.03"],
    ["术前 PR（每 ms）", "1.02"],
], [3.2, 1.3], fsize=12, row_h=Inches(0.4), header_h=Inches(0.45))
txt(s, Inches(7.5), Inches(1.85), Inches(5.3), Inches(0.4), "风险分层", size=14,
    color=NAVY, bold=True)
table(s, Inches(7.5), Inches(2.3), Inches(5.3), [
    ["评分", "30d AVB 风险", "处理"],
    ["0–3 低危", "<2% (实测<1%)", "可次日出院"],
    ["4–5 中危", "2–5%", "个体化监测"],
    ["≥6 高危", "≥5% (8.7–20.8%)", "延长监测数天"],
], [1.4, 2.0, 2.0], fsize=11.5, row_h=Inches(0.62), header_h=Inches(0.5))
txt(s, Inches(7.5), Inches(5.0), Inches(5.3), Inches(1.2),
    "注：一过性（次日消退）异常不计分；\n房颤/房扑者 PR 不可测，需用 D-PACE AF 版（待验证）。",
    size=12, color=GREY, line_spacing=1.2)
footer(s, slide_no, "D-PACE: EuroIntervention 2024 (Bologna/Catania, n=1290/936)")

# ============ SLIDE 10: TIMING ============
s = add_slide(); slide_no += 1
header(s, "⑥ 起搏器植入时机分布", "Timing of PPM")
txt(s, Inches(0.55), Inches(1.3), Inches(12.2), Inches(0.4),
    "在「最终需要起搏器」的患者中，植入时间的累计分布：", size=14, color=BLUE, bold=True)
# horizontal bars
bars = [("24 小时内", 33, TEAL), ("48 小时内", 50, BLUE),
        ("住院期间(中位~2天)", 90, NAVY), ("2 周内(含迟发)", 98, GREEN)]
y = Inches(2.0)
maxw = Inches(8.2)
for lab, pct, col in bars:
    txt(s, Inches(0.55), y, Inches(2.6), Inches(0.55), lab, size=13, color=DARK,
        bold=True, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, Inches(3.2), y+Inches(0.05), maxw, Inches(0.45), LIGHT)
    rect(s, Inches(3.2), y+Inches(0.05), Emu(int(maxw*pct/100)), Inches(0.45), col)
    txt(s, Inches(3.2)+Emu(int(maxw*pct/100))+Inches(0.1), y, Inches(1.3),
        Inches(0.55), f"~{pct}%", size=14, color=col, bold=True,
        anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.78)
bullets(s, Inches(0.55), Inches(5.4), Inches(12.2), Inches(1.7), [
    "HAVB 事件本身 60–96% 发生在术后 24h 内；>48h 的迟发约 2–7%",
    "约 90% 在首次住院期植入（中位 2 天，IQR 0.5–3.5）；迟发者中位第 7 天，其中 79.6% 在 14 天内",
    "新发 RBBB 亚组更早更集中：中位植入 1 天，绝大多数在 1 周内",
], size=13, gap=6)
footer(s, slide_no, "CFR 2021; EuroIntervention (managing heart block); CIRCEP 2024")

# ============ SLIDE 11: MANAGEMENT PATHWAY ============
s = add_slide(); slide_no += 1
header(s, "⑦ 处理流程：Rodés-Cabau / 2020 ACC 五组分流", "Management Pathway")
groups = [
    ("组1", "无 RBBB\n且术后无 ECG 改变", "术后即拔临时电极，遥测 24h；\n无新异常→出院", GREEN),
    ("组2", "术前 RBBB\n术后无 ECG 改变", "临时电极≥24h；再遥测24h\n无变化→总≥48h出院", ORANGE),
    ("组3", "术前 RBBB/LBBB/IVCD/\n一度AVB 且有 ECG 改变", "临时电极≥24h；变化消退\n再观察24h；持续→倾向PPM", ORANGE),
    ("组4", "新发 LBBB", "临时起搏1天+遥测/每日ECG;\n加重或PR延长→高危", BLUE),
    ("组5", "术中/术后\n高度AVB或完全阻滞", "任何时间出现\n→ 永久起搏器 (PPM)", RED),
]
x = Inches(0.4)
cw = Inches(2.48)
gap = Inches(0.05)
for tag, cond, act, col in groups:
    rect(s, x, Inches(1.55), cw, Inches(0.6), col)
    txt(s, x, Inches(1.55), cw, Inches(0.6), tag, size=16, color=WHITE, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x, Inches(2.2), cw, Inches(1.55), LIGHT)
    txt(s, x+Inches(0.05), Inches(2.28), cw-Inches(0.1), Inches(1.45), cond,
        size=12, color=DARK, bold=True, align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)
    rect(s, x, Inches(3.8), cw, Inches(2.0), LIGHT2)
    txt(s, x+Inches(0.05), Inches(3.9), cw-Inches(0.1), Inches(1.9), act,
        size=11.5, color=GREY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP,
        line_spacing=1.15)
    x += cw + gap
txt(s, Inches(0.4), Inches(6.0), Inches(12.4), Inches(0.8),
    "通则：出现高度AVB/完全阻滞的持续病例 = Class I 起搏指征（ESC 2021）；"
    "ESC 建议对高度/完全 AVB 可观察至 7 天判断是否可逆（伴慢逸搏可缩短）。",
    size=12, color=DARK, line_spacing=1.2)
footer(s, slide_no, "Rodés-Cabau JACC 2019;74:1086; Lilly 2020 ACC (JACC 2020;76:2391); ESC Pacing 2021")

# ============ SLIDE 12: RBBB & DISCHARGE ============
s = add_slide(); slide_no += 1
header(s, "⑧ 术前 RBBB 管理与拔管/出院决策", "Baseline RBBB & Discharge")
rect(s, Inches(0.55), Inches(1.5), Inches(5.9), Inches(4.6), LIGHT)
txt(s, Inches(0.75), Inches(1.62), Inches(5.5), Inches(0.5),
    "术前 RBBB（组2）", size=16, color=NAVY, bold=True)
bullets(s, Inches(0.75), Inches(2.15), Inches(5.5), Inches(3.8), [
    "无论 PR/QRS 有无新变化，临时起搏能力+持续监测≥24h",
    "24h 内无 HAVB、PR/QRS 无进行性变化→可考虑拔除",
    "拔管≠出院：再遥测+每日ECG≥24h（总院内≥48h，最少2天）",
    "拔管前提：ICU/step-down、保留静脉通路以便紧急起搏",
    "风险可延续至7天→常延长监测，出院时加动态心电监测",
], size=13, gap=9)
rect(s, Inches(6.75), Inches(1.5), Inches(6.0), Inches(4.6), LIGHT2)
txt(s, Inches(6.95), Inches(1.62), Inches(5.6), Inches(0.5),
    "拔管 / 出院一般原则", size=16, color=NAVY, bold=True)
bullets(s, Inches(6.95), Inches(2.15), Inches(5.6), Inches(3.8), [
    "术后即刻/次日 ECG 无变化、无 RBBB → 遥测24h后可早出院",
    "低危 D-PACE(0–3)：迟发风险<1%，适合次日出院无需继续监测",
    "有持续新发束支阻滞/PR延长 → 延长院内监测(常至1周)",
    "早出院趋势下，头 2 周动态心电监测捕捉迟发 HAVB",
    "任何时间高度AVB/完全阻滞 → 永久起搏器",
], size=13, gap=9)
footer(s, slide_no, "2020 ACC 共识; 2019 JACC 声明; D-PACE 2024")

# ============ SLIDE 13: KEY TAKEAWAYS ============
s = add_slide(); slide_no += 1
header(s, "关键要点总结", "Key Takeaways")
bullets(s, Inches(0.6), Inches(1.6), Inches(12.1), Inches(4.9), [
    "术前 RBBB 是 TAVR 后传导恶化与 PPM 的最强预测因子；应提前预案。",
    "「一过性 vs 持续性」是核心——只有次日 ECG 仍存在的异常才有预测价值。",
    "总体 PPM 率约 11%；术前/新发 RBBB 者可达 46–67%；ECG 干净者 <2%。",
    "时机：需起搏者约 1/3 在 24h、约 1/2 在 48h、~90% 住院期、迟发者约 80% 在 2 周内。",
    "评分工具：术前初筛用 Emory；判断能否早出院用 D-PACE（次日计算，低危<1%）。",
    "床旁硬指标：PR≥230 ms、ΔPR≥24 ms、持续性新发束支阻滞。",
    "术前 RBBB：临时起搏≥24h，拔管后仍需监测≥24h（总≥48h）方可出院。",
    "早出院时代应结合出院后动态心电监测，重点覆盖术后前 2 周。",
], size=15, gap=11)
footer(s, slide_no)

# ============ SLIDE 14: REFERENCES ============
s = add_slide(); slide_no += 1
header(s, "参考文献", "References")
refs = [
    "1. Lilly SM, et al. 2020 ACC Expert Consensus Decision Pathway on Management of Conduction Disturbances in Patients Undergoing TAVR. J Am Coll Cardiol. 2020;76(20):2391-2411.",
    "2. Rodés-Cabau J, et al. Management of Conduction Disturbances Associated With TAVR. J Am Coll Cardiol. 2019;74(8):1086-1106.",
    "3. Marchetti M, et al. Development and validation of the D-PACE scoring system to predict delayed high-grade conduction disturbances after TAVI. EuroIntervention. 2024.",
    "4. Kiani S, et al. Development of a Risk Score to Predict New Pacemaker Implantation After TAVR (Emory Risk Score). JACC Cardiovasc Interv. 2019;12(21):2133-2142.",
    "5. New-Onset RBBB After TAVR: Incidence and Outcomes. Circ Arrhythm Electrophysiol. 2024 (CIRCEP.123.012377).",
    "6. New-Onset RBBB After TAVR: Incidence and Risk Factors for PPI (7 Israeli centers). PMID 40767801.",
    "7. Evaluation and Management of Heart Block After TAVR. Cardiac Failure Review (CFR). 2021.",
    "8. Managing heart block after TAVI: monitoring, device selection and pacemaker indications. EuroIntervention.",
    "9. A Systematic Review of Delayed High-Grade AV Block After TAVI. (PMC10994975).",
    "10. Glikson M, et al. 2021 ESC Guidelines on cardiac pacing and cardiac resynchronization therapy. Eur Heart J. 2021.",
    "11. Jorgensen TH, et al. Immediate post-procedural 12-lead ECG as predictor of late conduction defects after TAVR. JACC Cardiovasc Interv. 2018;11.",
    "12. Emory Risk Score validation. CJC Open 2022; J Interv Cardiol 2020;1807909.",
]
tb = s.shapes.add_textbox(Inches(0.55), Inches(1.4), Inches(12.3), Inches(5.5))
tf = tb.text_frame; tf.word_wrap = True
for i, r in enumerate(refs):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.space_after = Pt(6); p.line_spacing = 1.05
    run = p.add_run(); run.text = r
    run.font.size = Pt(11.5); run.font.name = FONT; run.font.color.rgb = DARK
footer(s, slide_no)

# Disclaimer slide
s = add_slide(); slide_no += 1
rect(s, 0, 0, SW, SH, NAVY)
txt(s, Inches(1.0), Inches(2.6), Inches(11.3), Inches(2.0),
    "免责声明", size=30, color=TEAL, bold=True)
txt(s, Inches(1.0), Inches(3.5), Inches(11.3), Inches(2.5),
    "本幻灯片为文献证据的教学性汇总，数据来自公开发表的指南、专家共识与队列研究，"
    "各研究在定义、瓣膜类型与随访时长上存在异质性。具体临床决策须结合患者个体情况、"
    "所在中心方案与最新指南，由具备资质的医师判断。",
    size=15, color=WHITE, line_spacing=1.4)
footer(s, "")

prs.save("TAVR_ECG_Pacemaker.pptx")
print("Saved TAVR_ECG_Pacemaker.pptx with", len(prs.slides._sldIdLst), "slides")

# -*- coding: utf-8 -*-
"""Build an editable 16:9 PPTX in the Magpie Murders (red/black/white) style."""
import os, math, random
from PIL import Image, ImageDraw

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn

OUT = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(OUT, "assets")
os.makedirs(ASSETS, exist_ok=True)

BLACK = (11, 11, 11)
WHITE = (244, 240, 232)
RED = (207, 20, 43)
GRAY = (150, 146, 140)
DARKGRAY = (70, 68, 66)

SANS = "Microsoft YaHei"
SERIF = "Georgia"


# ---------------------------------------------------------------- assets
def _new(w, h, scale=4):
    im = Image.new("RGBA", (w * scale, h * scale), (0, 0, 0, 0))
    return im, ImageDraw.Draw(im), scale


def _draw_poly(d, pts, color, scale):
    d.polygon([(int(x * scale), int(y * scale)) for x, y in pts], fill=color)


def make_magpie(path, color):
    im, d, s = _new(300, 300)
    c = color + (255,)
    _draw_poly(d, [(126, 166), (14, 248), (100, 232), (28, 292), (136, 216)], c, s)
    # body ellipse approximated by a dense polygon
    pts = []
    for i in range(36):
        a = 2 * math.pi * i / 36
        pts.append((158 + 62 * math.cos(a), 160 + 48 * math.sin(a)))
    _draw_poly(d, pts, c, s)
    d.ellipse((187 * s, 84 * s, 255 * s, 152 * s), fill=c)
    _draw_poly(d, [(249, 114), (296, 127), (249, 141)], c, s)
    _draw_poly(d, [(128, 133), (192, 150), (164, 197), (118, 175)], c, s)
    d.ellipse((224 * s, 104 * s, 234 * s, 114 * s), fill=BLACK + (255,))
    im.save(path)


def make_skyline(path, color):
    im, d, s = _new(1200, 200)
    c = color + (255,)
    random.seed(1888)
    x = 0
    while x < 1200:
        w = random.randint(18, 54)
        h = random.randint(38, 138)
        d.rectangle((x * s, (200 - h) * s, (x + w) * s, 200 * s), fill=c)
        if random.random() < .22:
            d.polygon([(x * s, (200 - h) * s), ((x + w // 2) * s, (200 - h - 28) * s),
                       ((x + w) * s, (200 - h) * s)], fill=c)
        x += w + random.randint(2, 9)
    im.save(path)


def make_knife(path, color):
    im, d, s = _new(100, 400)
    c = color + (255,)
    _draw_poly(d, [(50, 0), (74, 82), (64, 186), (50, 208), (36, 186), (26, 82)], c, s)
    d.rectangle((31 * s, 202 * s, 69 * s, 366 * s), fill=c)
    d.ellipse((31 * s, 352 * s, 69 * s, 390 * s), fill=c)
    im.save(path)


def make_lady(path, color):
    im, d, s = _new(120, 260)
    c = color + (255,)
    _draw_poly(d, [(34, 30), (60, 4), (86, 30), (90, 38), (30, 38)], c, s)
    d.ellipse((31 * s, 14 * s, 89 * s, 44 * s), fill=c)
    d.ellipse((81 * s, 38 * s, 99 * s, 56 * s), fill=c)
    d.polygon([(48 * s, 68 * s), (72 * s, 68 * s), (94 * s, 116 * s),
               (108 * s, 260 * s), (12 * s, 260 * s), (26 * s, 116 * s)], fill=c)
    im.save(path)


def make_newspaper(path, color):
    im, d, s = _new(200, 240)
    paper = WHITE + (255,)
    ink = BLACK + (255,)
    d.rectangle((0, 0, 200 * s, 240 * s), fill=paper)
    d.rectangle((10 * s, 10 * s, 190 * s, 36 * s), fill=ink)
    d.rectangle((10 * s, 46 * s, 96 * s, 116 * s), fill=ink)
    for y in (126, 144, 162, 180, 198):
        d.rectangle((10 * s, y * s, 190 * s, (y + 8) * s), fill=ink)
    im.save(path)


MAGPIE_RED = os.path.join(ASSETS, "magpie_red.png")
MAGPIE_WHITE = os.path.join(ASSETS, "magpie_white.png")
SKYLINE = os.path.join(ASSETS, "skyline_white.png")
KNIFE = os.path.join(ASSETS, "knife_white.png")
LADY = os.path.join(ASSETS, "lady_white.png")
PAPER = os.path.join(ASSETS, "newspaper_white.png")
make_magpie(MAGPIE_RED, RED)
make_magpie(MAGPIE_WHITE, WHITE)
make_skyline(SKYLINE, WHITE)
make_knife(KNIFE, WHITE)
make_lady(LADY, WHITE)
make_newspaper(PAPER, WHITE)


# ---------------------------------------------------------------- helpers
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


def rgb(c):
    return RGBColor(*c)


def set_alpha_alpha(shape, alpha):
    """alpha 0..100000 (percent opacity)."""
    solid = shape.fill._xPr
    pass


def rect(slide, x, y, w, h, fill=None, line=None, lw=1.0, rot=0):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(fill)
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = rgb(line)
        sh.line.width = Pt(lw)
    sh.rotation = rot
    return sh


def oval(slide, x, y, w, h, fill=None, line=None, lw=1.0):
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(fill)
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = rgb(line)
        sh.line.width = Pt(lw)
    return sh


def txt(slide, text, x, y, w, h, size=18, color=WHITE, bold=False, font=SANS,
        align=PP_ALIGN.LEFT, spacing=0, line_spacing=1.15, valign=MSO_ANCHOR.TOP,
        italic=False):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = valign
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.alignment = align
        p.line_spacing = line_spacing
        for r in p.runs:
            r.font.size = Pt(size)
            r.font.name = font
            r.font.bold = bold
            r.font.italic = italic
            r.font.color.rgb = rgb(color)
            rPr = r._r.get_or_add_rPr()
            if spacing:
                rPr.set("spc", str(int(spacing * 100)))
            ea = rPr.find(qn("a:ea"))
            if ea is None:
                ea = OxmlElement("a:ea")
                rPr.append(ea)
            ea.set("typeface", font)
    return tb


def pic(slide, path, x, y, w=None, h=None, rot=0):
    if w and h:
        return slide.shapes.add_picture(path, Inches(x), Inches(y), width=Inches(w), height=Inches(h))
    if w:
        return slide.shapes.add_picture(path, Inches(x), Inches(y), width=Inches(w))
    return slide.shapes.add_picture(path, Inches(x), Inches(y), height=Inches(h))


def clear_transition(slide):
    for node in list(slide._element):
        if node.tag.endswith("}transition"):
            slide._element.remove(node)


def set_transition(slide, kind):
    """Different slide transition per page to keep the suspense pacing varied."""
    clear_transition(slide)
    t = OxmlElement("p:transition")
    t.set("spd", "med")
    t.set("advClick", "1")
    if kind == "fade":
        t.append(OxmlElement("p:fade"))
    elif kind == "cut":
        t.append(OxmlElement("p:cut"))
    elif kind == "dissolve":
        t.append(OxmlElement("p:dissolve"))
    elif kind == "blinds":
        e = OxmlElement("p:blinds"); e.set("dir", "horz"); t.append(e)
    elif kind == "checker":
        e = OxmlElement("p:checker"); e.set("dir", "horz"); t.append(e)
    elif kind == "push":
        e = OxmlElement("p:push"); e.set("dir", "r"); t.append(e)
    elif kind.startswith("wipe_"):
        e = OxmlElement("p:wipe"); e.set("dir", kind.split("_")[1]); t.append(e)
    else:
        t.append(OxmlElement("p:fade"))
    slide._element.append(t)


def clear_timing(slide):
    for node in list(slide._element):
        if node.tag == qn("p:timing"):
            slide._element.remove(node)


def _sub(parent, tag, **attrs):
    el = OxmlElement(tag)
    for k, v in attrs.items():
        el.set(k, str(v))
    parent.append(el)
    return el


def add_text_reveal(slide, start=120, step=165, duration=430):
    """Staggered 'fade + float up' entrance: shapes revealed row by row, top to bottom."""
    clear_timing(slide)
    bands = {}
    for sh in slide.shapes:
        if (sh.left <= 0 and sh.top <= 0 and
                sh.width >= prs.slide_width * 0.99 and sh.height >= prs.slide_height * 0.99):
            continue  # keep full-bleed background static
        key = round(sh.top / Inches(0.22))
        bands.setdefault(key, []).append(sh)
    waves = []
    for i, key in enumerate(sorted(bands)):
        spids = [sh.shape_id for sh in sorted(bands[key], key=lambda s: s.left)]
        waves.append((start + i * step, spids))

    tm = OxmlElement("p:timing")
    tnLst = _sub(tm, "p:tnLst")
    rootPar = _sub(tnLst, "p:par")
    rootCtn = _sub(rootPar, "p:cTn", id="1", dur="indefinite", restart="never",
                   nodeType="tmRoot")
    rootChildren = _sub(rootCtn, "p:childTnLst")
    seq = _sub(rootChildren, "p:seq", concurrent="1", nextAc="seek")
    seqCtn = _sub(seq, "p:cTn", id="2", dur="indefinite", nodeType="mainSeq")
    seqChildren = _sub(seqCtn, "p:childTnLst")

    topPar = _sub(seqChildren, "p:par")
    topCtn = _sub(topPar, "p:cTn", id="3", fill="hold")
    topCond = _sub(topCtn, "p:stCondLst")
    beginCond = _sub(topCond, "p:cond", evt="onBegin", delay="0")
    _sub(_sub(beginCond, "p:tgtEl"), "p:sldTgt")
    topChildren = _sub(topCtn, "p:childTnLst")

    cid = [4]

    def nid():
        v = cid[0]
        cid[0] += 1
        return str(v)

    for delay, spids in waves:
        wavePar = _sub(topChildren, "p:par")
        waveCtn = _sub(wavePar, "p:cTn", id=nid(), fill="hold")
        _sub(_sub(waveCtn, "p:stCondLst"), "p:cond", delay=str(delay))
        waveChildren = _sub(waveCtn, "p:childTnLst")
        for spid in spids:
            one = _sub(waveChildren, "p:par")
            ctn = _sub(one, "p:cTn", id=nid(), fill="hold")
            _sub(_sub(ctn, "p:stCondLst"), "p:cond", delay="0")
            effChildren = _sub(ctn, "p:childTnLst")

            fade = _sub(effChildren, "p:animEffect", transition="in", filter="fade")
            fb = _sub(fade, "p:cBhvr")
            _sub(fb, "p:cTn", id=nid(), dur=str(duration))
            _sub(_sub(fb, "p:tgtEl"), "p:spTgt", spid=str(spid))

            anim = _sub(effChildren, "p:anim", calcmode="lin", valueType="num")
            ab = _sub(anim, "p:cBhvr", additive="base")
            _sub(ab, "p:cTn", id=nid(), dur=str(duration), fill="hold")
            _sub(_sub(ab, "p:tgtEl"), "p:spTgt", spid=str(spid))
            _sub(_sub(ab, "p:attrNameLst"), "p:attrName").text = "ppt_y"
            tavLst = _sub(anim, "p:tavLst")
            _sub(_sub(_sub(tavLst, "p:tav", tm="0"), "p:val"),
                 "p:strVal", val="1+#ppt_h/10")
            _sub(_sub(_sub(tavLst, "p:tav", tm="100000"), "p:val"),
                 "p:strVal", val="#ppt_y")

    prevCond = _sub(seq, "p:prevCondLst")
    pc = _sub(prevCond, "p:cond", evt="onPrev", delay="0")
    _sub(_sub(pc, "p:tgtEl"), "p:sldTgt")
    nextCond = _sub(seq, "p:nextCondLst")
    nc = _sub(nextCond, "p:cond", evt="onNext", delay="0")
    _sub(_sub(nc, "p:tgtEl"), "p:sldTgt")

    slide._element.append(tm)


def base(page, left="CASE FILE · UNSOLVED", right="WHITECHAPEL / LONDON"):
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, 13.333, 7.5, BLACK)
    rect(s, 0, 0, 0.09, 7.5, RED)
    rect(s, 0.72, 0.44, 0.09, 0.09, RED, rot=45)
    txt(s, left, 0.95, 0.34, 6.0, 0.28, 9.5, GRAY, spacing=.22)
    txt(s, right, 8.4, 0.34, 4.2, 0.28, 9.5, GRAY, align=PP_ALIGN.RIGHT, spacing=.18)
    rect(s, 0.72, 7.02, 11.9, 0.012, DARKGRAY)
    txt(s, "喜鹊谋杀案风格 · 新闻传播视角", 0.72, 7.10, 7.0, 0.25, 9.5, GRAY, spacing=.12)
    txt(s, "%02d / 09" % page, 10.7, 7.08, 1.9, 0.28, 12, WHITE, font=SERIF,
        align=PP_ALIGN.RIGHT)
    txt(s, "09", 12.64, 7.08, 0.3, 0.25, 9.5, RED, font=SERIF)
    return s


def eyebrow(s, text):
    txt(s, text, 0.72, 0.72, 9.0, 0.28, 11, RED, font=SERIF, spacing=.26)


# ---------------------------------------------------------------- slides
# 01 cover
s = base(1, "OPEN CASE · 1888", "THE MAKING OF A MONSTER")
rect(s, 0.72, 1.16, 5.8, 0.025, RED)
txt(s, "JACK", 0.66, 1.28, 7.2, 1.75, 108, RED, font=SERIF, spacing=-.03)
txt(s, "THE RIPPER", 0.72, 2.90, 8.0, 1.05, 52, RED, font=SERIF, spacing=.06)
txt(s, "雾都迷案", 0.72, 4.10, 7.0, 0.85, 44, WHITE, font="SimSun", spacing=.14)
txt(s, "谁制造了「开膛手杰克」", 0.75, 5.00, 7.0, 0.4, 17, WHITE, spacing=.20)
rect(s, 0.72, 5.55, 4.9, 0.02, RED)
txt(s, "新闻传播视角下的百年悬案\n红 · 黑 · 白 ／ 悬疑推理叙事 ／ 第四次作业",
    0.72, 5.72, 7.0, 0.95, 12, GRAY, spacing=.12, line_spacing=1.5)
pic(s, MAGPIE_RED, 8.30, 0.55, w=4.45)
kn = pic(s, KNIFE, 10.62, 1.50, w=0.62, h=3.7)
kn.rotation = 14
txt(s, "1888", 9.75, 6.00, 2.8, 0.9, 64, DARKGRAY, font=SERIF, align=PP_ALIGN.RIGHT)

# 02 contents
s = base(2, "CONTENTS / 案件卷宗", "CASE FILES 01—05")
eyebrow(s, "CONTENTS / 案件卷宗")
txt(s, "五页卷宗，", 0.72, 1.10, 5.0, 0.85, 44, WHITE, font="SimSun", spacing=.10)
txt(s, "一桩悬案", 4.55, 1.10, 4.0, 0.85, 44, RED, font="SimSun", spacing=.10)
txt(s, "从犯罪现场到新闻现场——凶手未获判决，名字却已流传百年",
    0.72, 2.02, 10.6, 0.35, 13, GRAY, spacing=.10)
items = [
    ("01", "案发：五个人，五个夜晚", "1888 年白教堂，被简化为「经典五案」的五条生命。"),
    ("02", "命名：一封红墨水信", "「Dear Boss」如何把一个化名变成全球品牌。"),
    ("03", "造势：小报制造「杰克」", "道德剧、医生传说、移民指控与数百封模仿信。"),
    ("04", "失真：标签盖过受害者", "她们是谁？关于五名女性的传播学追问。"),
    ("05", "反思：悬案为何不散", "高热度不等于硬事实：一份媒介素养清单。"),
    ("＋", "结语：雾仍未散", "当叙事比真相更早结案。"),
]
for idx, (n, title, desc) in enumerate(items):
    col = idx % 2
    row = idx // 2
    x = 0.72 + col * 6.05
    y = 2.62 + row * 1.32
    rect(s, x, y, 5.55, 0.012, DARKGRAY)
    txt(s, n, x, y + 0.12, 0.95, 0.6, 34, RED, font=SERIF)
    txt(s, title, x + 1.05, y + 0.14, 4.5, 0.38, 19, WHITE, spacing=.04)
    txt(s, desc, x + 1.05, y + 0.58, 4.45, 0.55, 11.5, GRAY, line_spacing=1.4)
pic(s, MAGPIE_WHITE, 10.35, 5.05, w=2.1)

# 03 timeline
s = base(3, "CASE 01 / 1888.08—11", "来源：BBC History / Britannica")
eyebrow(s, "CASE 01 / THE CANONICAL FIVE")
txt(s, "案发：", 0.72, 1.12, 2.1, 0.8, 40, WHITE, font="SimSun")
txt(s, "五个人", 2.75, 1.12, 2.6, 0.8, 40, RED, font="SimSun")
txt(s, "，五个夜晚", 5.25, 1.12, 4.5, 0.8, 40, WHITE, font="SimSun")
txt(s, "所谓「经典五案」（canonical five）是后世归纳，并非当时的定论",
    0.72, 2.00, 11.0, 0.35, 13, GRAY)
rect(s, 0.78, 2.62, 11.9, 0.015, RED)
events = [
    ("08.31", "玛丽·安·尼科尔斯", "Mary Ann Nichols · 43岁 · 巴克街"),
    ("09.08", "安妮·查普曼", "Annie Chapman · 47岁 · 汉伯里街"),
    ("09.30", "伊丽莎白·斯特赖德", "Elizabeth Stride · 44岁 · 伯纳街"),
    ("09.30", "凯瑟琳·埃多斯", "Catherine Eddowes · 46岁 · 米特广场"),
    ("11.09", "玛丽·简·凯利", "Mary Jane Kelly · 25岁 · 米勒公寓"),
]
for i, (d, name, en) in enumerate(events):
    x = 0.80 + i * 2.42
    rect(s, x - 0.075, 2.545, 0.15, 0.15, RED, rot=45)
    txt(s, d, x, 2.86, 2.2, 0.35, 19, RED, font=SERIF)
    txt(s, name, x, 3.28, 2.35, 0.55, 15.5, WHITE, spacing=.02)
    txt(s, en, x, 3.85, 2.35, 0.8, 10, GRAY, font=SERIF, line_spacing=1.4)
txt(s, "1888", 4.55, 3.25, 4.3, 1.5, 92, (34, 33, 32), font=SERIF, align=PP_ALIGN.CENTER)
txt(s, "时间 1888.08.31—11.09    地点 白教堂一带，半径约一英里    共性 深夜遇害、割喉、剖腹    定名 后世称「经典五案」",
    0.72, 4.62, 11.9, 0.35, 11.5, GRAY)
rect(s, 0.72, 5.08, 0.035, 1.10, RED)
txt(s, "「她们的生平被缠进了一张由假设、谣言与无根据揣测织成的网里。」",
    1.05, 5.06, 10.5, 0.45, 16, WHITE)
txt(s, "HALLIE RUBENHOLD · 《THE FIVE》", 1.05, 5.62, 7.0, 0.3, 10, RED, spacing=.18)
pic(s, SKYLINE, 0.0, 5.62, w=13.333, h=1.2)

# 04 letter
s = base(4, "CASE 02 / THE NAME", "来源：The National Archives / Britannica")
eyebrow(s, "CASE 02 / A LETTER IN RED INK")
txt(s, "命名：", 0.72, 1.12, 2.1, 0.8, 40, WHITE, font="SimSun")
txt(s, "一封红墨水信", 2.75, 1.12, 5.0, 0.8, 40, RED, font="SimSun")
rect(s, 0.92, 2.02, 4.85, 4.35, (124, 16, 29))
rect(s, 0.72, 1.84, 4.85, 4.35, WHITE)
txt(s, "Dear Boss,", 1.06, 2.14, 4.0, 0.55, 29, RED, font=SERIF)
txt(s, "1888.09.25 · 寄给中央新闻社", 1.06, 2.74, 4.2, 0.3, 11, DARKGRAY)
rect(s, 1.06, 3.12, 4.1, 0.012, DARKGRAY)
txt(s, "「我听说警方已经抓到我了，\n他们只是还拿我没辙。\n下一次，我要把女士的耳朵剪下来……\n别介意我用这个行号署名。」",
    1.06, 3.30, 4.1, 1.85, 14, BLACK, line_spacing=1.45)
txt(s, "Yours truly,\nJack the Ripper", 1.06, 5.28, 4.2, 0.8, 17, RED, font=SERIF,
    line_spacing=1.2)
stamp = rect(s, 4.25, 5.55, 1.05, 0.42, None, RED, 1.5, rot=-8)
txt(s, "HOAX?", 4.28, 5.60, 1.0, 0.3, 11, RED, font=SERIF, align=PP_ALIGN.CENTER)
steps = [
    ("09.25", "信件署期，用红墨水写成", "两页纸，拼写与标点错误百出。"),
    ("09.27", "中央新闻社收件", "最初被当作大量恶作剧信件之一。"),
    ("09.30", "埃多斯遇害，右耳部分被割", "信中「剪耳朵」的细节突然显得「应验」。"),
    ("10.01", "警方印发信件传单", "全国报纸全文转载，名字传遍世界。"),
]
for i, (d, t, p) in enumerate(steps):
    y = 1.98 + i * 1.05
    rect(s, 6.05, y + 0.98, 6.5, 0.01, DARKGRAY)
    txt(s, d, 6.05, y, 1.25, 0.35, 22, RED, font=SERIF)
    txt(s, t, 7.35, y + 0.02, 5.1, 0.35, 15, WHITE)
    txt(s, p, 7.35, y + 0.43, 5.1, 0.5, 11, GRAY)
rect(s, 9.05, 6.30, 3.55, 0.42, None, RED, 1.5)
txt(s, "警方后来认定：极可能是记者伪造", 9.12, 6.37, 3.4, 0.3, 10.5, RED,
    align=PP_ALIGN.CENTER)

# 05 press
s = base(5, "CASE 03 / THE PRESS", "来源：BBC Culture 2026 / BBC History")
eyebrow(s, "CASE 03 / THE PRESS")
txt(s, "JACK", 0.66, 1.02, 8.2, 1.85, 116, RED, font=SERIF)
txt(s, "THE NAME THAT SOLD THE PAPERS", 0.78, 2.78, 9.0, 0.35, 13, WHITE, spacing=.28)
rect(s, 8.55, 1.55, 0.03, 1.55, WHITE)
txt(s, "「谋杀在那个秋天就被写成了小说，\n成了一出至今仍在回响的道德剧。」",
    8.82, 1.55, 4.0, 1.2, 14, WHITE, line_spacing=1.5)
txt(s, "CHRISTOPHER FRAYLING · BBC CULTURE, 2026", 8.82, 2.78, 4.0, 0.35, 9, RED, spacing=.10)
cols = [
    ("① 小报的连续剧", "一桩凶案被拆成每日连载：醒目标题、悬赏、读者来信、现场想象，读者追更「下一集」。"),
    ("② 现成的嫌疑人", "「刀法精准 ⇒ 医生」；「东区贫困与移民 ⇒ 外来者」。报纸先给出剧本，警方与公众照着找人。"),
    ("③ 模仿的信件潮", "「Jack the Ripper」成为可复用的署名（trade name）；警方与报社收到数百封模仿信，真假难辨。"),
]
for i, (t, p) in enumerate(cols):
    x = 0.72 + i * 4.08
    rect(s, x, 3.72, 3.75, 0.035, RED)
    txt(s, t, x, 3.92, 3.7, 0.4, 18, WHITE)
    txt(s, p, x, 4.42, 3.65, 1.6, 11.5, GRAY, line_spacing=1.65)

# 06 suspects
s = base(6, "CASE 04 / SUSPECT FILES", "来源：BBC Culture 2026 · Frayling")
eyebrow(s, "CASE 04 / THE SUSPECT LIST")
txt(s, "名册：", 0.72, 1.12, 2.1, 0.8, 40, WHITE, font="SimSun")
txt(s, "媒体与警方共同开出", 2.75, 1.12, 7.2, 0.8, 34, RED, font="SimSun")
txt(s, "1894 年，警方档案中一份备忘录写下三名「认真考虑过」的嫌疑人",
    0.72, 1.99, 11.0, 0.32, 12, GRAY)
sus = [
    ("FILE 01", "蒙塔古·德鲁伊特", "Montague Druitt · 律师",
     "被误认为医生；1888 年 12 月尸体在泰晤士河中被发现，其时案件已近尾声。"),
    ("FILE 02", "阿伦·科斯明斯基", "Aaron Kosminski · 波兰裔理发师",
     "生活在白教堂，1899 年死于精神病院；后世「DNA 破案」说法争议不断。"),
    ("FILE 03", "迈克尔·奥斯特罗格", "Michael Ostrog · 俄国医生／惯犯",
     "身份与行踪均充满不确定性，更多来自当时的传闻网络。"),
]
for i, (n, name, role, desc) in enumerate(sus):
    x = 0.72 + i * 4.08
    rect(s, x, 2.48, 3.75, 3.15, None, DARKGRAY, 1.0)
    rect(s, x, 2.48, 3.75, 0.055, RED)
    txt(s, n, x + 0.24, 2.76, 2.5, 0.3, 12, RED, font=SERIF, spacing=.18)
    txt(s, name, x + 0.24, 3.18, 3.3, 0.65, 19, WHITE, spacing=.02)
    txt(s, role, x + 0.24, 3.86, 3.3, 0.5, 10.5, GRAY)
    txt(s, desc, x + 0.24, 4.48, 3.25, 1.0, 11.5, GRAY, line_spacing=1.5)
    txt(s, "%02d" % (i + 1), x + 2.45, 4.62, 1.05, 0.8, 50, DARKGRAY,
        font=SERIF, align=PP_ALIGN.RIGHT)
    rect(s, x + 0.22, 5.06, 1.55, 0.38, None, RED, 1.3, rot=-5)
    txt(s, "NO EVIDENCE", x + 0.27, 5.13, 1.45, 0.25, 9.5, RED, font=SERIF)
txt(s, "200+", 0.72, 5.88, 2.0, 0.8, 42, RED, font=SERIF)
txt(s, "百余年间被点名的嫌疑人", 2.35, 6.18, 4.2, 0.3, 12, GRAY)
txt(s, "名单是传播史，", 8.0, 5.92, 4.6, 0.4, 18, WHITE)
txt(s, "不是破案史。", 8.0, 6.30, 4.6, 0.4, 18, RED)

# 07 victims
s = base(7, "CASE 05 / THE VOICELESS", "来源：BBC Radio 4 / Hallie Rubenhold《The Five》")
eyebrow(s, "CASE 05 / THE VOICELESS")
txt(s, "失语：", 0.72, 1.12, 2.1, 0.8, 40, WHITE, font="SimSun")
txt(s, "标签盖过人", 2.75, 1.12, 5.0, 0.8, 40, RED, font="SimSun")
names = [("波莉", "MARY ANN NICHOLS"), ("安妮", "ANNIE CHAPMAN"),
         ("丽兹", "ELIZABETH STRIDE"), ("凯特", "CATHERINE EDDOWES"),
         ("玛丽", "MARY JANE KELLY")]
for i, (n, en) in enumerate(names):
    x = 0.72 + i * 2.42
    pic(s, LADY, x + 0.78, 2.02, w=0.86, h=1.85)
    txt(s, n, x, 3.96, 2.4, 0.35, 15, WHITE, align=PP_ALIGN.CENTER)
    txt(s, en, x, 4.30, 2.4, 0.3, 9, RED, font=SERIF, align=PP_ALIGN.CENTER)
txt(s, "“PROSTITUTES”", 0.72, 4.82, 6.0, 0.6, 30, DARKGRAY, font=SERIF, spacing=.14)
rect(s, 0.70, 5.10, 5.75, 0.035, RED, rot=355)
txt(s, "「只有两名受害者的资料显示曾有性工作证据；其余人只是因为在街头露宿而被贴上标签。」",
    0.72, 5.60, 8.6, 0.75, 14.5, WHITE, line_spacing=1.5)
rect(s, 9.65, 5.56, 0.025, 1.05, RED)
txt(s, "传播学追问：\n谁在定义「受害者」？\n谁的故事被凶手的魅力掩盖？",
    9.88, 5.56, 2.75, 1.15, 11, GRAY, line_spacing=1.5)

# 08 logic
s = base(8, "VERDICT / 传播公式", "媒介素养 · 三步核验")
eyebrow(s, "VERDICT / WHY THE MYTH SURVIVES")
txt(s, "悬案为何", 0.72, 1.12, 3.1, 0.8, 40, WHITE, font="SimSun")
txt(s, "不散", 3.75, 1.12, 2.0, 0.8, 40, RED, font="SimSun")
txt(s, "未完成叙事  ＋  高识别符号\n信息真空  ＋  媒体反复展演  ＝  百年神话",
    0.72, 2.18, 7.5, 1.35, 20, WHITE, font=SERIF, line_spacing=1.65)
txt(s, "符号：红墨水 · 刀 · 雾都 · 匿名信", 0.72, 3.40, 7.4, 0.3, 10.5, GRAY, spacing=.12)
rect(s, 0.72, 3.72, 7.4, 0.02, RED)
txt(s, "高热度 ≠ 硬事实", 0.72, 4.02, 7.4, 0.85, 34, RED, font="SimSun", spacing=.06)
checks = [
    ("01", "看信源", "谁最先发布？是警方档案，还是小报标题？"),
    ("02", "查证据链", "「应验的细节」是否早已公开？能否交叉验证？"),
    ("03", "听谁失语", "被标签化的人有没有机会讲完自己的故事？"),
]
for i, (n, t, p) in enumerate(checks):
    y = 1.75 + i * 1.35
    rect(s, 8.55, y, 4.05, 1.13, None, DARKGRAY, 1.0)
    txt(s, n, 8.82, y + 0.22, 0.65, 0.5, 24, RED, font=SERIF)
    txt(s, t, 9.55, y + 0.18, 2.7, 0.35, 16, WHITE)
    txt(s, p, 9.55, y + 0.58, 2.75, 0.45, 10.5, GRAY)
txt(s, "CASE OPEN · 叙事仍在继续", 8.55, 6.25, 4.05, 0.4, 17, RED, font=SERIF,
    align=PP_ALIGN.RIGHT, spacing=.08)

# 09 end
s = base(9, "THE END / 未结之案", "红黑白 · 悬疑推理风格")
txt(s, "?", 10.55, 0.62, 2.3, 3.0, 210, (92, 16, 25), font=SERIF, align=PP_ALIGN.CENTER)
txt(s, "真相，仍在雾中", 0.72, 1.35, 9.0, 1.0, 46, WHITE, font="SimSun", spacing=.12)
txt(s, "THE FOG HAS NOT LIFTED", 0.75, 2.35, 8.0, 0.4, 19, RED, font=SERIF, spacing=.22)
rect(s, 0.72, 3.02, 5.2, 0.025, RED)
txt(s, "资料来源\nBBC Culture《The creation of the myth of Jack the Ripper》2026\n"
       "BBC History《Jack the Ripper (?)》；BBC News 2016 / 2024\n"
       "Britannica《Who were Jack the Ripper's victims?》\n"
       "The National Archives（UK）：\"Dear Boss\" letter 原件转录\n"
       "BBC Radio 4 / Hallie Rubenhold《The Five》",
    0.72, 3.30, 8.4, 2.9, 11.5, GRAY, line_spacing=1.65)
txt(s, "WHOEVER IT WAS,\nIT WAS NOT THE CHARACTER\nINVENTED BY THE PRESS.",
    8.55, 1.72, 4.05, 1.6, 15, WHITE, font=SERIF, align=PP_ALIGN.RIGHT, line_spacing=1.5)
pic(s, MAGPIE_RED, 10.0, 4.45, w=2.55)

# 每页不同的切换效果：硬切、百叶窗、棋盘、推近、溶解……强化悬疑节奏
TRANSITIONS = [
    "fade", "wipe_r", "wipe_d", "cut", "blinds",
    "checker", "fade", "push", "dissolve",
]
for slide, kind in zip(prs.slides, TRANSITIONS):
    set_transition(slide, kind)
    add_text_reveal(slide)

path = os.path.join(OUT, "雾都迷案_开膛手杰克_多样转场.pptx")
prs.save(path)
print("saved:", path)
print("slides:", len(prs.slides.__iter__.__self__._sldIdLst))

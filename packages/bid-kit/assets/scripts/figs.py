# -*- coding: utf-8 -*-
"""配图：施工总平面布置图 + 工艺流程图（线性道路/管网类工程）"""
import os
from PIL import Image, ImageDraw, ImageFont

NAVY = (31, 78, 121)
LBLUE = (222, 235, 247)
GREY = (89, 89, 89)
LGREY = (205, 205, 205)
ORANGE = (237, 125, 49)
PALE = (226, 239, 218)
WHITE = (255, 255, 255)
W = 1600
FONT_B = r'C:\Windows\Fonts\simhei.ttf'
FONT_R = r'C:\Windows\Fonts\simsun.ttc'


def F(s, bold=False):
    return ImageFont.truetype(FONT_B if bold else FONT_R, s)


def plan_road(out, road, segs, facs, note=None):
    im = Image.new('RGB', (W, 760), WHITE)
    d = ImageDraw.Draw(im)
    d.rectangle([6, 6, W - 7, 753], outline=NAVY, width=3)
    d.text((W // 2, 40), '施工总平面布置示意图', font=F(34, True), fill=NAVY, anchor='mm')

    y0 = 200
    d.line([90, y0, W - 90, y0], fill=GREY, width=7)
    d.text((96, y0 - 48), road, font=F(22), fill=GREY, anchor='lm')
    span = (W - 180) / max(1, len(segs))
    for i, s in enumerate(segs):
        x1, x2 = 90 + i * span + 8, 90 + (i + 1) * span - 8
        d.rectangle([x1, y0 - 13, x2, y0 + 13], fill=NAVY)
        d.text(((x1 + x2) / 2, y0 - 34), s, font=F(20), fill=NAVY, anchor='mm')
    for x, lab in [(90, '起点'), (W - 90, '终点')]:
        d.ellipse([x - 10, y0 - 10, x + 10, y0 + 10], fill=ORANGE)
        d.text((x, y0 + 34), lab, font=F(19), fill=ORANGE, anchor='mm')

    bw = 240
    gap = (W - 120 - bw * len(facs)) / max(1, len(facs) - 1)
    for i, f in enumerate(facs):
        x = 60 + i * (bw + gap)
        d.rounded_rectangle([x, 300, x + bw, 400], radius=8, fill=LBLUE, outline=NAVY, width=2)
        for k, ln in enumerate(f.split('\n')):
            tw = d.textbbox((0, 0), ln, font=F(20))[2]
            d.text((x + (bw - tw) / 2, 325 + k * 28), ln, font=F(20), fill=(0, 0, 0))

    d.line([120, 470, W - 140, 470], fill=LGREY, width=4)
    d.text((W // 2, 500), '作业流向：路面破除 → 管线施工 → 路基基层 → 沥青面层 → 附属工程',
           font=F(22), fill=ORANGE, anchor='mm')
    lx, ly = 420, 560
    d.rectangle([lx, ly, lx + 32, ly + 18], fill=NAVY)
    d.text((lx + 42, ly - 3), '施工段', font=F(20), fill=GREY)
    d.rectangle([lx + 180, ly, lx + 212, ly + 18], outline=NAVY, width=2)
    d.text((lx + 222, ly - 3), '临时设施', font=F(20), fill=GREY)
    d.ellipse([lx + 380, ly + 2, lx + 394, ly + 16], fill=ORANGE)
    d.text((lx + 404, ly - 3), '起终点', font=F(20), fill=GREY)
    d.text((W // 2, 660), note or
           '说明：临时设施沿红线内侧布置并随施工段推进调整；稳定土拌合站设于道路起点外侧，减少运输干扰；'
           '渣土经洗车池冲洗后外运。', font=F(19), fill=GREY, anchor='mm')
    os.makedirs(os.path.dirname(out) or '.', exist_ok=True)
    im.save(out, dpi=(200, 200))
    return out


def flow(out, steps, title='主要工序施工工艺流程图'):
    im = Image.new('RGB', (W, 470), WHITE)
    d = ImageDraw.Draw(im)
    d.rectangle([6, 6, W - 7, 463], outline=NAVY, width=3)
    d.text((W // 2, 40), title, font=F(32, True), fill=NAVY, anchor='mm')
    n = max(1, len(steps))
    bw = min(238, int((W - 90 - 24 * (n - 1)) / n))
    gap = (W - 90 - bw * n) / max(1, n - 1)
    y = 140
    for i, s in enumerate(steps):
        x = 45 + i * (bw + gap)
        d.rounded_rectangle([x, y, x + bw, y + 145], radius=8,
                            fill=LBLUE if i % 2 == 0 else PALE, outline=NAVY, width=2)
        for k, ln in enumerate(s.split('\n')):
            tw = d.textbbox((0, 0), ln, font=F(20))[2]
            d.text((x + (bw - tw) / 2, y + 45 + k * 28), ln, font=F(20), fill=(0, 0, 0))
        if i < n - 1:
            d.line([x + bw + 3, y + 72, x + bw + gap - 4, y + 72], fill=GREY, width=3)
            d.polygon([(x + bw + gap - 4, y + 72), (x + bw + gap - 16, y + 66),
                       (x + bw + gap - 16, y + 78)], fill=GREY)
    d.text((W // 2, 355), '关键工序设质量控制点，逐道工序验收合格后进入下道工序',
           font=F(21), fill=ORANGE, anchor='mm')
    os.makedirs(os.path.dirname(out) or '.', exist_ok=True)
    im.save(out, dpi=(200, 200))
    return out


def site_net(out, title, nodes, note=''):
    """多点作业（管网/小区类）：中心调度 + 作业面辐射"""
    H = 1210
    im = Image.new('RGB', (W, H), WHITE)
    d = ImageDraw.Draw(im)
    d.rectangle([6, 6, W - 7, H - 7], outline=NAVY, width=3)
    d.text((W // 2, 40), title, font=F(34, True), fill=NAVY, anchor='mm')
    cx, cy = W // 2, 300
    d.rounded_rectangle([cx - 130, cy - 52, cx + 130, cy + 52], radius=8, fill=NAVY)
    for k, ln in enumerate(['项目调度中心', '（材料周转库）']):
        tw = d.textbbox((0, 0), ln, font=F(21))[2]
        d.text((cx - tw / 2, cy - 26 + k * 30), ln, font=F(21), fill=WHITE)
    n = max(1, len(nodes))
    for i, txt in enumerate(nodes):
        ang = 3.14159 * (1.15 + 1.70 * i / max(1, n - 1))
        import math
        x = cx + 560 * math.cos(ang)
        y = cy + 175 * math.sin(ang)
        d.line([cx, cy, x, y], fill=LGREY, width=2)
        d.rounded_rectangle([x - 115, y - 45, x + 115, y + 45], radius=8, fill=LBLUE, outline=NAVY, width=2)
        for k, ln in enumerate(txt.split('\n')):
            tw = d.textbbox((0, 0), ln, font=F(19))[2]
            d.text((x - tw / 2, y - 20 + k * 28), ln, font=F(19), fill=(0, 0, 0))
    d.text((W // 2, cy + 205), '作业流向：管线探测 → 沟槽开挖 → 管道安装 → 试压冲洗 → 接驳通水 → 路面恢复',
           font=F(21), fill=ORANGE, anchor='mm')
    y1 = 745
    d.line([70, y1 + 150, W - 70, y1 + 150], fill=GREY, width=8)
    items = [('施工围挡及\n警示标志', 120, ORANGE, WHITE), ('管材堆放区\n（≤20m）', 380, LBLUE, (0, 0, 0)),
             ('沟槽开挖段\n钢板桩支护', 650, PALE, (0, 0, 0)), ('试压泵、\n冲洗设备', 920, LBLUE, (0, 0, 0)),
             ('洗车池及\n三级沉淀池', 1190, PALE, (0, 0, 0)), ('临时便道\n（钢板铺设）', 1440, LBLUE, (0, 0, 0))]
    for txt, x, fill, fg in items:
        d.rectangle([x, y1 + 10, x + 200, y1 + 100], fill=fill, outline=NAVY, width=2)
        for k, ln in enumerate(txt.split('\n')):
            tw = d.textbbox((0, 0), ln, font=F(19))[2]
            d.text((x + (200 - tw) / 2, y1 + 28 + k * 30), ln, font=F(19), fill=fg)
        d.line([x + 100, y1 + 100, x + 100, y1 + 146], fill=GREY, width=2)
    d.text((W // 2, y1 + 195), '市政道路（半幅封闭施工，保留 3.5m 通行）', font=F(21), fill=GREY, anchor='mm')
    d.text((W // 2, H - 50), note or '说明：临时设施随作业面推进移动，不设集中生活区；施工废水经三级沉淀后排放。',
           font=F(19), fill=GREY, anchor='mm')
    os.makedirs(os.path.dirname(out) or '.', exist_ok=True)
    im.save(out, dpi=(200, 200))
    return out

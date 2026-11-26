# -*- coding: utf-8 -*-
"""施工进度计划横道图 —— 标准生成器（后续所有项目统一用这个）

用法：
    python make_gantt.py                 # 生成 37/38/39 三份（示例）
    from make_gantt import gantt         # 或在别的脚本里调用

标准（与 2026-11 已定稿的三份附表一致）：
    1. 只有「计划开工」+「计划工期」，不写计划竣工日期
    2. 整月等宽；首月/末月按该月实际施工天数 ÷ 该月天数 定宽（最窄 22%）
    3. 年份条带：整条通栏，每年一段，年份文字在段内居中；窄段（<150px）竖排两行
    4. 月份名写「x月」，位于本列正中
    5. 横道按天定位，标签「起~止天」写在横道右端（贴边时改左端）
    6. 关键节点按 42 字折行，橙色
    7. 画布 2600px 宽，300dpi；插入 Word 时宽度 19.84cm
"""
import os
import calendar
from datetime import date, timedelta
from PIL import Image, ImageDraw, ImageFont

NAVY = (31, 78, 121)
GREY = (89, 89, 89)
LGREY = (205, 205, 205)
ORANGE = (237, 125, 49)
WHITE = (255, 255, 255)
YEARBG = (226, 234, 244)
W = 2600
S = W / 1984.0
FONT_BOLD = r'C:\Windows\Fonts\simhei.ttf'
FONT_NORMAL = r'C:\Windows\Fonts\simsun.ttc'


def F(sz, bold=False):
    return ImageFont.truetype(FONT_BOLD if bold else FONT_NORMAL, int(sz * S))


def month_segments(start, days):
    """按自然月切分：[(year, month, start_offset, end_offset, days_in_month)]"""
    segs = []
    y, m = start.year, start.month
    while True:
        first = date(y, m, 1)
        ny, nm = (y + 1, 1) if m == 12 else (y, m + 1)
        nxt = date(ny, nm, 1)
        sd = max(0, (first - start).days)
        ed = min(days, (nxt - start).days)
        if sd >= days:
            break
        segs.append((y, m, sd, max(sd + 1, ed), calendar.monthrange(y, m)[1]))
        if ed >= days:
            break
        y, m = ny, nm
    return segs


def gantt(out, rows, start, days, note='', title='施工进度计划横道图'):
    """rows: [(工序名, 起始天, 结束天)]；start: date；days: 总工期"""
    months = month_segments(start, days)
    weights = [max(0.22, min(1.0, (ed - sd) / full)) for (y, m, sd, ed, full) in months]

    rowh = int((62 if len(rows) > 6 else 68) * S)
    top = int(215 * S)
    H = top + rowh * len(rows) + int(150 * S)
    im = Image.new('RGB', (W, H), WHITE)
    d = ImageDraw.Draw(im)
    d.rectangle([6, 6, W - 7, H - 7], outline=NAVY, width=3)
    d.text((W // 2, int(44 * S)), title, font=F(40, True), fill=NAVY, anchor='mm')
    d.text((W // 2, int(92 * S)),
           f'计划开工：{start.year} 年 {start.month} 月 {start.day} 日　　计划工期：{days} 日历天',
           font=F(24), fill=GREY, anchor='mm')

    left = int(560 * S)
    right = W - int(90 * S)
    unit = (right - left) / sum(weights)
    cum = [left]
    for w in weights:
        cum.append(cum[-1] + w * unit)

    def x_of(day):
        for i, (y, m, sd, ed, full) in enumerate(months):
            if day <= sd:
                return cum[i]
            if sd < day <= ed:
                return cum[i] + (day - sd) / (ed - sd) * (cum[i + 1] - cum[i])
        return right

    # 年份条带（整条 + 每年一段，年份居中）
    yb_top, yb_bot = int(132 * S), int(170 * S)
    d.rectangle([left, yb_top, right, yb_bot], fill=YEARBG, outline=NAVY, width=2)
    groups = []
    for i, (y, m, sd, ed, full) in enumerate(months):
        if not groups or groups[-1][0] != y:
            groups.append([y, i, i])
        else:
            groups[-1][2] = i
    for y, i0, i1 in groups:
        x1, x2 = cum[i0], cum[i1 + 1]
        seg_w = x2 - x1
        if i0 > 0:
            d.line([(x1, yb_top), (x1, yb_bot)], fill=NAVY, width=2)
        if seg_w < int(150 * S):
            cy = (yb_top + yb_bot) / 2
            d.text(((x1 + x2) / 2, cy - int(11 * S)), str(y), font=F(15, True), fill=NAVY, anchor='mm')
            d.text(((x1 + x2) / 2, cy + int(11 * S)), '年', font=F(15, True), fill=NAVY, anchor='mm')
        else:
            d.text(((x1 + x2) / 2, (yb_top + yb_bot) / 2), f'{y} 年', font=F(20, True), fill=NAVY, anchor='mm')

    # 月份线 + 月名
    for i, (y, m, sd, ed, full) in enumerate(months):
        d.line([(cum[i], yb_bot), (cum[i], top - int(10 * S))], fill=LGREY, width=2)
        d.text(((cum[i] + cum[i + 1]) / 2, yb_bot + int(22 * S)), f'{m}月', font=F(18), fill=GREY, anchor='mm')
    d.line([(right, yb_bot), (right, top - int(10 * S))], fill=LGREY, width=2)

    # 横道
    for i, (name, s, e) in enumerate(rows):
        y = top + i * rowh
        if i % 2 == 0:
            d.rectangle([int(30 * S), y - int(4 * S), W - int(30 * S), y + rowh - int(10 * S)], fill=(247, 249, 252))
        d.text((int(44 * S), y + rowh * 0.42), name, font=F(22), fill=(34, 34, 34), anchor='lm')
        x1, x2 = x_of(s - 1), x_of(e)
        d.rectangle([x1, y + int(6 * S), max(x2, x1 + int(10 * S)), y + int(42 * S)], fill=NAVY)
        lbl = f'{s}~{e}天'
        if x2 > right - int(150 * S):
            d.text((x1 - int(8 * S), y + int(24 * S)), lbl, font=F(18), fill=GREY, anchor='rm')
        else:
            d.text((x2 + int(8 * S), y + int(24 * S)), lbl, font=F(18), fill=GREY, anchor='lm')

    # 关键节点
    if note:
        yy = top + rowh * len(rows) + int(12 * S)
        cur = ''
        for p in note.split('｜'):
            cand = p if not cur else cur + '｜' + p
            if len(cand) > 42 and cur:
                d.text((int(44 * S), yy), cur, font=F(20), fill=ORANGE)
                yy += int(30 * S)
                cur = p
            else:
                cur = cand
        if cur:
            d.text((int(44 * S), yy), cur, font=F(20), fill=ORANGE)
    d.text((W - int(40 * S), H - int(26 * S)), '注：与施工组织设计进度安排一致', font=F(18), fill=GREY, anchor='rs')
    os.makedirs(os.path.dirname(out) or '.', exist_ok=True)
    im.save(out, dpi=(300, 300))
    return out


if __name__ == '__main__':
    jobs = [
        ('37', date(2026, 12, 20), 450, [
            ('施工准备', 1, 30), ('软基处理（深层搅拌桩）', 31, 160), ('路基填筑', 100, 250),
            ('排水管道', 120, 260), ('给水及电力通信', 200, 320), ('路面结构（水稳＋沥青）', 280, 400),
            ('附属工程（路灯、交通、绿化）', 380, 440), ('收尾及竣工验收', 430, 450)],
         '关键节点：第30天临建完成｜第160天软基处理完成｜第260天排水贯通｜第320天管线完成｜第400天路面贯通｜第450天竣工验收'),
        ('38', date(2026, 12, 12), 300, [
            ('施工准备（临建、探测、交底）', 1, 20), ('管线探测与测量放线', 12, 45),
            ('路面破除与沟槽开挖', 30, 180), ('管道安装与接口', 45, 210),
            ('试压、冲洗与消毒', 120, 250), ('新旧管接驳与通水', 160, 265),
            ('路面恢复', 140, 285), ('附属工程及竣工验收', 250, 300)],
         '关键节点：第20天临建完成｜第45天探测交底完成｜第180天主干管沟槽交面｜第250天试压完成｜第265天全部通水｜第300天竣工验收'),
        ('39', date(2026, 12, 20), 270, [
            ('施工准备（分片对接、探测）', 1, 20), ('西周华翔工业园区改造', 21, 150),
            ('寿丰综合楼（东浦）改造', 40, 170), ('乌沙路新桥海丰村改造', 60, 190),
            ('灵南片区、泗洲头改造', 80, 210), ('新旧管接驳与通水', 150, 250),
            ('路面恢复', 130, 255), ('收尾及竣工验收', 230, 270)],
         '关键节点：第20天五片区全面开工｜第150天园区改造具备通水条件｜第210天各村网改造完成｜第250天接驳通水｜第270天竣工验收'),
    ]
    for tag, start, days, rows, note in jobs:
        print(gantt(f'gantt_out/{tag}_gantt.png', rows, start, days, note))

# -*- coding: utf-8 -*-
"""A3 四页版：无封面无目录、无附表、图片加框、小四宋体 1.5 倍行距"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x1F, 0x4E, 0x79)
SONG, HEI = '宋体', '黑体'

SEVEN = [
    ('（一）雨季及台风季施工', [
        '施工期跨越雨季与台风季，开工前疏通沿线排水系统，沟槽两侧设挡水土埂与排水沟，配备足量水泵与发电机；'
        '暴雨预警发布后停止土方与吊装作业，覆盖裸露边坡与材料堆场，加固围挡、临建及大型机械。'
        '台风来临前完成临建加固、桩机放倒或锚固、材料入库，人员按预案撤离至安全区域。']),
    ('（二）高温与夜间施工', [
        '高温时段调整作业时间，避开11:00至15:00露天作业，供应防暑降温用品；混凝土浇筑与养护采取降温与保湿措施。'
        '确需夜间施工的，提前办理夜间施工许可并公告周边，控制照明眩光与机械噪声。']),
    ('（三）应急处置', [
        '编制坍塌、触电、机械伤害、地下管线破坏、暴雨内涝等专项应急预案，成立应急小组并配备抢险物资与设备，'
        '每季度组织一次演练；发生事故按“先控制、后处置、再报告”的原则处理，一小时内报告建设单位与监管部门。']),
    ('（四）成品保护与竣工资料', [
        '已完工程分段挂牌保护，路面、井盖、管道口采取覆盖与封堵措施，严禁车辆碾压与杂物落入；'
        '竣工资料与工程实体同步形成，按分部分项整理检测报告、隐蔽验收记录与影像资料，竣工前完成组卷归档。']),
    ('（五）与建设单位、监理及管线单位的配合', [
        '主动接受建设单位与监理的检查指导，按时报送施工进度、材料报验与检测资料；开工前与供水、电力、通信、'
        '燃气等管线单位签订保护协议，明确联络人与应急响应流程，重要管线施工实行旁站监护。',
        '严格执行隐蔽工程验收与工序交接制度，对监理提出的问题在24小时内书面反馈整改情况，整改完成经复验合格后方可继续施工；'
        '主动配合建设单位组织的各项检查、验收与阶段性考核，确保工程按期、优质交付。']),
]


def _ea(run, name=SONG):
    rPr = run._element.get_or_add_rPr()
    rf = rPr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rPr.insert(0, rf)
    rf.set(qn('w:eastAsia'), name)


def add_border(inline, color='1F4E79', pt=1.0):
    """给图片加实线边框"""
    spPr = inline._inline.graphic.graphicData.pic.spPr
    ln = OxmlElement('a:ln')
    ln.set('w', str(int(pt * 12700)))
    for tag in ('a:solidFill',):
        fill = OxmlElement(tag)
        clr = OxmlElement('a:srgbClr')
        clr.set('val', color)
        fill.append(clr)
        ln.append(fill)
    dash = OxmlElement('a:prstDash')
    dash.set('val', 'solid')
    ln.append(dash)
    spPr.append(ln)


def build(P, tag, K):
    doc = Document()
    s = doc.sections[0]
    s.orientation = WD_ORIENT.LANDSCAPE
    s.page_width, s.page_height = Cm(42.0), Cm(29.7)
    s.top_margin, s.bottom_margin = Cm(1.2), Cm(1.0)
    s.left_margin, s.right_margin = Cm(1.4), Cm(1.2)

    st = doc.styles['Normal']
    st.font.name = 'Times New Roman'
    st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), SONG)
    st.paragraph_format.line_spacing = 1.5
    st.paragraph_format.space_before = Pt(0)
    st.paragraph_format.space_after = Pt(0)

    def H(text, size=15, before=8, after=4):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.keep_with_next = True
        p.paragraph_format.line_spacing = 1.5
        r = p.add_run(text)
        r.font.size = Pt(size)
        r.bold = True
        r.font.color.rgb = NAVY
        _ea(r, HEI)

    def Para(text):
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Pt(24)
        p.paragraph_format.line_spacing = 1.5
        r = p.add_run(text)
        r.font.size = Pt(12)
        _ea(r)

    def Figure(path, caption, width=None):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.0
        shape = p.add_run().add_picture(path, width=Cm(width or P.get('fig', 20.0)))
        add_border(shape)
        q = doc.add_paragraph()
        q.alignment = WD_ALIGN_PARAGRAPH.CENTER
        q.paragraph_format.line_spacing = 1.0
        q.paragraph_format.space_after = Pt(8)
        r = q.add_run(caption)
        r.font.size = Pt(10.5)
        _ea(r)

    def Table(headers, rows, widths, fsize=9.5, center=()):
        total = K.get('tw', 31.0)
        s_ = total / float(sum(widths))
        widths = [round(w * s_, 2) for w in widths]
        t = doc.add_table(rows=1, cols=len(headers))
        t.style = 'Table Grid'
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        lay = OxmlElement('w:tblLayout')
        lay.set(qn('w:type'), 'fixed')
        t._tbl.tblPr.append(lay)
        for i, h in enumerate(headers):
            c = t.rows[0].cells[i]
            c.text = ''
            tcPr = c._tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:fill'), '1F4E79')
            tcPr.append(shd)
            pr = c.paragraphs[0]
            pr.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pr.paragraph_format.line_spacing = 1.0
            r = pr.add_run(h)
            r.bold = True
            r.font.size = Pt(fsize)
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            _ea(r, HEI)
        for row in rows:
            cells = t.add_row().cells
            for i, v in enumerate(row):
                cells[i].text = ''
                pr = cells[i].paragraphs[0]
                pr.paragraph_format.line_spacing = 1.0
                pr.paragraph_format.first_line_indent = Pt(0)
                if i in center:
                    pr.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = pr.add_run(str(v))
                r.font.size = Pt(fsize)
                _ea(r)
        for i, w in enumerate(widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)
                tcPr = row.cells[i]._tc.get_or_add_tcPr()
                tcW = tcPr.find(qn('w:tcW'))
                if tcW is None:
                    tcW = OxmlElement('w:tcW')
                    tcPr.append(tcW)
                tcW.set(qn('w:w'), str(int(w * 567)))
                tcW.set(qn('w:type'), 'dxa')
        # 同步 tblGrid 与表宽，避免渲染时列错位
        grid = t._tbl.find(qn('w:tblGrid'))
        if grid is not None:
            cols = grid.findall(qn('w:gridCol'))
            for gc, w in zip(cols, widths):
                gc.set(qn('w:w'), str(int(w * 567)))
        tblW = t._tbl.tblPr.find(qn('w:tblW'))
        if tblW is None:
            tblW = OxmlElement('w:tblW')
            t._tbl.tblPr.append(tblW)
        tblW.set(qn('w:w'), str(int(sum(widths) * 567)))
        tblW.set(qn('w:type'), 'dxa')
        trPr = t.rows[0]._tr.get_or_add_trPr()
        trPr.append(OxmlElement('w:tblHeader'))
        for row in t.rows:
            row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
        doc.add_paragraph().paragraph_format.line_spacing = 1.0

    def PB():
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.0
        p.add_run().add_break(WD_BREAK.PAGE)

    brk = set(K.get('brk', []))

    # 一
    H('一、工程概况与特点分析', 15, before=0)
    for t in P['overview'][:K['ov']]:
        Para(t)
    H('（一）主要工程量')
    Para(P.get('qty_note', '本工程主要工程量（按招标工程量清单摘录）如下表。'))
    Table(P['qty_head'], P['qty_rows'][:K['qty']], P['qty_w'])
    H('（二）工程特点与难点')
    for i, t in enumerate(P['features'][:K['ft']], 1):
        Para(f'（{i}）{t}')
    H('（三）施工总体目标')
    Para(f"质量目标：一次性验收合格，分部分项工程合格率100%，无质量事故。工期目标：{P['days']}日历天内完成全部工程内容，按节点交付。"
         f"安全目标：无重大伤亡事故，无火灾、坍塌及地下管线破坏事故。文明目标：达到象山县文明工地标准，扬尘、噪声达标，污水不外排。")
    if 1 in brk:
        PB()

    # 二
    H('二、施工总体部署', 15, before=10)
    Para(P['org'])
    H('（一）施工段划分与流水组织')
    for t in P['segments'][:K['sg']]:
        Para(t)
    H('（二）总体施工流程')
    Para(P['flow'])
    Figure(f"{P.get('fig_dir', 'figp3')}/{tag}_plan.png", '图 1  施工总平面布置示意图')
    if 2 in brk:
        PB()

    # 三
    H('三、主要分部分项工程施工方案', 15, before=10)
    for title, paras in P['methods'][:K['m1']]:
        H(title, 13, before=6)
        for t in paras[:K['pp']]:
            Para(t)
    Figure(f"{P.get('fig_dir', 'figp3')}/{tag}_flow.png", '图 3  主要工序施工工艺流程图')
    for title, paras in P['methods2'][:K['m2']]:
        H(title, 13, before=6)
        for t in paras[:K['pp']]:
            Para(t)
    H('（十一）主要施工机械与劳动力配置', 13, before=8)
    eq = P['equip']
    names = '、'.join(f"{e[1]}{e[2]}×{e[3]}台" for e in eq[:8])
    Para(f'主要施工机械按施工段与工序配套配置：{names}，并配备发电机、水泵等辅助设备，进场前完成检验与备案，'
         '关键设备一用一备，确保工序不因设备故障中断。')
    Para('劳动力按基础、主体、面层、附属四个阶段组织，主要工种包括机械操作工、管道安装工、混凝土工、钢筋工、'
         '电工焊工及普工，高峰期约' + str(P.get('labor_peak', 110)) + '人，全部经三级教育后上岗，特种作业人员持证。')
    if 3 in brk:
        PB()

    # 四
    H('四、质量保证措施', 15, before=10)
    for i, t in enumerate(P['quality'][:K['qu']], 1):
        Para(f'（{i}）{t}')
    H('（七）关键工序质量控制点')
    Table(['工序', '控制要点', '检验方法'], P['qc_points'][:5], [4.6, 9.4, 4.6])
    if 4 in brk:
        PB()

    # 五
    H('五、安全生产与文明施工措施', 15, before=10)
    H('（一）安全生产措施', 13, before=6)
    for i, t in enumerate(P['safety'][:K['sf']], 1):
        Para(f'（{i}）{t}')
    H('（二）文明施工与环境保护措施', 13, before=6)
    for i, t in enumerate(P['civil'][:K['cv']], 1):
        Para(f'（{i}）{t}')
    if 5 in brk:
        PB()

    # 六
    H('六、进度保证与交通组织措施', 15, before=10)
    H('（一）进度控制措施', 13, before=6)
    for i, t in enumerate(P['progress'][:K['pg']], 1):
        Para(f'（{i}）{t}')
    H('（二）进度计划的编制与动态管理', 13, before=6)
    Para('以总进度计划为纲，按“总计划—月计划—周计划—日安排”四级分解，采用网络计划技术分级管理：'
         '开工前编制总进度计划报监理审批，每月25日前报送下月计划，每周生产例会对照检查上周完成情况，'
         '日调度会明确次日作业面与资源投入。')
    Para(f"计划编制以{P['days']}日历天总工期为约束，倒排关键节点：开工、管线交面、主体完成、路面（管道）贯通、"
         '竣前验收，每个节点设置提前量，节点滞后3天以上即启动预警，滞后7天以上启动纠偏方案。')
    Para('纠偏措施按“先内部调剂、后增加投入”的顺序实施：优先在施工段之间调剂人员机械，其次延长作业时间或增加作业面，'
         '必要时报建设单位批准后调整工序搭接关系；每月对进度执行情况进行考核，考核结果与班组结算挂钩。')
    Para('建立进度信息台账，逐日记录完成的工程量、投入的人员机械与影响因素；因设计变更、管线迁改、恶劣天气等'
         '非承包人原因造成的工期影响，及时办理签证与工期顺延手续，做到有据可查。')
    H('（三）交通组织与管线保护措施', 13, before=6)
    for i, t in enumerate(P['traffic'][:K['tf']], 1):
        Para(f'（{i}）{t}')
    if 6 in brk:
        PB()

    # 七
    H('七、季节性施工与应急保障措施', 15, before=10)
    for title, body in SEVEN[:K.get('s7', 5)]:
        H(title, 13, before=6)
        for b in body:
            Para(b)

    out = P.get('out') or f"{tag}-技术标（施工组织设计·A3四页）.docx"
    doc.save(out)
    return out

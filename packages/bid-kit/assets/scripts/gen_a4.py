# -*- coding: utf-8 -*-
"""三份技术标 v3：小四宋体 1.5 倍行距、无附表、三张图、连续排版凑足 10 页"""
import sys
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x1F, 0x4E, 0x79)
SONG, HEI = '宋体', '黑体'


def _ea(run, name=SONG):
    rPr = run._element.get_or_add_rPr()
    rf = rPr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rPr.insert(0, rf)
    rf.set(qn('w:eastAsia'), name)


def build(P, tag, K):
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.top_margin, s.bottom_margin = Cm(2.2), Cm(2.0)
    s.left_margin, s.right_margin = Cm(2.4), Cm(2.2)

    st = doc.styles['Normal']
    st.font.name = 'Times New Roman'
    st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), SONG)
    st.paragraph_format.line_spacing = 1.5
    st.paragraph_format.space_before = Pt(0)
    st.paragraph_format.space_after = Pt(0)

    def H(text, size=14, before=10, after=4):
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
        p.add_run().add_picture(path, width=Cm(width or P.get('fig', 16.2)))
        q = doc.add_paragraph()
        q.alignment = WD_ALIGN_PARAGRAPH.CENTER
        q.paragraph_format.line_spacing = 1.0
        q.paragraph_format.space_after = Pt(8)
        r = q.add_run(caption)
        r.font.size = Pt(10.5)
        _ea(r)

    def Table(headers, rows, widths, fsize=9.5, center=()):
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
        # 表头跨页重复、行内不拆分
        trPr = t.rows[0]._tr.get_or_add_trPr()
        th = OxmlElement('w:tblHeader')
        trPr.append(th)
        for row in t.rows:
            trPr2 = row._tr.get_or_add_trPr()
            cs = OxmlElement('w:cantSplit')
            trPr2.append(cs)
        doc.add_paragraph().paragraph_format.line_spacing = 1.0

    def PB():
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.0
        p.add_run().add_break(WD_BREAK.PAGE)

    A3 = bool(K.get('a3'))
    BRK = set(K.get('brk', []))

    # 封面
    if not A3:
        for _ in range(4):
            doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(P['name'])
    r.font.size = Pt(22)
    r.bold = True
    _ea(r, HEI)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('（' + P['sub'] + '）')
    r.font.size = Pt(14)
    _ea(r)
    for _ in range(3):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('技  术  标')
    r.font.size = Pt(30)
    r.bold = True
    r.font.color.rgb = NAVY
    _ea(r, HEI)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('（施工组织设计）')
    r.font.size = Pt(16)
    _ea(r)
    for _ in range(5):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('二〇二六年')
    r.font.size = Pt(14)
    _ea(r)
    PB()

    # 目录
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run('目　　录')
    r.font.size = Pt(18)
    r.bold = True
    r.font.color.rgb = NAVY
    _ea(r, HEI)
    TOC = [('一、工程概况与特点分析', ''), ('二、施工总体部署', ''), ('三、主要分部分项工程施工方案', ''),
           ('四、质量保证措施', ''), ('五、安全生产与文明施工措施', ''), ('六、进度保证与交通组织措施', ''),
           ('七、季节性施工与应急保障措施', '')]
    tocpg = K.get('toc') or {}
    for name, _ in TOC:
        q = doc.add_paragraph()
        q.paragraph_format.line_spacing = 1.5
        q.paragraph_format.space_after = Pt(2)
        r = q.add_run(name)
        r.font.size = Pt(12)
        _ea(r)
        if name in tocpg:
            r2 = q.add_run('　' + '…' * 12 + '　' + str(tocpg[name]))
            r2.font.size = Pt(12)
            _ea(r2)
    q = doc.add_paragraph()
    q.paragraph_format.line_spacing = 1.5
    q.paragraph_format.space_before = Pt(14)
    r = q.add_run('附图：图 1 施工总平面布置示意图　图 2 施工进度计划横道图　图 3 主要工序施工工艺流程图')
    r.font.size = Pt(10.5)
    _ea(r)
    q = doc.add_paragraph()
    q.paragraph_format.line_spacing = 1.5
    q.paragraph_format.space_before = Pt(10)
    r = q.add_run('编制依据：本工程招标文件、招标工程量清单、施工图纸及现行国家、行业施工质量验收规范与安全技术规程。')
    r.font.size = Pt(10.5)
    _ea(r)
    PB()

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

    # 二
    H('二、施工总体部署', 15, before=12)
    Para(P['org'])
    H('（一）施工段划分与流水组织')
    for t in P['segments'][:K['sg']]:
        Para(t)
    H('（二）总体施工流程')
    Para(P['flow'])
    Figure(f"{P.get('fig_dir', 'figp3')}/{tag}_plan.png", '图 1  施工总平面布置示意图')

    # 三
    H('三、主要分部分项工程施工方案', 15, before=12)
    for title, paras in P['methods'][:K['m1']]:
        H(title, 13, before=8)
        for t in paras[:K['pp']]:
            Para(t)
    Figure(f"{P.get('fig_dir', 'figp3')}/{tag}_flow.png", '图 3  主要工序施工工艺流程图')
    for title, paras in P['methods2'][:K['m2']]:
        H(title, 13, before=8)
        for t in paras[:K['pp']]:
            Para(t)
    if P.get('method_table'):
        Table(P['method_table'][0], P['method_table'][1][:5], P['method_table'][2])
    H('（十一）主要施工机械与劳动力配置', 13, before=8)
    eq = P['equip']
    names = '、'.join(f"{e[1]}{e[2]}×{e[3]}台" for e in eq[:8])
    Para(f'主要施工机械按施工段与工序配套配置：{names}，并配备发电机、水泵等辅助设备，进场前完成检验与备案，'
         '关键设备一用一备，确保工序不因设备故障中断。')
    Para('劳动力按基础、主体、面层、附属四个阶段组织，主要工种包括机械操作工、管道安装工、混凝土工、钢筋工、'
         '电工焊工及普工，高峰期约' + str(P.get('labor_peak', 110)) + '人，全部经三级教育后上岗，特种作业人员持证。')

    # 四
    H('四、质量保证措施', 15, before=12)
    for i, t in enumerate(P['quality'][:K['qu']], 1):
        Para(f'（{i}）{t}')
    H('（七）关键工序质量控制点')
    Table(['工序', '控制要点', '检验方法'], P['qc_points'][:5], [4.2, 8.2, 4.0])

    # 五
    H('五、安全生产与文明施工措施', 15, before=12)
    H('（一）安全生产措施', 13, before=6)
    for i, t in enumerate(P['safety'][:K['sf']], 1):
        Para(f'（{i}）{t}')
    H('（二）文明施工与环境保护措施', 13, before=6)
    for i, t in enumerate(P['civil'][:K['cv']], 1):
        Para(f'（{i}）{t}')

    # 六
    H('六、进度保证与交通组织措施', 15, before=12)
    for i, t in enumerate(P['progress'][:K['pg']], 1):
        Para(f'（{i}）{t}')
    Figure(f"{P.get('fig_dir', 'figp3')}/{tag}_gantt.png", '图 2  施工进度计划横道图')
    H('（二）交通组织与管线保护措施', 13, before=6)
    for i, t in enumerate(P['traffic'][:K['tf']], 1):
        Para(f'（{i}）{t}')

    # 七
    H('七、季节性施工与应急保障措施', 15, before=12)
    H('（一）雨季及台风季施工', 13, before=6)
    Para('施工期跨越雨季与台风季，开工前疏通沿线排水系统，沟槽两侧设挡水土埂与排水沟，配备足量水泵与发电机；'
         '暴雨预警发布后停止土方与吊装作业，覆盖裸露边坡与材料堆场，加固围挡、临建及大型机械。'
         '台风来临前完成临建加固、塔吊与桩机放倒或锚固、材料入库，人员按预案撤离至安全区域。')
    H('（二）高温与夜间施工', 13, before=6)
    Para('高温时段调整作业时间，避开11:00至15:00露天作业，供应防暑降温用品；混凝土浇筑与养护采取降温与保湿措施。'
         '确需夜间施工的，提前办理夜间施工许可并公告周边，控制照明眩光与机械噪声。')
    H('（三）应急处置', 13, before=6)
    Para('编制坍塌、触电、机械伤害、地下管线破坏、暴雨内涝等专项应急预案，成立应急小组并配备抢险物资与设备，'
         '每季度组织一次演练；发生事故按“先控制、后处置、再报告”的原则处理，一小时内报告建设单位与监管部门。')
    H('（四）成品保护与竣工资料', 13, before=6)
    Para('已完工程分段挂牌保护，路面、井盖、管道口采取覆盖与封堵措施，严禁车辆碾压与杂物落入；'
         '竣工资料与工程实体同步形成，按分部分项整理检测报告、隐蔽验收记录与影像资料，竣工前完成组卷归档。')
    H('（五）与建设单位、监理及管线单位的配合', 13, before=6)
    Para('主动接受建设单位与监理的检查指导，按时报送施工进度、材料报验与检测资料；'
         '开工前与供水、电力、通信、燃气等管线单位签订保护协议，明确联络人与应急响应流程，重要管线施工实行旁站监护。')
    Para('严格执行隐蔽工程验收与工序交接制度，对监理提出的问题在24小时内书面反馈整改情况，整改完成经复验合格后方可继续施工；'
         '主动配合建设单位组织的各项检查、验收与阶段性考核，确保工程按期、优质交付。')

    out = P.get('out_a4') or f"{tag}-技术标（施工组织设计·A4十页）.docx"
    doc.save(out)
    return out

# -*- coding: utf-8 -*-
"""附表生成：在模版上替换 4 张表 + 2 张图（结构与页数不变）"""
import os
import shutil
from docx import Document
from docx.shared import Cm
from docx.text.paragraph import Paragraph

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(os.path.dirname(HERE), 'templates', '附表模板.docx')


def _write_row(row, vals):
    cells = row.cells
    for j, v in enumerate(vals):
        if j >= len(cells):
            break
        cp = cells[j].paragraphs[0]
        if cp.runs:
            cp.runs[0].text = str(v)
            for extra in cp.runs[1:]:
                extra.text = ''
        else:
            cp.add_run(str(v))


def _find_gantt_paragraph(body):
    """附表三（进度图）所在段落：表4 标题之前、含图片的空段落"""
    idx4 = None
    for i, el in enumerate(body):
        txt = ''.join(t.text or '' for t in el.iter(W + 't'))
        if txt.strip().startswith('表4'):
            idx4 = i
            break
    limit = idx4 if idx4 is not None else len(body)
    for i in range(limit - 1, -1, -1):
        el = body[i]
        if el.tag.endswith('}p') and el.findall('.//' + A + 'blip'):
            return el
    raise RuntimeError('未找到附表三的图片段落')


def build(out_path, equip, labor, land, gantt_png, site_png,
          gantt_width_cm=19.84, site_width_cm=23.0, template=TEMPLATE):
    shutil.copy(template, out_path)
    doc = Document(out_path)
    body = list(doc.element.body)

    # 附表三：进度横道图
    p_el = _find_gantt_paragraph(body)
    for run in list(p_el.findall(W + 'r')):
        if run.findall('.//' + A + 'blip'):
            p_el.remove(run)
    para = Paragraph(p_el, doc)
    para.alignment = 1
    para.add_run().add_picture(gantt_png, width=Cm(gantt_width_cm))

    # 附表四：总平面图（1×1 表格内的图）
    cell = doc.tables[2].rows[0].cells[0]
    for p_ in cell.paragraphs:
        for run in list(p_._p.findall(W + 'r')):
            if run.findall('.//' + A + 'blip'):
                p_._p.remove(run)
    for el in list(cell._tc.iter(W + 'drawing')):
        el.getparent().remove(el)
    p0 = cell.paragraphs[0]
    p0.alignment = 1
    p0.add_run().add_picture(site_png, width=Cm(site_width_cm))

    # 附表一 机械设备、附表二 劳动力、临时用地表
    for i, vals in enumerate(equip, start=1):
        if i < len(doc.tables[0].rows):
            _write_row(doc.tables[0].rows[i], vals)
    for i, vals in enumerate(labor, start=2):
        if i < len(doc.tables[1].rows):
            _write_row(doc.tables[1].rows[i], vals)
    for i, vals in enumerate(land, start=1):
        if i < len(doc.tables[3].rows):
            _write_row(doc.tables[3].rows[i], vals)

    doc.save(out_path)
    return out_path

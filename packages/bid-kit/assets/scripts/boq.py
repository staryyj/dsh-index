# -*- coding: utf-8 -*-
"""清单解析：xlsx 读取 + 关键工程量分类汇总"""
import re
import zipfile
from xml.etree import ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
PKG = '{http://schemas.openxmlformats.org/package/2006/relationships}'
UNITS = ('m2', 'm²', 'm3', 'm³', 'm', 't', 'kg', '个', '座', '套', '根', '处', '株', '棵', '组',
         '台', '项', '块', '只', '副', '段', 'km')
KEEP_SHEET = ('分部分项', '道路工程', '排水工程', '交通工程', '路灯工程', '绿化工程', '智能交通',
              '再生水工程', '给水工程', '通信工程', '附属工程', '人行道工程')


def sheets(path):
    z = zipfile.ZipFile(path)
    shared = []
    if 'xl/sharedStrings.xml' in z.namelist():
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root.findall(NS + 'si'):
            shared.append(''.join(t.text or '' for t in si.iter(NS + 't')))
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    relmap = {r.get('Id'): r.get('Target') for r in rels.findall(PKG + 'Relationship')}
    out = []
    for sh in wb.find(NS + 'sheets').findall(NS + 'sheet'):
        name = sh.get('name')
        tgt = relmap.get(sh.get(RNS + 'id'), '')
        cand = ['xl/' + tgt.lstrip('/').replace('xl/', ''), 'xl/worksheets/' + tgt.split('/')[-1]]
        p = next((x for x in cand if x in z.namelist()), None)
        if not p:
            continue
        root = ET.fromstring(z.read(p))
        rows = []
        for r in root.findall('.//' + NS + 'row'):
            vals = []
            for c in r.findall(NS + 'c'):
                t = c.get('t')
                v = c.find(NS + 'v')
                if t == 's' and v is not None:
                    try:
                        vals.append(shared[int(v.text)])
                    except Exception:
                        vals.append('')
                elif t == 'inlineStr':
                    is_ = c.find(NS + 'is')
                    vals.append(''.join(x.text or '' for x in is_.iter(NS + 't')) if is_ is not None else '')
                else:
                    vals.append(v.text if v is not None else '')
            rows.append(vals)
        out.append((name, rows))
    return out


def num(s):
    try:
        return float(str(s).replace(',', '').strip())
    except Exception:
        return None


def parse_items(path):
    """→ [(名称, 特征, 单位, 数量)]"""
    rows = []
    for name, rs in sheets(path):
        if not (('分部分项' in name) or (name in KEEP_SHEET)):
            continue
        for r in rs:
            vals = [str(x).strip() for x in r]
            code_i = next((i for i, v in enumerate(vals)
                           if re.fullmatch(r'\d{9,12}|04B\d{3}', v)), None)
            if code_i is None:
                continue
            nm = vals[code_i + 1] if code_i + 1 < len(vals) else ''
            unit = qty = None
            for i in range(code_i + 2, len(vals)):
                if vals[i] in UNITS:
                    unit = vals[i]
                    for j in range(i + 1, min(i + 4, len(vals))):
                        q = num(vals[j])
                        if q is not None:
                            qty = q
                            break
                    break
            feat = ' '.join(vals[code_i + 2:code_i + 8])
            if qty is not None:
                rows.append((nm, feat, unit, qty))
    return rows


CATS = [
    ('沥青混凝土面层', lambda nm, f: '沥青混凝土' in nm, ('m2', 'm²'), ('拆除',)),
    ('水泥稳定碎石基层', lambda nm, f: '水泥稳定' in nm or '水稳' in nm, ('m2', 'm²'), ('拆除',)),
    ('级配碎石底基层', lambda nm, f: '碎石' in nm and '稳定' not in nm, ('m2', 'm²'), ('拆除',)),
    ('水泥混凝土路面', lambda nm, f: '水泥混凝土' in nm, ('m2', 'm²'), ('拆除',)),
    ('旧板病害处理', lambda nm, f: '板块处理' in nm or '病害' in nm, ('m2', 'm²', 'm', 'm3', 'm³'), ()),
    ('植筋', lambda nm, f: '植筋' in nm, ('根',), ()),
    ('防裂贴', lambda nm, f: '防裂贴' in nm, ('m2', 'm²'), ()),
    ('应力吸收层', lambda nm, f: '应力吸收' in nm, ('m2', 'm²'), ()),
    ('拆除路面', lambda nm, f: '拆除路面' in nm, ('m2', 'm²'), ()),
    ('拆除基层', lambda nm, f: '拆除基层' in nm, ('m2', 'm²'), ()),
    ('铣刨路面', lambda nm, f: '铣刨' in nm, ('m2', 'm²'), ()),
    ('人行道铺装', lambda nm, f: '人行道块料铺设' in nm or '人行道板' in nm, ('m2', 'm²'), ('拆除',)),
    ('拆除人行道', lambda nm, f: '拆除人行道' in nm, ('m2', 'm²'), ()),
    ('侧平石', lambda nm, f: any(k in nm for k in ('侧平石', '侧石', '平石', '缘石')), ('m',), ('拆除',)),
    ('拆除侧平石', lambda nm, f: '拆除侧' in nm, ('m',), ()),
    ('雨水管', lambda nm, f: any(k in nm for k in ('塑料管', '混凝土管', '波纹管', '缠绕管')) and '雨水' in f,
     None, ()),
    ('污水管', lambda nm, f: any(k in nm for k in ('塑料管', '混凝土管', '波纹管', '缠绕管')) and '污水' in f,
     None, ()),
    ('给水管', lambda nm, f: ('球墨' in nm or '给水' in nm) and '给水' in (nm + f), None, ('拆除',)),
    ('再生水管', lambda nm, f: '再生水' in (nm + f), None, ('拆除',)),
    ('检查井', lambda nm, f: '检查井' in nm and '井盖' not in nm, ('座', '个'), ('加高',)),
    ('检查井加高', lambda nm, f: '检查井加高' in nm or ('加高' in nm and '井' in nm), ('座',), ()),
    ('雨水口', lambda nm, f: '雨水口' in nm, ('座', '个'), ('拆除',)),
    ('通信排管', lambda nm, f: ('排管' in nm or '通信' in nm) and ('通信' in (nm + f) or 'CPVC' in f), ('m',), ()),
    ('路灯灯具', lambda nm, f: ('照明灯' in nm or '路灯' in nm or '灯杆' in nm), ('套', '个', '根'), ('拆除', '信号')),
    ('绿化', lambda nm, f: any(k in nm for k in ('乔木', '灌木', '苗木')), ('株', '棵', 'm2'), ()),
]


def summarize(items):
    """→ {类别: {单位: 合计}}"""
    out = {}
    for label, pred, units, exclude in CATS:
        acc = {}
        for nm, feat, u, q in items:
            if units and u not in units:
                continue
            if not pred(nm, feat):
                continue
            if exclude and any(k in (nm + feat) for k in exclude):
                continue
            acc[u or '项'] = acc.get(u or '项', 0.0) + q
        if acc:
            out[label] = acc
    return out


def top_level(path):
    """清单总造价（若表内有招标控制价费用表）"""
    for name, rows in sheets(path):
        for r in rows:
            vals = [str(x) for x in r]
            if any('合计' in v for v in vals):
                for v in vals:
                    n = num(v)
                    if n and n > 10000:
                        return n
    return None


if __name__ == '__main__':
    import sys, json
    p = sys.argv[1]
    items = parse_items(p)
    summary = summarize(items)
    print(json.dumps({'items': len(items), 'summary': summary}, ensure_ascii=False, indent=2))

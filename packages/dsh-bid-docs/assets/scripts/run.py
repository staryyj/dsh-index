# -*- coding: utf-8 -*-
"""投标方案一键生成：inspect（读招标资料）/ build（出技术标+附表）

用法：
  python run.py inspect --archive 项目包.zip --work work/
  python run.py build   --work work/

inspect  解压招标资料 → 识别最新版招标文件与清单 → 抽取关键参数与清单量
build    读 work/plan.json（可先由人/AI 调整）→ 出图 → 生成技术标 A3 与附表 → 转 PDF 报页数
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import apx              # noqa: E402
import boq              # noqa: E402
import figs             # noqa: E402
import gen_a3           # noqa: E402
import gantt            # noqa: E402
import prose            # noqa: E402

NODE = os.environ.get('BIDDOC_NODE') or os.environ.get('DSH_NODE_EXECUTABLE')
LO_CLI = os.environ.get('BIDDOC_LO_CLI') or os.environ.get(
    'DSH_LIBREOFFICE_CLI',
    r'D:\deepseek\resources\app.asar.unpacked\dsh\node_modules\@deepseek-ai\libreoffice-kit\lib\cli.js')


# ---------------------------------------------------------------- 归档
def _unpack(archive, dest):
    os.makedirs(dest, exist_ok=True)
    if archive.lower().endswith('.zip'):
        with zipfile.ZipFile(archive) as z:
            z.extractall(dest)
    else:
        raise SystemExit('目前只支持 .zip 项目包；rar 请先解压后传目录')
    return dest


def _walk(root):
    for base, _, files in os.walk(root):
        for f in files:
            yield os.path.join(base, f)


def _tender_pdf(root):
    """取最新版招标文件：优先含 V版本号最大者，其次 补遗 目录 > 根目录"""
    cands = []
    for p in _walk(root):
        low = os.path.basename(p).lower()
        if not low.endswith('.pdf'):
            continue
        if '招标文件' not in os.path.basename(p):
            continue
        m = re.search(r'[Vv](\d+)', os.path.basename(p))
        ver = int(m.group(1)) if m else 0
        bubu = 1 if ('补遗' in p or '澄清' in p) else 0
        cands.append((ver, bubu, p))
    if not cands:
        return None
    cands.sort()
    return cands[-1][2]


def _boq_xlsx(root):
    """清单工作簿识别：按工作表名判断（分部分项/道路工程/排水工程…），排除通讯与计算规则"""
    best = None
    for p in _walk(root):
        if not p.lower().endswith('.xlsx'):
            continue
        name = os.path.basename(p)
        if any(k in name for k in ('通信', '计算规则', '费用表')):
            continue
        score = 0
        if '清单' in name:
            score += 2
        if '控制价' in name:
            score += 2
        try:
            names = [n for n, _ in boq.sheets(p)]
        except Exception:
            continue
        has = any(('分部分项' in n) or (n in boq.KEEP_SHEET) for n in names)
        if not has:
            continue
        score += 3
        if any('道路工程' in n for n in names):
            score += 2
        if best is None or score > best[0]:
            best = (score, p)
    return best[1] if best else None


# ---------------------------------------------------------------- 招标参数
def _dump_pdf_text(pdf, out_txt):
    import pdfx
    import pdftext
    doc = pdfx.PDF(pdf)
    nums = pdfx.page_numbers(doc)
    cache = {}
    chunks = []
    for i, num in enumerate(nums, 1):
        h, _ = doc.objs.get(num, (b'', None))
        fs = pdftext.parse_resources_fonts(h) if h else {}
        for n, fn in fs.items():
            if fn not in cache:
                cache[fn] = pdftext.font_decoder(doc, fn)
        try:
            t = pdftext.extract_page(doc, num, cache) if fs else ''
        except Exception:
            t = ''
        chunks.append(f'\n<<<PAGE {i}>>>\n{t}')
    txt = ''.join(chunks)
    with open(out_txt, 'w', encoding='utf-8') as f:
        f.write(txt)
    return txt, len(nums)


def extract_params(txt, hint_name=''):
    """从招标文件文本抽关键参数。PDF 文本可能乱序，采用"多模式 + 候选"策略。"""
    flat = re.sub(r'\s+', '', txt)
    p = {}

    # 名称：优先取含"标段"的最长工程名；否则用文件名 hint
    hint = (hint_name or '').strip()
    if hint and ('工程' in hint or '标段' in hint) and len(hint) >= 12:
        p['name'] = hint
    else:
        cands = re.findall(r'[\u4e00-\u9fa5（）()A-Za-z0-9、·]{8,120}?标段（[^）]{2,80}）', flat)
        bad = ('工程名称', '投标人', '序号', '0000', '/', '招标', '资格要求', '工程量',
               '单位盖章', '有限公司', '施工项目', '联系电话', '设计单位')
        cands = [c for c in cands if c.endswith('）') and '工程' in c
                 and not any(b in c for b in bad) and 16 <= len(c) <= 100]
        p['name'] = max(cands, key=len) if cands else ''
        if not p['name']:
            p['name_hint'] = '未能自动识别项目全称，请用 --name 传入或用页面图核对'

    # 工期：优先紧跟"日历天"的数字，其次全文 3 位数候选取最常见
    m = re.findall(r'(\d{2,4})\s*日历天', flat)
    if not m:
        m = re.findall(r'(\d{3})\s*日历天', flat)
    if m:
        from collections import Counter
        p['days'] = int(Counter(m).most_common(1)[0][0])
    else:
        nums = [int(x) for x in re.findall(r'(?<!\d)(\d{3})(?!\d)', flat)]
        nums = [x for x in nums if 60 <= x <= 999]
        from collections import Counter as _C
        p['days_candidates'] = [v for v, _ in _C(nums).most_common(6)]

    # 本次招标建安工程造价（金额形如 2028.087100）
    cands = re.findall(r'(\d{3,5}\.\d{4,6})', flat)
    if cands:
        from collections import Counter
        p['money_candidates'] = [c for c, _ in Counter(cands).most_common(4)]

    # 长度
    lengths = [int(x) for x in re.findall(r'(?<!\d)(\d{3,5})(?!\d)\s*米', flat)]
    lengths = [x for x in lengths if 100 <= x <= 20000]
    if lengths:
        from collections import Counter
        p['length'] = Counter(lengths).most_common(1)[0][0]
        p['length_candidates'] = sorted(set(lengths))[:6]

    # 宽度
    m = re.search(r'(\d{2})\s*[-－~～]\s*(\d{2})\s*米', flat)
    if m:
        p['width_text'] = f'{m.group(1)}～{m.group(2)}'

    # 保证金（大写）
    m = re.search(r'([一二三四五六七八九十百零壹贰叁肆伍陆柒捌玖拾]{2,8}万元整)', flat)
    if m:
        p['bond'] = m.group(1)

    # 资质 / 建造师 / 技术标页数 / 暗标 / 政府采购
    m = re.search(r'市政公用工程施工总承包(叁|贰|壹)级', flat)
    if m:
        p['qual'] = '市政' + m.group(1) + '级'
    m = re.search(r'(壹|贰)级.{0,4}建造师', flat)
    if m:
        p['pm'] = m.group(1) + '级建造师'
    m = re.search(r'技术标文件.{0,24}?(\d{1,2})页', flat) or re.search(r'(\d{1,2})\s*页', flat)
    if m:
        p['tech_pages'] = int(m.group(1))
    p['anbiao'] = '暗标' in flat
    return p


def render_key_pages(pdf, work, pages=(1, 7, 8, 9)):
    """把公告/前附表等关键页渲染成 PNG，供人工或视觉核对（PDF 文本可能乱序）"""
    out_dir = os.path.join(work, 'pages')
    if os.path.isdir(out_dir):
        shutil.rmtree(out_dir, ignore_errors=True)
    page_arg = ','.join(str(x) for x in pages)
    try:
        if NODE and os.path.exists(LO_CLI):
            subprocess.run([NODE, LO_CLI, 'render', '--input', pdf, '--output-dir', out_dir,
                            '--pages', page_arg, '--dpi', '110'], capture_output=True, timeout=300)
        else:
            subprocess.run(['pdftoppm', '-png', '-r', '110', '-f', str(min(pages)),
                            '-l', str(max(pages)), pdf, os.path.join(out_dir, 'page')],
                           capture_output=True, timeout=300)
    except Exception:
        return []
    return sorted(os.path.join(out_dir, f) for f in os.listdir(out_dir) if f.endswith('.png'))


# ---------------------------------------------------------------- inspect
def cmd_inspect(args):
    work = os.path.abspath(args.work)
    src = os.path.abspath(args.archive)
    raw = os.path.join(work, 'raw')
    for d in (work, raw):
        os.makedirs(d, exist_ok=True)
    if os.path.isdir(src):
        src_real, work_real = os.path.realpath(src), os.path.realpath(work)
        if work_real.startswith(src_real + os.sep) or work_real == src_real:
            raw = src                    # 工作目录就在资料目录里，直接读取，避免自我复制
        else:
            shutil.copytree(src, raw, dirs_exist_ok=True)
    else:
        _unpack(src, raw)

    hint_name = args.name or os.path.basename(src).replace('.zip', '')
    tender = _tender_pdf(raw)
    params = {}
    pages = 0
    if tender:
        txt, pages = _dump_pdf_text(tender, os.path.join(work, 'tender.txt'))
        hint = args.name or os.path.basename(src).replace('.zip', '')
        params = extract_params(txt, hint_name=hint)
    params.setdefault('name', hint_name)
    params['tender_file'] = os.path.basename(tender) if tender else None
    params['tender_pages'] = pages

    wb = _boq_xlsx(raw)
    summary = {}
    if wb:
        try:
            summary = boq.summarize(boq.parse_items(wb))
        except Exception as e:
            params['boq_error'] = str(e)
    params['boq_file'] = os.path.basename(wb) if wb else None

    pages_dir = render_key_pages(tender, work) if tender else []
    params['page_images'] = pages_dir
    json.dump(params, open(os.path.join(work, 'params.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=2)
    json.dump(summary, open(os.path.join(work, 'boq.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=2)

    lines = ['# 招标资料核对表', '', '> PDF 抽文本可能乱序，工期/造价/长度等数字请对照 pages/ 下页面图核对；候选值见 money_candidates；项目全称可用 --name 传入。', '', f'- 招标文件：{params.get("tender_file")}（{pages} 页）',
             f'- 清单工作簿：{params.get("boq_file")}', '']
    lines += [f'- {k}：{v}' for k, v in params.items() if k not in ('tender_file', 'boq_file', 'tender_pages')]
    if summary:
        lines += ['', '## 清单主要工程量']
        for k, v in summary.items():
            lines.append(f'- {k}：' + '，'.join(f'{u} {n:,.2f}' for u, n in v.items()))
    open(os.path.join(work, 'params.md'), 'w', encoding='utf-8').write('\n'.join(lines))
    print('\n'.join(lines))
    print(f'\n[输出] {work}\\params.json, boq.json, tender.txt, params.md')
    return 0


# ---------------------------------------------------------------- build
def _draft_plan(work):
    params = json.load(open(os.path.join(work, 'params.json'), encoding='utf-8'))
    boq_summary = json.load(open(os.path.join(work, 'boq.json'), encoding='utf-8')) if \
        os.path.exists(os.path.join(work, 'boq.json')) else {}
    name = params.get('name') or '（项目名称）'
    return {
        'tag': '00',
        'short': '项目',
        'name': name,
        'road': name,
        'length': params.get('length'),
        'width_text': params.get('width_text', '—'),
        'days': params.get('days', 360),
        'money': params.get('money', '—'),
        'kind': 'road',
        'boq': boq_summary,
    }


def cmd_build(args):
    work = os.path.abspath(args.work)
    plan_path = os.path.join(work, 'plan.json')
    if not os.path.exists(plan_path):
        plan = _draft_plan(work)
        json.dump(plan, open(plan_path, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
        print(f'[提示] 未找到 plan.json，已按招标资料生成草稿：{plan_path}')
        print('       请补齐 name / road / length / width_text / days / money / kind / struct 后重跑 build；')
        print('       工期、造价、长度请对照 work/pages/ 下的页面图核对（PDF 抽文本可能乱序）。')
        return 0
    plan = json.load(open(plan_path, encoding='utf-8'))
    tag = str(plan.get('tag') or '00')
    out_dir = os.path.join(work, 'out')
    fig_dir = os.path.join(work, 'figs')
    os.makedirs(out_dir, exist_ok=True)
    os.makedirs(fig_dir, exist_ok=True)

    P = prose.build_P(plan)
    # 图
    road_label = plan.get('road') or plan.get('name')
    figs.plan_road(os.path.join(fig_dir, f'{tag}_plan.png'), road_label,
                   plan.get('segments') or ['第一施工段', '第二施工段', '第三施工段'],
                   plan.get('facilities') or ['项目部\n办公区', '材料堆场\n管材堆放', '稳定土\n拌合站', '洗车池\n临时便道'])
    figs.flow(os.path.join(fig_dir, f'{tag}_flow.png'), P['flow_steps'])
    start = plan.get('start_date') or '2026-12-20'
    y, m, d = [int(x) for x in str(start).split('-')]
    gantt.gantt(os.path.join(fig_dir, f'{tag}_gantt.png'), P['gantt'], date(y, m, d), int(P['days']),
                note=plan.get('gantt_note') or '关键节点：按施工组织设计进度安排')

    # 技术标
    P['fig_dir'] = fig_dir
    P['out'] = os.path.join(out_dir, f"{tag}-{plan.get('short', '项目')}-技术标（施工组织设计·A3四页）.docx")
    K = dict(ov=2, qty=10, ft=3, sg=1, m1=3, m2=2, pp=2, qu=3, sf=3, cv=2, pg=2, tf=2, s7=3,
             fig=float(plan.get('fig', 15.2)), brk=[])
    gen_a3.build(P, tag, K)

    # 附表
    apx_out = os.path.join(out_dir, f"{tag}-{plan.get('short', '项目')}-附表.docx")
    apx.build(apx_out, P['equip'], [list(x) for x in P['labor']],
              plan.get('land') or prose.DEFAULT_LAND,
              os.path.join(fig_dir, f'{tag}_gantt.png'),
              os.path.join(fig_dir, f'{tag}_plan.png'))
    print(f'[生成] {P["out"]}')
    print(f'[生成] {apx_out}')

    # PDF + 页数
    for f in (P['out'], apx_out):
        pages = _to_pdf(f)
        print(f'[PDF] {os.path.basename(f)} → {pages} 页' if pages else f'[PDF] {os.path.basename(f)} 转换失败')
    return 0


def _to_pdf(docx_path):
    pdf = os.path.splitext(docx_path)[0] + '.pdf'
    if os.path.exists(pdf):
        os.remove(pdf)
    try:
        if NODE and os.path.exists(LO_CLI):
            subprocess.run([NODE, LO_CLI, 'convert', '--input', docx_path, '--output', pdf],
                           capture_output=True, timeout=300)
        else:
            subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir',
                            os.path.dirname(docx_path), docx_path], capture_output=True, timeout=300)
    except Exception:
        return None
    if not os.path.exists(pdf):
        return None
    try:
        import pdfx
        return len(pdfx.page_numbers(pdfx.PDF(pdf)))
    except Exception:
        return '?'


def main(argv=None):
    ap = argparse.ArgumentParser(description='投标技术标与附表生成')
    sub = ap.add_subparsers(dest='cmd', required=True)
    pi = sub.add_parser('inspect', help='解压招标资料并抽取参数/清单量')
    pi.add_argument('--archive', required=True, help='项目 zip 或已解压目录')
    pi.add_argument('--work', default='work', help='工作目录（默认 work/）')
    pi.add_argument('--name', default=None, help='项目全称（可选，覆盖自动识别）')
    pi.set_defaults(func=cmd_inspect)
    pb = sub.add_parser('build', help='按 plan.json 生成技术标与附表')
    pb.add_argument('--work', default='work')
    pb.set_defaults(func=cmd_build)
    args = ap.parse_args(argv)
    return args.func(args)


if __name__ == '__main__':
    sys.exit(main())

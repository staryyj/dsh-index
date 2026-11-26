# -*- coding: utf-8 -*-
"""Text extraction for the two tender PDFs: Identity-H fonts + ToUnicode, position aware."""
import re, sys, zlib
sys.path.insert(0, '.')
from pdfx import PDF, page_numbers, parse_tounicode

NUM = re.compile(rb'[-+]?\d*\.?\d+')


def parse_resources_fonts(hdr):
    """name -> objnum from an inline /Resources<</Font<<...>>>> block"""
    out = {}
    m = re.search(rb'/Font\s*<<(.*?)>>', hdr, re.S)
    if m:
        for name, num in re.findall(rb'/([A-Za-z0-9_+.\-]+)\s+(\d+)\s+0\s+R', m.group(1)):
            out[name.decode('latin1')] = int(num)
    return out


def content_refs(hdr):
    m = re.search(rb'/Contents\s*\[(.*?)\]', hdr, re.S)
    if m:
        return [int(x) for x in re.findall(rb'(\d+)\s+0\s+R', m.group(1))]
    m = re.search(rb'/Contents\s+(\d+)\s+0\s+R', hdr)
    return [int(m.group(1))] if m else []


def font_decoder(pdf, fontnum):
    hdr, _ = pdf.objs.get(fontnum, (b'', None))
    sub = re.search(rb'/Subtype\s*/(\w+)', hdr)
    sub = sub.group(1) if sub else b''
    two_byte = (sub == b'Type0') or (b'Identity-H' in hdr)
    cmap = {}
    m = re.search(rb'/ToUnicode\s+(\d+)\s+0\s+R', hdr)
    if m:
        try:
            cmap = parse_tounicode(pdf.stream(int(m.group(1))))
        except Exception:
            cmap = {}
    if cmap and max(cmap) > 255:
        two_byte = True

    def dec(raw: bytes) -> str:
        out = []
        if two_byte:
            for i in range(0, len(raw) - 1, 2):
                code = (raw[i] << 8) | raw[i + 1]
                out.append(cmap.get(code, ''))
        else:
            for b in raw:
                out.append(cmap.get(b, chr(b) if 32 <= b < 127 else ''))
        return ''.join(out)
    return dec


ESCAPES = {b'n': b'\n', b'r': b'\r', b't': b'\t', b'b': b'\b', b'f': b'\f',
           b'(': b'(', b')': b')', b'\\': b'\\'}


def read_string(buf, i):
    """parse a PDF string starting at buf[i]; returns (bytes, next_index)"""
    if buf[i:i + 1] == b'(':
        depth = 1
        i += 1
        out = bytearray()
        while i < len(buf) and depth:
            c = buf[i:i + 1]
            if c == b'\\':
                nxt = buf[i + 1:i + 2]
                if nxt in ESCAPES:
                    out += ESCAPES[nxt]
                    i += 2
                elif nxt.isdigit():
                    oct_ = re.match(rb'[0-7]{1,3}', buf[i + 1:i + 4]).group(0)
                    out.append(int(oct_, 8) & 0xFF)
                    i += 1 + len(oct_)
                else:
                    i += 2
            elif c == b'(':
                depth += 1
                out += c
                i += 1
            elif c == b')':
                depth -= 1
                if depth:
                    out += c
                i += 1
            else:
                out += c
                i += 1
        return bytes(out), i
    if buf[i:i + 1] == b'<':
        e = buf.find(b'>', i)
        h = re.sub(rb'[^0-9A-Fa-f]', b'', buf[i + 1:e])
        if len(h) % 2:
            h += b'0'
        return bytes.fromhex(h.decode()), e + 1
    return b'', i + 1


TOKEN = re.compile(rb'''
    (?P<num>[-+]?\d*\.?\d+)
  | (?P<name>/[A-Za-z0-9_+.\-]+)
  | (?P<op>[A-Za-z'"*][A-Za-z0-9*'"]*)
  | (?P<arr>[\[\]])
  | (?P<str>\(|<)
''', re.X)


def extract_page(pdf, pagenum, decoders):
    hdr, _ = pdf.objs.get(pagenum, (b'', None))
    fonts = parse_resources_fonts(hdr)
    buf = b''.join(pdf.stream(r) for r in content_refs(hdr))
    segs = []
    stack = []
    cur_font = None
    tm = [1, 0, 0, 1, 0, 0]
    tlm = list(tm)
    leading = 0.0
    i = 0
    n = len(buf)
    while i < n:
        c = buf[i:i + 1]
        if c in b' \t\r\n\x00':
            i += 1
            continue
        if c in b'(<':
            s, i = read_string(buf, i)
            stack.append(('str', s))
            continue
        if c in b'[]':
            stack.append(('delim', c))
            i += 1
            continue
        m = TOKEN.match(buf, i)
        if not m:
            i += 1
            continue
        i = m.end()
        if m.lastgroup == 'num':
            stack.append(('num', float(m.group('num'))))
            continue
        if m.lastgroup == 'name':
            stack.append(('name', m.group('name')[1:].decode('latin1')))
            continue
        op = m.group('op').decode('latin1')
        if op == 'Tf':
            names = [v for t, v in stack if t == 'name']
            if names:
                cur_font = fonts.get(names[-1])
        elif op in ('Td', 'TD'):
            nums = [v for t, v in stack if t == 'num'][-2:]
            if len(nums) == 2:
                if op == 'TD':
                    leading = -nums[1]
                tlm = [tlm[0], tlm[1], tlm[2], tlm[3], tlm[4] + nums[0] * tlm[0],
                       tlm[5] + nums[1] * tlm[3]]
                tm = list(tlm)
        elif op == 'Tm':
            nums = [v for t, v in stack if t == 'num'][-6:]
            if len(nums) == 6:
                tm = list(nums)
                tlm = list(nums)
        elif op == 'TL':
            nums = [v for t, v in stack if t == 'num'][-1:]
            if nums:
                leading = nums[0]
        elif op == 'T*':
            tlm = [tlm[0], tlm[1], tlm[2], tlm[3], tlm[4], tlm[5] - leading]
            tm = list(tlm)
        elif op in ('Tj', "'", '"', 'TJ'):
            dec = decoders.get(cur_font)
            if dec:
                if op == 'TJ':
                    parts = [v for t, v in stack if t == 'str']
                else:
                    parts = [v for t, v in stack if t == 'str'][-1:]
                s = ''.join(dec(p) for p in parts)
                if s.strip():
                    segs.append((round(tm[5], 1), round(tm[4], 1), s))
        if op in ('Tj', "'", '"', 'TJ', 'Tf', 'Td', 'TD', 'Tm', 'TL', 'T*', 'BT', 'ET'):
            stack = []
    # group into lines by y
    segs.sort(key=lambda t: (-t[0], t[1]))
    lines = []
    for y, x, s in segs:
        if lines and abs(lines[-1][0] - y) <= 2.0:
            lines[-1][1].append((x, s))
        else:
            lines.append([y, [(x, s)]])
    out = []
    for y, items in lines:
        items.sort()
        out.append(''.join(s for _, s in items))
    return '\n'.join(out)


def extract(path, out_txt):
    pdf = PDF(path)
    pn = page_numbers(pdf)
    cache = {}
    with open(out_txt, 'w', encoding='utf-8') as f:
        for idx, pnum in enumerate(pn, 1):
            hdr, _ = pdf.objs.get(pnum, (b'', None))
            fonts = parse_resources_fonts(hdr)
            for name, num in fonts.items():
                if num not in cache:
                    cache[num] = font_decoder(pdf, num)
            txt = extract_page(pdf, pnum, cache)
            f.write(f'\n===== PAGE {idx} =====\n{txt}\n')
    return len(pn)


if __name__ == '__main__':
    for src, dst in (('zb_V2.pdf', 'zb_V2.txt'), ('zb_V3.pdf', 'zb_V3.txt')):
        n = extract(src, dst)
        import os
        print(f'{src} -> {dst}  pages={n}  chars={os.path.getsize(dst)}')

# -*- coding: utf-8 -*-
"""Minimal PDF text extractor for Aspose/Word generated PDFs (Identity-H + ToUnicode)."""
import re, sys, zlib

WS = b'\x00\t\n\x0c\r '


class PDF:
    def __init__(self, path):
        self.data = open(path, 'rb').read()
        self.objs = {}          # num -> (dict_bytes, stream_bytes|None)
        self._scan_objects()
        self._expand_objstm()

    def _scan_objects(self):
        data = self.data
        for m in re.finditer(rb'(?<![0-9])(\d{1,7})\s+(\d{1,5})\s+obj\b', data):
            num = int(m.group(1))
            start = m.end()
            end = data.find(b'endobj', start)
            if end < 0:
                continue
            body = data[start:end]
            sd = body.find(b'stream')
            if sd >= 0:
                header = body[:sd]
                p = sd + 6
                if body[p:p + 2] == b'\r\n':
                    p += 2
                elif body[p:p + 1] in (b'\n', b'\r'):
                    p += 1
                mlen = re.search(rb'/Length\s+(\d+)(\s+0\s+R)?', header)
                if mlen and not mlen.group(2):
                    ln = int(mlen.group(1))
                    raw = body[p:p + ln]
                else:
                    e2 = body.find(b'endstream', p)
                    raw = body[p:e2] if e2 >= 0 else body[p:]
                    if raw.endswith(b'\r\n'):
                        raw = raw[:-2]
                    elif raw.endswith((b'\n', b'\r')):
                        raw = raw[:-1]
                self.objs[num] = (header, raw)
            else:
                self.objs[num] = (body, None)

    def _expand_objstm(self):
        for num, (hdr, raw) in list(self.objs.items()):
            if raw is None or b'/ObjStm' not in hdr:
                continue
            try:
                st = self.decode(hdr, raw)
            except Exception:
                continue
            n = int(re.search(rb'/N\s+(\d+)', hdr).group(1))
            first = int(re.search(rb'/First\s+(\d+)', hdr).group(1))
            nums = list(map(int, st[:first].split()))
            pairs = [(nums[i], nums[i + 1]) for i in range(0, min(len(nums), 2 * n), 2)]
            for i, (onum, off) in enumerate(pairs):
                s = first + off
                e = first + pairs[i + 1][1] if i + 1 < len(pairs) else len(st)
                self.objs.setdefault(onum, (st[s:e], None))

    def decode(self, hdr, raw):
        filters = re.findall(rb'/(FlateDecode|ASCIIHexDecode|ASCII85Decode)', hdr)
        out = raw
        for f in filters:
            if f == b'FlateDecode':
                try:
                    out = zlib.decompressobj().decompress(out)
                except zlib.error:
                    out = zlib.decompressobj(-15).decompress(out[2:])
            elif f == b'ASCIIHexDecode':
                h = re.sub(rb'[^0-9A-Fa-f]', b'', out.split(b'>')[0])
                out = bytes.fromhex(h.decode()) if len(h) % 2 == 0 else bytes.fromhex(h.decode() + '0')
        return out

    def stream(self, num):
        hdr, raw = self.objs.get(num, (b'', None))
        if raw is None:
            return b''
        return self.decode(hdr, raw)

    def resolve_len(self, header):
        return header


def parse_tounicode(cmap_bytes):
    """code -> unicode string"""
    out = {}
    txt = cmap_bytes
    for blk in re.findall(rb'beginbfchar(.*?)endbfchar', txt, re.S):
        for src, dst in re.findall(rb'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', blk):
            try:
                code = int(src, 16)
                s = bytes.fromhex(dst.decode()).decode('utf-16-be', 'ignore')
                out[code] = s
            except Exception:
                pass
    for blk in re.findall(rb'beginbfrange(.*?)endbfrange', txt, re.S):
        for lo, hi, dst in re.findall(rb'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', blk):
            try:
                a, b = int(lo, 16), int(hi, 16)
                base = int(dst, 16)
                for i in range(a, min(b, a + 512) + 1):
                    out[i] = chr(base + (i - a))
            except Exception:
                pass
        for lo, hi, arr in re.findall(rb'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*\[(.*?)\]', blk, re.S):
            try:
                a = int(lo, 16)
                items = re.findall(rb'<([0-9A-Fa-f]+)>', arr)
                for i, it in enumerate(items):
                    out[a + i] = bytes.fromhex(it.decode()).decode('utf-16-be', 'ignore')
            except Exception:
                pass
    return out


def page_numbers(pdf):
    """ordered /Type /Page objects (walk Kids if possible)"""
    root = None
    for num, (hdr, _) in pdf.objs.items():
        if b'/Type' in hdr and b'/Catalog' in hdr:
            root = hdr
            break
    pages = []
    if root:
        m = re.search(rb'/Pages\s+(\d+)\s+0\s+R', root)
        if m:
            def walk(n, depth=0):
                if depth > 30:
                    return
                hdr, _ = pdf.objs.get(n, (b'', None))
                if re.search(rb'/Type\s*/Page\b', hdr):
                    pages.append(n)
                    return
                kids = re.search(rb'/Kids\s*\[(.*?)\]', hdr, re.S)
                if kids:
                    for k in re.findall(rb'(\d+)\s+0\s+R', kids.group(1)):
                        walk(int(k), depth + 1)
            walk(int(m.group(1)))
    if not pages:
        pages = sorted(n for n, (hdr, _) in pdf.objs.items()
                       if re.search(rb'/Type\s*/Page[^s]', hdr))
    return pages

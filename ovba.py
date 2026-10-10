#!/usr/bin/env python3
"""Низкоуровневые функции сборки VBA-проекта (MS-OVBA, MS-CFB).

- перекодировка русских строк VBA в вызовы U("...");
- сжатие потоков по MS-OVBA 2.4.1;
- запись составного файла OLE (MS-CFB, сектора 512 байт).
"""
import os, struct

ROOT = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(ROOT, 'src')

# ---------------------------------------------------------------- encoder
def enc_text(s):
    """Кодирует строку в ASCII-формат для функции U()."""
    out = []
    cyr = False
    for ch in s:
        o = ord(ch)
        if 0x410 <= o <= 0x44F or o in (0x401, 0x451):
            if not cyr:
                out.append('{'); cyr = True
            if o == 0x401: out.append('p')
            elif o == 0x451: out.append('q')
            else: out.append(chr(o - 0x410 + 0x30))
        elif ch in '{~' or o > 0x7E or o < 0x20:
            out.append('~%04X' % o)
        else:
            if cyr and 0x30 <= o <= 0x71:
                out.append('{'); cyr = False
            out.append(ch)
    return ''.join(out)

def dec_text(s):
    """Эталонный декодер (как U() в VBA) — для самопроверки."""
    out = []; cyr = False; i = 0
    while i < len(s):
        c = s[i]
        if c == '{': cyr = not cyr; i += 1
        elif c == '~': out.append(chr(int(s[i+1:i+5], 16))); i += 5
        else:
            o = ord(c)
            if cyr and 0x30 <= o <= 0x6F: out.append(chr(o - 0x30 + 0x410))
            elif cyr and c == 'p': out.append('\u0401')
            elif cyr and c == 'q': out.append('\u0451')
            else: out.append(c)
            i += 1
    return ''.join(out)

def transform_line(line):
    """Заменяет строковые литералы с не-ASCII символами на U("...")."""
    res = []; i = 0; n = len(line)
    while i < n:
        c = line[i]
        if c == '"':
            j = i + 1; buf = []
            while j < n:
                if line[j] == '"':
                    if j + 1 < n and line[j+1] == '"':
                        buf.append('"'); j += 2; continue
                    break
                buf.append(line[j]); j += 1
            if j >= n: raise ValueError('Unterminated string: ' + line)
            content = ''.join(buf)
            if any(ord(ch) > 0x7E for ch in content):
                e = enc_text(content)
                assert dec_text(e) == content, content
                res.append('U("' + e.replace('"', '""') + '")')
            else:
                res.append(line[i:j+1])
            i = j + 1
        elif c == "'":
            res.append(line[i:]); break
        else:
            res.append(c); i += 1
    return ''.join(res)

def load_module(name):
    with open(os.path.join(SRC, name + '.bas'), encoding='utf-8') as f:
        lines = f.read().replace('\r\n', '\n').split('\n')
    out = []
    for ln in lines:
        t = transform_line(ln)
        # комментарии: русские буквы допустимы в cp1251, прочее заменяем
        t = ''.join(ch if (ord(ch) < 0x80 or _in_cp(ch)) else '?' for ch in t)
        if len(t.encode('cp1251')) > 1000:
            raise ValueError('Line too long in %s: %s' % (name, t[:80]))
        if 'Const' in t and 'U("' in t:
            raise ValueError('Const with Cyrillic in %s: %s' % (name, t))
        out.append(t)
    while out and out[-1] == '':
        out.pop()
    return '\r\n'.join(out) + '\r\n'

def _in_cp(ch):
    try: ch.encode('cp1251'); return True
    except UnicodeEncodeError: return False

THISDOC = ('Attribute VB_Name = "ThisDocument"\r\n'
           'Attribute VB_Base = "1Normal.ThisDocument"\r\n'
           'Attribute VB_GlobalNameSpace = False\r\n'
           'Attribute VB_Creatable = False\r\n'
           'Attribute VB_PredeclaredId = True\r\n'
           'Attribute VB_Exposed = True\r\n'
           'Attribute VB_TemplateDerived = True\r\n'
           'Attribute VB_Customizable = True\r\n')

# ---------------------------------------------------------------- MS-OVBA compression
def _copy_token_help(diff):
    bc = max((diff - 1).bit_length(), 4) if diff > 1 else 4
    length_mask = 0xFFFF >> bc
    max_len = length_mask + 3
    return bc, length_mask, max_len

def compress_chunk(data):
    out = bytearray()
    pos = 0; n = len(data)
    table = {}
    while pos < n:
        flag_pos = len(out); out.append(0); flags = 0
        for bit in range(8):
            if pos >= n: break
            bc, lmask, max_len = _copy_token_help(pos)
            best_len = 0; best_off = 0
            if pos + 3 <= n:
                key = bytes(data[pos:pos+3])
                for cand in reversed(table.get(key, [])):
                    off = pos - cand
                    if off > (1 << bc): break
                    l = 0
                    while pos + l < n and l < max_len and data[cand + l] == data[pos + l]:
                        l += 1
                    if l > best_len:
                        best_len, best_off = l, off
                        if l == max_len: break
            if best_len >= 3:
                tok = ((best_off - 1) << (16 - bc)) | (best_len - 3)
                out += struct.pack('<H', tok)
                flags |= (1 << bit)
                for k in range(best_len):
                    if pos + k + 3 <= n:
                        table.setdefault(bytes(data[pos+k:pos+k+3]), []).append(pos + k)
                pos += best_len
            else:
                out.append(data[pos])
                if pos + 3 <= n:
                    table.setdefault(bytes(data[pos:pos+3]), []).append(pos)
                pos += 1
        out[flag_pos] = flags
    return bytes(out)

def compress(data):
    out = bytearray([1])
    for start in range(0, len(data), 4096):
        chunk = data[start:start+4096]
        comp = compress_chunk(chunk)
        if len(comp) <= 4096:
            hdr = (len(comp) + 2 - 3) | (0b011 << 12) | 0x8000
            out += struct.pack('<H', hdr) + comp
        else:
            assert len(chunk) == 4096, 'incompressible short chunk'
            out += struct.pack('<H', 0x0FFF | (0b011 << 12)) + chunk
    return bytes(out)

def rec(rid, payload):
    return struct.pack('<HI', rid, len(payload)) + payload

def u16(s): return s.encode('utf-16-le')
def mb(s): return s.encode('cp1251')

def build_projectwm(modules):
    b = bytearray()
    for mname, _ in modules:
        b += mb(mname) + b'\x00' + u16(mname) + b'\x00\x00'
    b += b'\x00\x00'
    return bytes(b)

# ---------------------------------------------------------------- MS-CFB writer
ENDOFCHAIN = 0xFFFFFFFE; FREESECT = 0xFFFFFFFF; FATSECT = 0xFFFFFFFD; NOSTREAM = 0xFFFFFFFF

class Node:
    def __init__(self, name, data=None):
        self.name = name; self.data = data; self.children = []
        self.is_storage = data is None
        self.left = self.right = self.child = NOSTREAM
        self.start = ENDOFCHAIN; self.size = 0; self.sid = None

def cfb_cmp_key(name):
    return (len(name), name.upper())

def write_cfb(root_children):
    root = Node('Root Entry'); root.children = root_children
    nodes = []
    def assign(node):
        node.sid = len(nodes); nodes.append(node)
        for ch in node.children: assign(ch)
    assign(root)
    def build_tree(children):
        ch = sorted(children, key=lambda c: cfb_cmp_key(c.name))
        def rec_(lo, hi):
            if lo > hi: return NOSTREAM
            mid = (lo + hi) // 2
            ch[mid].left = rec_(lo, mid - 1); ch[mid].right = rec_(mid + 1, hi)
            return ch[mid].sid
        return rec_(0, len(ch) - 1)
    for n in nodes:
        if n.is_storage: n.child = build_tree(n.children)
    streams = [n for n in nodes if not n.is_storage]
    # mini stream
    mini = bytearray(); minifat = []
    big = []
    for s in streams:
        s.size = len(s.data)
        if s.size == 0:
            s.start = ENDOFCHAIN
            continue
        if s.size < 4096:
            s.start = len(mini) // 64
            nsec = (s.size + 63) // 64
            for k in range(nsec):
                minifat.append(s.start + k + 1 if k < nsec - 1 else ENDOFCHAIN)
            mini += s.data + b'\x00' * (nsec * 64 - s.size)
        else:
            big.append(s)
    sectors = []  # list of bytes(512)
    fat = []
    def add_chain(data):
        start = len(sectors)
        nsec = max(1, (len(data) + 511) // 512)
        for k in range(nsec):
            sectors.append(data[k*512:(k+1)*512].ljust(512, b'\x00'))
            fat.append(start + k + 1 if k < nsec - 1 else ENDOFCHAIN)
        return start
    for s in big: s.start = add_chain(s.data)
    if mini:
        root.start = add_chain(bytes(mini)); root.size = len(mini)
    else:
        root.start = ENDOFCHAIN; root.size = 0
    mf_bytes = b''.join(struct.pack('<I', x) for x in minifat)
    if minifat:
        mf_bytes += struct.pack('<I', FREESECT) * ((-len(minifat)) % 128)
        minifat_start = add_chain(mf_bytes); n_minifat = len(mf_bytes) // 512
    else:
        minifat_start = ENDOFCHAIN; n_minifat = 0
    # directory
    def dir_entry(n):
        name = u16(n.name) + b'\x00\x00'
        e = name.ljust(64, b'\x00')
        typ = 5 if n is root else (1 if n.is_storage else 2)
        e += struct.pack('<HBB', len(name), typ, 1)
        e += struct.pack('<III', n.left, n.right, n.child)
        e += b'\x00' * 16 + struct.pack('<I', 0) + b'\x00' * 16
        start = n.start if (not n.is_storage or n is root) else 0
        size = n.size if (not n.is_storage or n is root) else 0
        e += struct.pack('<IQ', start, size)
        assert len(e) == 128
        return e
    dbytes = b''.join(dir_entry(n) for n in nodes)
    empty = b'\x00' * 64 + struct.pack('<HBB', 0, 0, 0) + struct.pack('<III', NOSTREAM, NOSTREAM, NOSTREAM) + b'\x00' * 36 + struct.pack('<IQ', 0, 0)
    while len(dbytes) % 512: dbytes += empty
    dir_start = add_chain(dbytes)
    # FAT sectors
    nfat = 1
    while True:
        total = len(sectors) + nfat
        if (total + 127) // 128 <= nfat: break
        nfat += 1
    assert nfat <= 109
    fat_start = len(sectors)
    for k in range(nfat): fat.append(FATSECT)
    fat += [FREESECT] * (nfat * 128 - len(fat))
    fbytes = b''.join(struct.pack('<I', x) for x in fat)
    for k in range(nfat): sectors.append(fbytes[k*512:(k+1)*512])
    difat = [fat_start + k for k in range(nfat)] + [FREESECT] * (109 - nfat)
    hdr = bytes.fromhex('D0CF11E0A1B11AE1') + b'\x00' * 16
    hdr += struct.pack('<HHHHH', 0x3E, 3, 0xFFFE, 9, 6) + b'\x00' * 6
    hdr += struct.pack('<IIIIIIIII', 0, nfat, dir_start, 0, 4096, minifat_start, n_minifat, ENDOFCHAIN, 0)
    hdr += b''.join(struct.pack('<I', x) for x in difat)
    assert len(hdr) == 512
    return hdr + b''.join(sectors)

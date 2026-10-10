#!/usr/bin/env python3
"""Очистка контейнеров Word в base/ от личных данных.

- docProps: автор, «кем изменён», компания — пустые;
- vbaProject.bin пересобирается: p-code и __SRP_* удаляются (там пути установки Office),
  в ссылке MSForms путь к кэшу C:\\Users\\<имя>\\... заменяется на нейтральный;
- исходный код модулей и хранилище формы сохраняются как есть.
Запуск: python3 sanitize_base.py base/word_AbbrGOST.dotm base/word_form.docm
"""
import sys, io, re, struct, zipfile, olefile
import ovba as B
from oletools.olevba import decompress_stream

NEUTRAL_USER = b'User'

def scrub_path(b):
    return re.sub(rb'(?i)(\\Users\\)[^\\]+(\\)', rb'\1' + NEUTRAL_USER + rb'\2', b)

def sanitize_dir(raw):
    d = bytes(decompress_stream(bytearray(raw)))
    out = b''; i = 0; mod_offsets = {}
    cur_mod = None
    while i < len(d):
        rid, = struct.unpack('<H', d[i:i+2])
        if rid == 0x0009:
            out += d[i:i+12]; i += 12; continue
        sz, = struct.unpack('<I', d[i+2:i+6])
        payload = d[i+6:i+6+sz]
        if rid == 0x0030:                      # REFERENCECONTROL, расширенная часть
            n, = struct.unpack('<I', payload[:4])
            lib = scrub_path(payload[4:4+n])
            payload = struct.pack('<I', len(lib)) + lib + payload[4+n:]
        elif rid in (0x000D, 0x000E, 0x002F, 0x0033):   # прочие ссылки: только чистка имени пользователя
            if rid == 0x002F:
                n, = struct.unpack('<I', payload[:4])
                lib = scrub_path(payload[4:4+n])
                payload = struct.pack('<I', len(lib)) + lib + payload[4+n:]
            elif rid == 0x000D:
                n, = struct.unpack('<I', payload[:4])
                lib = scrub_path(payload[4:4+n])
                payload = struct.pack('<I', len(lib)) + lib + payload[4+n:]
            elif rid == 0x0033:
                payload = scrub_path(payload)
        elif rid == 0x0019:
            cur_mod = payload.decode('cp1251')
        elif rid == 0x0031:
            payload = struct.pack('<I', 0)          # исходник с начала потока
        elif rid == 0x002C:
            payload = struct.pack('<H', 0xFFFF)
        out += struct.pack('<HI', rid, len(payload)) + payload
        i += 6 + sz
    assert b'fergu' not in out.lower()
    return B.compress(out)

def module_source(stream_bytes, offset):
    return bytes(decompress_stream(bytearray(stream_bytes[offset:])))

def dir_offsets(raw):
    d = bytes(decompress_stream(bytearray(raw))); i = 0; res = {}; cur = None
    while i < len(d):
        rid, = struct.unpack('<H', d[i:i+2])
        if rid == 0x0009: i += 12; continue
        sz, = struct.unpack('<I', d[i+2:i+6]); p = d[i+6:i+6+sz]
        if rid == 0x001A: cur = p.decode('cp1251')
        if rid == 0x0031: res[cur] = struct.unpack('<I', p)[0]
        i += 6 + sz
    return res

def rebuild_vba(vbin):
    o = olefile.OleFileIO(io.BytesIO(vbin))
    rawdir = o.openstream('VBA/dir').read()
    offs = dir_offsets(rawdir)
    vba = B.Node('VBA')
    vba.children = [B.Node('_VBA_PROJECT', bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00])),
                    B.Node('dir', sanitize_dir(rawdir))]
    for mod, off in offs.items():
        src = module_source(o.openstream('VBA/' + mod).read(), off)
        vba.children.append(B.Node(mod, B.compress(src)))
    top = [vba]
    for st in o.listdir(storages=True):
        if len(st) == 1 and o.get_type(st[0]) == olefile.STGTY_STORAGE and st[0] != 'VBA':
            node = B.Node(st[0])
            node.children = [B.Node(s[1], o.openstream('/'.join(s)).read()) for s in o.listdir() if s[0] == st[0]]
            top.append(node)
    proj = o.openstream('PROJECT').read()
    assert b'fergu' not in proj.lower()
    top += [B.Node('PROJECT', proj), B.Node('PROJECTwm', o.openstream('PROJECTwm').read())]
    return B.write_cfb(top)

def clean_core(xml):
    s = xml.decode('utf-8')
    for tag in ('dc:creator', 'cp:lastModifiedBy', 'dc:description'):
        s = re.sub(r'<%s>[^<]*</%s>' % (tag, tag), '<%s></%s>' % (tag, tag), s)
    return s.encode('utf-8')

def clean_app(xml):
    s = xml.decode('utf-8')
    s = re.sub(r'<Company>[^<]*</Company>', '<Company></Company>', s)
    s = re.sub(r'<Manager>[^<]*</Manager>', '', s)
    return s.encode('utf-8')

def sanitize(path):
    zin = zipfile.ZipFile(path)
    items = [(it, zin.read(it.filename)) for it in zin.infolist()]
    zin.close()
    zout = zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED)
    for it, data in items:
        if it.filename == 'word/vbaProject.bin': data = rebuild_vba(data)
        elif it.filename == 'docProps/core.xml': data = clean_core(data)
        elif it.filename == 'docProps/app.xml': data = clean_app(data)
        zout.writestr(it, data)
    zout.close()

if __name__ == '__main__':
    for p in sys.argv[1:]:
        sanitize(p); print('cleaned', p)

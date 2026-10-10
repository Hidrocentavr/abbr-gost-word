#!/usr/bin/env python3
"""Сборка надстройки AbbrGOST.dotm на основе контейнеров, созданных самим Word.

base/word_AbbrGOST.dotm — рабочий шаблон пользователя (проект, ссылки stdole/Office, CMG/DPB/GC)
base/word_form.docm     — пустая UserForm1, созданная Word (хранилище формы, ссылка MSForms)
Код подменяется исходником; p-code и __SRP_* удаляются, _VBA_PROJECT = CC61FFFF000000,
поэтому Word перекомпилирует проект из исходного текста при открытии.
"""
import os, io, struct, zipfile, olefile
import ovba as B
from oletools.olevba import decompress_stream

ROOT = os.path.dirname(os.path.abspath(__file__))
STD_MODULES = ['modAbbrMain', 'modAbbrScan', 'modAbbrOut', 'modAbbrUtil']
FORM = 'UserForm1'

def ole_from_zip(path):
    z = zipfile.ZipFile(path)
    return olefile.OleFileIO(io.BytesIO(z.read('word/vbaProject.bin')))

def parse_dir(raw):
    d = decompress_stream(bytearray(raw))
    recs = []; i = 0
    while i < len(d):
        rid, = struct.unpack('<H', d[i:i+2])
        if rid == 0x0009:
            recs.append((rid, bytes(d[i:i+12]))); i += 12; continue
        sz, = struct.unpack('<I', d[i+2:i+6])
        recs.append((rid, bytes(d[i:i+6+sz]))); i += 6 + sz
    return recs

def ref_block(recs):
    """Сырые байты всех записей ссылок (от первой 0x16 до 0x0F)."""
    out = b''; on = False
    for rid, raw in recs:
        if rid == 0x16: on = True
        if rid == 0x0F: break
        if on: out += raw
    return out

def info_block(recs, name):
    out = b''
    for rid, raw in recs:
        if rid == 0x16: break
        if rid == 0x04: raw = B.rec(0x0004, B.mb(name))
        out += raw
    return out

def module_recs(mname, mtype, private=False):
    d = b''
    d += B.rec(0x0019, B.mb(mname)) + B.rec(0x0047, B.u16(mname))
    d += B.rec(0x001A, B.mb(mname)) + B.rec(0x0032, B.u16(mname))
    d += B.rec(0x001C, b'') + B.rec(0x0048, b'')
    d += B.rec(0x0031, struct.pack('<I', 0))
    d += B.rec(0x001E, struct.pack('<I', 0))
    d += B.rec(0x002C, struct.pack('<H', 0xFFFF))
    d += B.rec(0x0021 if mtype == 'std' else 0x0022, b'')
    if private: d += B.rec(0x0028, b'')
    d += B.rec(0x002B, b'')
    return d

def main():
    word = ole_from_zip(os.path.join(ROOT, 'base/word_AbbrGOST.dotm'))
    form = ole_from_zip(os.path.join(ROOT, 'base/word_form.docm'))
    wrecs = parse_dir(word.openstream('VBA/dir').read())
    frecs = parse_dir(form.openstream('VBA/dir').read())

    # ссылки: stdole + Office из шаблона пользователя, MSForms из формы
    refs = ref_block(wrecs)
    frefs = b''; on = False
    for rid, raw in frecs:
        if rid == 0x16 and raw[6:6+7] == b'MSForms': on = True
        if rid == 0x0F: break
        if on: frefs += raw
    assert frefs, 'MSForms reference not found'
    modules = [('ThisDocument', 'doc', False)] + [(m, 'std', False) for m in STD_MODULES] + [(FORM, 'doc', True)]
    d = info_block(wrecs, 'AbbrGOST') + refs + frefs
    d += B.rec(0x000F, struct.pack('<H', len(modules))) + B.rec(0x0013, struct.pack('<H', 0xFFFF))
    for m, t, p in modules: d += module_recs(m, t, p)
    d += B.rec(0x0010, b'')

    # исходники
    sources = {'ThisDocument': B.THISDOC}
    for m in STD_MODULES: sources[m] = B.load_module(m)
    from oletools.olevba import VBA_Parser
    fhdr = None
    for (_, _, vn, code) in VBA_Parser(os.path.join(ROOT, 'base/word_form.docm')).extract_macros():
        if vn.startswith(FORM):
            fhdr = [l for l in code.replace('\r\n', '\n').split('\n') if l.startswith('Attribute ')]
    assert fhdr and len(fhdr) >= 7, fhdr
    body = B.load_module('UserForm1.code')
    sources[FORM] = '\r\n'.join(fhdr) + '\r\n' + body

    # PROJECT: ID и CMG/DPB/GC — из шаблона, созданного Word
    wp = word.openstream('PROJECT').read().decode('cp1251').split('\r\n')
    keep = {k: next(l for l in wp if l.startswith(k + '=')) for k in ('ID', 'HelpContextID', 'VersionCompatible32', 'CMG', 'DPB', 'GC')}
    L = [keep['ID'], 'Document=ThisDocument/&H00000000',
         'Package={AC9F2F90-E877-11CE-9F68-00AA00574A4F}', 'BaseClass=%s' % FORM]
    L += ['Module=%s' % m for m in STD_MODULES]
    L += ['Name="AbbrGOST"', keep['HelpContextID'], keep['VersionCompatible32'], keep['CMG'], keep['DPB'], keep['GC'], '',
          '[Host Extender Info]', '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '', '[Workspace]',
          'ThisDocument=0, 0, 0, 0, C']
    L += ['%s=0, 0, 0, 0, C' % m for m in STD_MODULES]
    L += ['%s=0, 0, 0, 0, C, 32, 32, 448, 659, ' % FORM]
    project = ('\r\n'.join(L) + '\r\n').encode('cp1251')
    pwm = B.build_projectwm([(m, t) for m, t, _ in modules])

    vba = B.Node('VBA')
    vba.children = [B.Node('_VBA_PROJECT', bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00])),
                    B.Node('dir', B.compress(d))]
    for m, _, _ in modules:
        vba.children.append(B.Node(m, B.compress(sources[m].encode('cp1251'))))
    fst = B.Node(FORM)
    fst.children = [B.Node(n, form.openstream(FORM + '/' + n).read()) for n in ('\x01CompObj', '\x03VBFrame', 'f', 'o')]
    vbin = B.write_cfb([vba, fst, B.Node('PROJECT', project), B.Node('PROJECTwm', pwm)])

    out = os.path.join(ROOT, 'out'); os.makedirs(out, exist_ok=True)
    open(os.path.join(out, 'vbaProject.bin'), 'wb').write(vbin)
    zin = zipfile.ZipFile(os.path.join(ROOT, 'base/word_AbbrGOST.dotm'))
    zout = zipfile.ZipFile(os.path.join(out, 'AbbrGOST.dotm'), 'w', zipfile.ZIP_DEFLATED)
    cui = open(os.path.join(B.SRC, 'customUI14.xml'), 'rb').read()
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename == 'word/vbaProject.bin': data = vbin
        elif it.filename == 'customUI/customUI14.xml': data = cui
        zout.writestr(it, data)
    zout.close()
    # исходники для справки
    sd = os.path.join(out, 'src'); os.makedirs(sd, exist_ok=True)
    for m, s in sources.items(): open(os.path.join(sd, m + '.txt'), 'wb').write(s.encode('cp1251'))
    print('OK', len(vbin))

if __name__ == '__main__':
    main()

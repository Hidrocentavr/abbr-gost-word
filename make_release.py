#!/usr/bin/env python3
"""Сборка архива для релиза: надстройка, установщик, тестовый отчёт.

python3 make_release.py 2.1   →  dist/AbbrGOST_v2.1.zip
Используется GitHub Actions (.github/workflows/release.yml) при отправке метки vX.Y.
"""
import os, sys, shutil, zipfile, subprocess

ROOT = os.path.dirname(os.path.abspath(__file__))
os.chdir(ROOT)

INSTALL = r'''@echo off
chcp 65001 >nul
set "SRC=%~dp0AbbrGOST.dotm"
set "DST=%APPDATA%\Microsoft\Word\STARTUP"
set "BAK=%APPDATA%\AbbrGOST"
if not exist "%SRC%" (
  echo Не найден файл AbbrGOST.dotm рядом с установщиком. Распакуйте архив целиком.
  pause & exit /b 1
)
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul && (
  echo Закройте Word и запустите установку ещё раз.
  pause & exit /b 1
)
if not exist "%DST%" mkdir "%DST%"
if not exist "%BAK%" mkdir "%BAK%"
if exist "%DST%\AbbrGOST.dotm" copy /Y "%DST%\AbbrGOST.dotm" "%BAK%\AbbrGOST_previous.dotm" >nul
copy /Y "%SRC%" "%DST%\AbbrGOST.dotm" >nul || (
  echo Не удалось скопировать файл в "%DST%".
  pause & exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File -LiteralPath (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\AbbrGOST.dotm')" >nul 2>&1
echo.
echo Готово! Запустите Word, вкладка «Сокращения ГОСТ» → «Проверить».
echo Прежняя версия сохранена: %BAK%\AbbrGOST_previous.dotm
echo.
pause
'''

ROLLBACK = r'''@echo off
chcp 65001 >nul
set "BAK=%APPDATA%\AbbrGOST\AbbrGOST_previous.dotm"
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul && (
  echo Закройте Word и запустите ещё раз.
  pause & exit /b 1
)
if not exist "%BAK%" (
  echo Резервная копия не найдена.
  pause & exit /b 1
)
copy /Y "%BAK%" "%APPDATA%\Microsoft\Word\STARTUP\AbbrGOST.dotm" >nul
echo Возвращена прежняя версия надстройки.
pause
'''

UNINSTALL = r'''@echo off
chcp 65001 >nul
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul && (
  echo Закройте Word и запустите удаление ещё раз.
  pause & exit /b 1
)
del /Q "%APPDATA%\Microsoft\Word\STARTUP\AbbrGOST.dotm" 2>nul
echo Надстройка удалена. Ваш словарь сохранён в папке %APPDATA%\AbbrGOST
pause
'''

README = '''СОКРАЩЕНИЯ ГОСТ {ver} — надстройка для Microsoft Word (Windows)

УСТАНОВКА
1. Распакуйте архив целиком и закройте Word.
2. Запустите «Установить.cmd». Текущая версия сохранится в резервную копию.
3. Откройте Word → вкладка «Сокращения ГОСТ» → «Проверить».
Если Windows спросит про файл из интернета — «Подробнее» → «Выполнить в любом случае».

ПЕРВЫЙ ЗАПУСК
Откройте «Тестовый отчёт.docx» — в нём собраны типовые ошибки — и нажмите «Проверить».

ЕСЛИ ЧТО-ТО ПОШЛО НЕ ТАК
«Вернуть прежнюю версию.cmd» возвращает версию, которая стояла до обновления.

ВАШ СЛОВАРЬ
%APPDATA%\\AbbrGOST\\dictionary.txt и exceptions.txt — сохраняются при обновлении.
'''

def main():
    ver = (sys.argv[1] if len(sys.argv) > 1 else 'dev').lstrip('v')
    subprocess.check_call([sys.executable, 'build.py'])
    D = os.path.join('dist', 'AbbrGOST')
    shutil.rmtree('dist', ignore_errors=True)
    os.makedirs(D)
    shutil.copy(os.path.join('out', 'AbbrGOST.dotm'), D)
    subprocess.check_call([sys.executable, 'make_testdoc.py', os.path.join(D, 'Тестовый отчёт.docx')])
    for name, txt in [('Установить.cmd', INSTALL), ('Вернуть прежнюю версию.cmd', ROLLBACK), ('Удалить.cmd', UNINSTALL)]:
        open(os.path.join(D, name), 'w', encoding='utf-8', newline='\r\n').write(txt)
    open(os.path.join(D, 'Прочитай меня.txt'), 'w', encoding='utf-8-sig', newline='\r\n').write(README.format(ver=ver))
    # проверка: в надстройке нет личных данных
    z = zipfile.ZipFile(os.path.join(D, 'AbbrGOST.dotm'))
    blob = b''.join(z.read(n) for n in z.namelist())
    for bad in (b'fergu', 'Шевченко'.encode('utf-8')):
        assert bad not in blob, 'personal data in AbbrGOST.dotm'
    out = os.path.join('dist', 'AbbrGOST_v%s.zip' % ver)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zf:
        for r, _, fs in os.walk(D):
            for f in sorted(fs):
                p = os.path.join(r, f)
                zf.write(p, os.path.relpath(p, 'dist'))
    print(out)

if __name__ == '__main__':
    main()

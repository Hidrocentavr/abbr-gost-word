Attribute VB_Name = "modAbbrUtil"
Option Explicit
' =====================================================================
'  Сокращения ГОСТ — вспомогательные функции
'  Все русские строковые литералы при сборке перекодируются в U("..."),
'  поэтому надстройка работает при любой системной кодировке Windows.
' =====================================================================

Private mStop As Object
Private mRef As Object

' ---------- декодер строк ----------
Public Function U(ByVal s As String) As String
    Dim i As Long, n As Long, c As String, code As Long, cyr As Boolean
    Dim out As String
    n = Len(s)
    i = 1
    Do While i <= n
        c = Mid$(s, i, 1)
        If c = "{" Then
            cyr = Not cyr
            i = i + 1
        ElseIf c = "~" Then
            out = out & ChrW(CLng(Val("&H" & Mid$(s, i + 1, 4) & "&")))
            i = i + 5
        Else
            code = AscW(c)
            If cyr And code >= &H30 And code <= &H6F Then
                out = out & ChrW(code - &H30 + &H410)
            ElseIf cyr And c = "p" Then
                out = out & ChrW(&H401)
            ElseIf cyr And c = "q" Then
                out = out & ChrW(&H451)
            Else
                out = out & c
            End If
            i = i + 1
        End If
    Loop
    U = out
End Function

' ---------- классы символов ----------
Public Function IsUpperL(ByVal c As Long) As Boolean
    IsUpperL = (c >= 65 And c <= 90) Or (c >= &H410 And c <= &H42F) Or c = &H401
End Function

Public Function IsLowerL(ByVal c As Long) As Boolean
    IsLowerL = (c >= 97 And c <= 122) Or (c >= &H430 And c <= &H44F) Or c = &H451
End Function

Public Function IsLatinL(ByVal c As Long) As Boolean
    IsLatinL = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
End Function

Public Function IsCyrL(ByVal c As Long) As Boolean
    IsCyrL = (c >= &H410 And c <= &H44F) Or c = &H401 Or c = &H451
End Function

Public Function IsDigitL(ByVal c As Long) As Boolean
    IsDigitL = (c >= 48 And c <= 57)
End Function

Public Function IsAlnumL(ByVal c As Long) As Boolean
    IsAlnumL = IsUpperL(c) Or IsLowerL(c) Or IsDigitL(c) Or (c >= &H391 And c <= &H3C9)
End Function

' ---------- разбиение на слова ----------
' Возвращает количество слов; toks() — слова, starts() — позиции (с 1).
Public Function Tokenize(ByVal s As String, ByRef toks() As String, ByRef starts() As Long) As Long
    Dim n As Long, i As Long, j As Long, c As Long, cnt As Long
    n = Len(s)
    ReDim toks(0 To 31)
    ReDim starts(0 To 31)
    i = 1
    Do While i <= n
        c = AscW(Mid$(s, i, 1))
        If IsAlnumL(c) Then
            j = i + 1
            Do While j <= n
                c = AscW(Mid$(s, j, 1))
                If IsAlnumL(c) Then
                    j = j + 1
                Else
                    Exit Do
                End If
            Loop
            If cnt > UBound(toks) Then
                ReDim Preserve toks(0 To cnt * 2)
                ReDim Preserve starts(0 To cnt * 2)
            End If
            toks(cnt) = Mid$(s, i, j - i)
            starts(cnt) = i
            cnt = cnt + 1
            i = j
        Else
            i = i + 1
        End If
    Loop
    Tokenize = cnt
End Function

Public Function IsRoman(ByVal t As String) As Boolean
    Dim i As Long
    If Len(t) > 5 Then Exit Function
    For i = 1 To Len(t)
        If InStr("IVX", Mid$(t, i, 1)) = 0 Then Exit Function
    Next
    IsRoman = True
End Function

' Похоже ли слово на аббревиатуру: ПЛК, ЧПУ, ROS2, IoT, EtherCAT
Public Function IsAbbrToken(ByVal t As String) As Boolean
    Dim i As Long, c As Long, up As Long, lo As Long, n As Long
    n = Len(t)
    If n < 2 Or n > 16 Then Exit Function
    If Not IsUpperL(AscW(Left$(t, 1))) Then Exit Function
    For i = 1 To n
        c = AscW(Mid$(t, i, 1))
        If IsUpperL(c) Then
            up = up + 1
        ElseIf IsLowerL(c) Then
            lo = lo + 1
        End If
    Next
    If up < 2 Then Exit Function
    If lo > up Then Exit Function
    If IsRoman(t) Then Exit Function
    IsAbbrToken = True
End Function

' Заглавные буквы аббревиатуры (для сопоставления с расшифровкой)
Public Function UpperLetters(ByVal t As String) As String
    Dim i As Long, c As Long, r As String
    For i = 1 To Len(t)
        c = AscW(Mid$(t, i, 1))
        If IsUpperL(c) Then r = r & Mid$(t, i, 1)
    Next
    UpperLetters = r
End Function

Public Function HasLatin(ByVal t As String) As Boolean
    Dim i As Long
    For i = 1 To Len(t)
        If IsLatinL(AscW(Mid$(t, i, 1))) Then HasLatin = True: Exit Function
    Next
End Function

Public Function HasCyr(ByVal t As String) As Boolean
    Dim i As Long
    For i = 1 To Len(t)
        If IsCyrL(AscW(Mid$(t, i, 1))) Then HasCyr = True: Exit Function
    Next
End Function

Public Function HasDigit(ByVal t As String) As Boolean
    Dim i As Long
    For i = 1 To Len(t)
        If IsDigitL(AscW(Mid$(t, i, 1))) Then HasDigit = True: Exit Function
    Next
End Function

Public Function IsMixedScript(ByVal t As String) As Boolean
    IsMixedScript = HasLatin(t) And HasCyr(t)
End Function

' Латинские буквы, внешне совпадающие с русскими
Private Function HomoLat() As String
    HomoLat = "ABCEHKMOPTXYaceopxy"
End Function

Private Function HomoCyr() As String
    HomoCyr = ChrW(&H410) & ChrW(&H412) & ChrW(&H421) & ChrW(&H415) & ChrW(&H41D) & ChrW(&H41A) & _
              ChrW(&H41C) & ChrW(&H41E) & ChrW(&H420) & ChrW(&H422) & ChrW(&H425) & ChrW(&H423) & _
              ChrW(&H430) & ChrW(&H441) & ChrW(&H435) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H445) & ChrW(&H443)
End Function

' Приводит слово со смешанными алфавитами к одному алфавиту.
' Возвращает "" если однозначно привести нельзя.
Public Function FixMixed(ByVal t As String) As String
    Dim i As Long, c As Long, nl As Long, nc As Long, ch As String, k As Long, r As String
    Dim hl As String, hc As String
    hl = HomoLat(): hc = HomoCyr()
    For i = 1 To Len(t)
        c = AscW(Mid$(t, i, 1))
        If IsLatinL(c) Then nl = nl + 1
        If IsCyrL(c) Then nc = nc + 1
    Next
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        c = AscW(ch)
        If nc >= nl Then
            If IsLatinL(c) Then
                k = InStr(1, hl, ch, vbBinaryCompare)
                If k = 0 Then Exit Function
                ch = Mid$(hc, k, 1)
            End If
        Else
            If IsCyrL(c) Then
                k = InStr(1, hc, ch, vbBinaryCompare)
                If k = 0 Then Exit Function
                ch = Mid$(hl, k, 1)
            End If
        End If
        r = r & ch
    Next
    FixMixed = r
End Function

' Буквы a встречаются в b в том же порядке (без учёта регистра)
Public Function SubseqIn(ByVal a As String, ByVal b As String) As Boolean
    Dim i As Long, p As Long
    a = UCase$(a): b = UCase$(b)
    p = 1
    For i = 1 To Len(a)
        p = InStr(p, b, Mid$(a, i, 1), vbBinaryCompare)
        If p = 0 Then Exit Function
        p = p + 1
    Next
    SubseqIn = True
End Function

Private Sub InitWordLists()
    Dim a As Variant, i As Long
    If Not mStop Is Nothing Then Exit Sub
    Set mStop = CreateObject("Scripting.Dictionary")
    a = Split("и в во на по для с со к ко о об от из при за под над без до у а или как его ее их of and for the in on to a an with by de", " ")
    For i = LBound(a) To UBound(a)
        If Not mStop.Exists(a(i)) Then mStop.Add a(i), 1
    Next
    Set mRef = CreateObject("Scripting.Dictionary")
    a = Split("см рис рисунок рисунке табл таблица таблице формула формулы раздел разделе приложение приложении стр с п пп например т е англ от от англ", " ")
    For i = LBound(a) To UBound(a)
        If Not mRef.Exists(a(i)) Then mRef.Add a(i), 1
    Next
End Sub

Public Function IsStop(ByVal w As String) As Boolean
    InitWordLists
    IsStop = mStop.Exists(LCase$(w))
End Function

Public Function IsRefWord(ByVal w As String) As Boolean
    InitWordLists
    IsRefWord = mRef.Exists(LCase$(w))
End Function

' Сопоставление букв аббревиатуры L с началами слов, идя назад от слова lastIdx.
' Возвращает индекс первого слова расшифровки или -1.
Public Function BackMatch(words() As String, ByVal lastIdx As Long, ByVal L As String) As Long
    Dim j As Long, i As Long, w As String, k As Long, ok As Boolean
    BackMatch = -1
    j = Len(L)
    i = lastIdx
    Do While i >= 0 And j >= 1
        w = UCase$(words(i))
        ok = False
        For k = j To 1 Step -1
            If Left$(w, 1) = Mid$(L, k, 1) Then
                If SubseqIn(Mid$(L, k + 1, j - k), Mid$(w, 2)) Then
                    ok = True
                    Exit For
                End If
            End If
        Next
        If ok Then
            j = k - 1
            If j = 0 Then
                BackMatch = i
                Exit Function
            End If
        ElseIf Not IsStop(words(i)) Then
            Exit Function
        End If
        i = i - 1
        If lastIdx - i > 12 Then Exit Function
    Loop
End Function

Public Function AllStopBefore(words() As String, ByVal idx As Long) As Boolean
    Dim i As Long
    For i = 0 To idx - 1
        If Not IsStop(words(i)) Then Exit Function
    Next
    AllStopBefore = True
End Function

' Одинаковые ли расшифровки с точностью до падежных окончаний
Public Function SameDec(ByVal a As String, ByVal b As String) As Boolean
    Dim wa As Variant, wb As Variant, i As Long, x As String, y As String, k As Long, m As Long
    a = LCase$(NormSpaces(Replace(a, "-", " ")))
    b = LCase$(NormSpaces(Replace(b, "-", " ")))
    If a = b Then SameDec = True: Exit Function
    wa = Split(a, " "): wb = Split(b, " ")
    If UBound(wa) <> UBound(wb) Then Exit Function
    For i = 0 To UBound(wa)
        x = wa(i): y = wb(i)
        If x <> y Then
            m = Len(x)
            If Len(y) < m Then m = Len(y)
            k = 0
            Do While k < m
                If Mid$(x, k + 1, 1) <> Mid$(y, k + 1, 1) Then Exit Do
                k = k + 1
            Loop
            If k < 3 Or k < m - 3 Then Exit Function
        End If
    Next
    SameDec = True
End Function

' ---------- строки ----------
Public Function NormSpaces(ByVal s As String) As String
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, Chr(11), " ")
    s = Replace(s, Chr(12), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormSpaces = Trim$(s)
End Function

' Нормализация заголовка: верхний регистр, без нумерации и точки в конце
Public Function NormHead(ByVal s As String) As String
    Dim i As Long, c As String
    s = UCase$(NormSpaces(s))
    i = 1
    Do While i <= Len(s)
        c = Mid$(s, i, 1)
        If (c >= "0" And c <= "9") Or c = "." Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    If i > 1 Then s = Trim$(Mid$(s, i))
    Do While Len(s) > 0
        If InStr(".:;", Right$(s, 1)) > 0 Then
            s = Left$(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop
    NormHead = Trim$(s)
End Function

Public Function WordCount(ByVal s As String) As Long
    s = NormSpaces(s)
    If Len(s) = 0 Then Exit Function
    WordCount = UBound(Split(s, " ")) + 1
End Function

Public Function IsCapsText(ByVal s As String) As Boolean
    Dim i As Long, c As Long, up As Long, lo As Long
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If IsUpperL(c) Then
            up = up + 1
        ElseIf IsLowerL(c) Then
            lo = lo + 1
        End If
    Next
    IsCapsText = (up >= 8 And lo * 5 <= up)
End Function

Public Function EndsWithDigit(ByVal s As String) As Boolean
    s = NormSpaces(s)
    If Len(s) = 0 Then Exit Function
    EndsWithDigit = IsDigitL(AscW(Right$(s, 1)))
End Function

Public Function StripEndPunct(ByVal s As String) As String
    s = Trim$(s)
    Do While Len(s) > 0
        If InStr(".,;:", Right$(s, 1)) > 0 Then
            s = Trim$(Left$(s, Len(s) - 1))
        Else
            Exit Do
        End If
    Loop
    StripEndPunct = s
End Function

Public Function EndsWithPunct(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    EndsWithPunct = (InStr(".,;:", Right$(s, 1)) > 0)
End Function

Public Function StripLeadDash(ByVal s As String) As String
    s = Trim$(s)
    Do While Len(s) > 0
        If InStr("-" & ChrW(&H2013) & ChrW(&H2014) & ChrW(&H2012), Left$(s, 1)) > 0 Then
            s = Trim$(Mid$(s, 2))
        Else
            Exit Do
        End If
    Loop
    StripLeadDash = s
End Function

' Строка перечня "КЛЮЧ — расшифровка". sepKind: 1 тире, 2 дефис, 3 табуляция без тире
Public Function SplitEntry(ByVal line As String, ByRef key As String, ByRef dec As String, ByRef sepKind As Long) As Boolean
    Dim seps As Variant, kinds As Variant, i As Long, p As Long, best As Long, bestKind As Long, bestLen As Long
    line = Replace(line, ChrW(160), " ")
    seps = Array(vbTab, " " & ChrW(&H2014) & " ", " " & ChrW(&H2013) & " ", " - ", ChrW(&H2014), ChrW(&H2013))
    kinds = Array(3, 1, 1, 2, 1, 1)
    best = 0
    For i = 0 To UBound(seps)
        p = InStr(1, line, seps(i), vbBinaryCompare)
        If p > 1 Then
            If best = 0 Or p < best Then
                best = p: bestKind = kinds(i): bestLen = Len(seps(i))
            End If
        End If
    Next
    If best = 0 Then Exit Function
    key = Trim$(Left$(line, best - 1))
    dec = Mid$(line, best + bestLen)
    If bestKind = 3 Then
        If InStr("-" & ChrW(&H2013) & ChrW(&H2014), Left$(Trim$(dec), 1)) > 0 Then
            If Left$(Trim$(dec), 1) = "-" Then bestKind = 2 Else bestKind = 1
        End If
    End If
    dec = Trim$(StripLeadDash(dec))
    If Len(key) = 0 Or Len(key) > 40 Or Len(dec) = 0 Then Exit Function
    sepKind = bestKind
    SplitEntry = True
End Function

' Ключ сортировки: кириллица, затем латиница, затем прочие обозначения
Public Function SortKey(ByVal k As String) As String
    Dim c As Long, g As String
    If Len(k) = 0 Then SortKey = "9": Exit Function
    c = AscW(Left$(k, 1))
    If IsCyrL(c) Then
        g = "1"
    ElseIf IsLatinL(c) Then
        g = "2"
    ElseIf IsDigitL(c) Then
        g = "4"
    Else
        g = "3"
    End If
    k = UCase$(k)
    k = Replace(k, ChrW(&H401), ChrW(&H415))
    SortKey = g & k
End Function

' ---------- файлы словаря ----------
Public Function AppDir() As String
    Dim fso As Object, p As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    p = Environ$("APPDATA") & "\AbbrGOST"
    If Not fso.FolderExists(p) Then fso.CreateFolder p
    AppDir = p
End Function

Public Function DictPath() As String
    DictPath = AppDir() & "\dictionary.txt"
End Function

Public Function ExcPath() As String
    ExcPath = AppDir() & "\exceptions.txt"
End Function

Public Function ReadUtf8(ByVal path As String) As String
    Dim st As Object
    On Error GoTo fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText(-1)
    st.Close
    Exit Function
fail:
    ReadUtf8 = ""
End Function

Public Sub WriteUtf8(ByVal path As String, ByVal text As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    st.SaveToFile path, 2
    st.Close
End Sub

Public Function FileExists(ByVal path As String) As Boolean
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(path)
End Function

' Словарь из файла: строки "КЛЮЧ<TAB>расшифровка" или "КЛЮЧ — расшифровка"
Public Function LoadKV(ByVal path As String) As Object
    Dim d As Object, txt As String, lines As Variant, i As Long, ln As String
    Dim k As String, v As String, sk As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 0
    txt = ReadUtf8(path)
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If Len(ln) > 0 Then
            If Left$(ln, 1) <> "#" Then
                If SplitEntry(ln, k, v, sk) Then
                    If Not d.Exists(k) Then d.Add k, v
                Else
                    k = Trim$(ln)
                    If Not d.Exists(k) Then d.Add k, ""
                End If
            End If
        End If
    Next
    Set LoadKV = d
End Function

' Добавить или заменить строку с ключом key в файле
Public Sub UpsertKV(ByVal path As String, ByVal key As String, ByVal val As String)
    Dim txt As String, lines As Variant, i As Long, ln As String, k As String, v As String, sk As Long
    Dim found As Boolean, newLine As String, out As String
    If Len(val) > 0 Then
        newLine = key & vbTab & val
    Else
        newLine = key
    End If
    txt = ReadUtf8(path)
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = lines(i)
        k = ""
        If Len(Trim$(ln)) > 0 And Left$(Trim$(ln), 1) <> "#" Then
            If SplitEntry(Trim$(ln), k, v, sk) = False Then k = Trim$(ln)
        End If
        If k = key And Not found Then
            ln = newLine
            found = True
        End If
        If i < UBound(lines) Or Len(ln) > 0 Then out = out & ln & vbCrLf
    Next
    If Not found Then out = out & newLine & vbCrLf
    WriteUtf8 path, out
End Sub

' Удалить строку с ключом key из файла
Public Sub DeleteKV(ByVal path As String, ByVal key As String)
    Dim txt As String, lines As Variant, i As Long, ln As String, k As String, v As String, sk As Long, out As String
    txt = ReadUtf8(path)
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = lines(i)
        k = ""
        If Len(Trim$(ln)) > 0 And Left$(Trim$(ln), 1) <> "#" Then
            If SplitEntry(Trim$(ln), k, v, sk) = False Then k = Trim$(ln)
        End If
        If k <> key Then
            If i < UBound(lines) Or Len(ln) > 0 Then out = out & ln & vbCrLf
        End If
    Next
    WriteUtf8 path, out
End Sub

Public Sub EnsureDefaults()
    If Not FileExists(DictPath()) Then WriteUtf8 DictPath(), DefaultDictText()
    If Not FileExists(ExcPath()) Then WriteUtf8 ExcPath(), DefaultExcText()
End Sub

Private Function DL(ByVal k As String, ByVal v As String) As String
    DL = k & vbTab & v & vbCrLf
End Function

Public Function DefaultDictText() As String
    Dim s As String
    s = "# Личный словарь сокращений надстройки «Сокращения ГОСТ»" & vbCrLf
    s = s & "# Формат: СОКРАЩЕНИЕ<TAB>расшифровка в именительном падеже (или через « — »)" & vbCrLf
    s = s & "# Строки, начинающиеся с #, — комментарии. Файл можно править в Блокноте." & vbCrLf
    s = s & DL("АСУ", "автоматизированная система управления")
    s = s & DL("АЦП", "аналого-цифровой преобразователь")
    s = s & DL("БД", "база данных")
    s = s & DL("ВКР", "выпускная квалификационная работа")
    s = s & DL("ГНСС", "глобальная навигационная спутниковая система")
    s = s & DL("ДПТ", "двигатель постоянного тока")
    s = s & DL("ЕСКД", "Единая система конструкторской документации")
    s = s & DL("ЕСПД", "Единая система программной документации")
    s = s & DL("ИИ", "искусственный интеллект")
    s = s & DL("КА", "космический аппарат")
    s = s & DL("КД", "конструкторская документация")
    s = s & DL("ЛВС", "локальная вычислительная сеть")
    s = s & DL("МКС", "Международная космическая станция")
    s = s & DL("МКЭ", "метод конечных элементов")
    s = s & DL("МНК", "метод наименьших квадратов")
    s = s & DL("НИОКР", "научно-исследовательские и опытно-конструкторские работы")
    s = s & DL("НИР", "научно-исследовательская работа")
    s = s & DL("ОЗУ", "оперативное запоминающее устройство")
    s = s & DL("ОКР", "опытно-конструкторская работа")
    s = s & DL("ОС", "операционная система")
    s = s & DL("ПЗУ", "постоянное запоминающее устройство")
    s = s & DL("ПИД", "пропорционально-интегрально-дифференцирующий")
    s = s & DL("ПК", "персональный компьютер")
    s = s & DL("ПЛК", "программируемый логический контроллер")
    s = s & DL("ПО", "программное обеспечение")
    s = s & DL("ПР", "промышленный робот")
    s = s & DL("РТК", "робототехнический комплекс")
    s = s & DL("РТС", "робототехническая система")
    s = s & DL("САПР", "система автоматизированного проектирования")
    s = s & DL("САУ", "система автоматического управления")
    s = s & DL("СУ", "система управления")
    s = s & DL("ТЗ", "техническое задание")
    s = s & DL("ТП", "технологический процесс")
    s = s & DL("ЦАП", "цифро-аналоговый преобразователь")
    s = s & DL("ЧПУ", "числовое программное управление")
    s = s & DL("ШИМ", "широтно-импульсная модуляция")
    s = s & DL("ЭВМ", "электронная вычислительная машина")
    s = s & DL("API", "Application Programming Interface (программный интерфейс приложения)")
    s = s & DL("CAD", "Computer-Aided Design (система автоматизированного проектирования)")
    s = s & DL("CAE", "Computer-Aided Engineering (система инженерного анализа)")
    s = s & DL("CAM", "Computer-Aided Manufacturing (система подготовки производства)")
    s = s & DL("CNC", "Computer Numerical Control (числовое программное управление)")
    s = s & DL("CPU", "Central Processing Unit (центральный процессор)")
    s = s & DL("FPGA", "Field-Programmable Gate Array (программируемая логическая интегральная схема)")
    s = s & DL("GPU", "Graphics Processing Unit (графический процессор)")
    s = s & DL("HMI", "Human-Machine Interface (человеко-машинный интерфейс)")
    s = s & DL("IMU", "Inertial Measurement Unit (инерциальный измерительный модуль)")
    s = s & DL("IoT", "Internet of Things (интернет вещей)")
    s = s & DL("LiDAR", "Light Detection and Ranging (лазерный дальномер-сканер)")
    s = s & DL("PID", "Proportional-Integral-Derivative (пропорционально-интегрально-дифференцирующий)")
    s = s & DL("PLC", "Programmable Logic Controller (программируемый логический контроллер)")
    s = s & DL("ROS", "Robot Operating System (операционная система для роботов)")
    s = s & DL("SCADA", "Supervisory Control and Data Acquisition (диспетчерское управление и сбор данных)")
    s = s & DL("SLAM", "Simultaneous Localization and Mapping (одновременная локализация и построение карты)")
    s = s & DL("USB", "Universal Serial Bus (универсальная последовательная шина)")
    DefaultDictText = s
End Function

Public Function DefaultExcText() As String
    Dim s As String
    s = "# Исключения: общепринятые сокращения и единицы, которые не нужно расшифровывать." & vbCrLf
    s = s & "# По одному на строку. Эти слова надстройка не проверяет и не вносит в перечень." & vbCrLf
    s = s & "РФ" & vbCrLf & "СССР" & vbCrLf & "США" & vbCrLf & "ГОСТ" & vbCrLf & "ОСТ" & vbCrLf
    s = s & "СТО" & vbCrLf & "ТУ" & vbCrLf & "СП" & vbCrLf & "СНиП" & vbCrLf & "ISO" & vbCrLf
    s = s & "ИСО" & vbCrLf & "IEC" & vbCrLf & "МЭК" & vbCrLf & "IEEE" & vbCrLf & "ООО" & vbCrLf
    s = s & "ОАО" & vbCrLf & "ЗАО" & vbCrLf & "АО" & vbCrLf & "ПАО" & vbCrLf & "РАН" & vbCrLf
    s = s & "УДК" & vbCrLf & "ББК" & vbCrLf & "ISBN" & vbCrLf & "DOI" & vbCrLf & "URL" & vbCrLf
    s = s & "СИ" & vbCrLf & "МГц" & vbCrLf & "ГГц" & vbCrLf & "МПа" & vbCrLf & "ГПа" & vbCrLf
    s = s & "МВт" & vbCrLf & "ГВт" & vbCrLf & "МДж" & vbCrLf & "ГДж" & vbCrLf & "МОм" & vbCrLf
    s = s & "КБ" & vbCrLf & "МБ" & vbCrLf & "ГБ" & vbCrLf & "ТБ" & vbCrLf & "Мбит" & vbCrLf & "Гбит" & vbCrLf
    DefaultExcText = s
End Function

' ---------- настройки ----------
Public Function GetOpt(ByVal name As String, ByVal def As String) As String
    GetOpt = GetSetting("AbbrGOST", "Main", name, def)
End Function

Public Sub SetOpt(ByVal name As String, ByVal v As String)
    SaveSetting "AbbrGOST", "Main", name, v
End Sub

' 0 — ГОСТ 7.32-2017, 1 — ГОСТ Р 7.0.11-2011, 2 — ГОСТ Р 2.105-2019
Public Function GostIdx() As Long
    GostIdx = Val(GetOpt("Gost", "0"))
    If GostIdx < 0 Or GostIdx > 2 Then GostIdx = 0
End Function

Public Function GostName() As String
    Select Case GostIdx()
        Case 1: GostName = "ГОСТ Р 7.0.11-2011"
        Case 2: GostName = "ГОСТ Р 2.105-2019"
        Case Else: GostName = "ГОСТ 7.32-2017"
    End Select
End Function

Public Function ListTitleFor() As String
    Select Case GostIdx()
        Case 1: ListTitleFor = "СПИСОК СОКРАЩЕНИЙ И УСЛОВНЫХ ОБОЗНАЧЕНИЙ"
        Case 2: ListTitleFor = "Обозначения и сокращения"
        Case Else: ListTitleFor = "ПЕРЕЧЕНЬ СОКРАЩЕНИЙ И ОБОЗНАЧЕНИЙ"
    End Select
End Function

Public Function OptComments() As Boolean
    OptComments = (GetOpt("Comments", "0") = "1")
End Function

Public Function OptStrict() As Boolean
    OptStrict = (GetOpt("Strict", "1") = "1")
End Function

Public Function OptTabStyle() As Boolean
    OptTabStyle = (GetOpt("ListStyle", "0") = "0")
End Function

' ---------- поддержка автора ----------
Public Function TipUrl() As String
    TipUrl = "https://pay.cloudtips.ru/p/b787322c"
End Function

Public Sub OpenTipPage()
    On Error Resume Next
    CreateObject("Shell.Application").ShellExecute TipUrl()
End Sub

' Напоминание о чаевых — не чаще раза в 30 дней
Public Function TipDue() As Boolean
    Dim last As String
    last = GetOpt("TipLast", "")
    If Len(last) = 0 Then
        TipDue = True
    ElseIf IsDate(last) Then
        TipDue = (Date - CDate(last) >= 30)
    Else
        TipDue = True
    End If
    If TipDue Then SetOpt "TipLast", CStr(Date)
End Function

Public Function AppTitle() As String
    AppTitle = "Сокращения ГОСТ"
End Function

Public Function TagC() As String
    TagC = "[Сокращения ГОСТ] "
End Function

Public Function Dash() As String
    Dash = ChrW(&H2014)
End Function

Public Function Q(ByVal s As String) As String
    Q = ChrW(&HAB) & s & ChrW(&HBB)
End Function

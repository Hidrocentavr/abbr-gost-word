Attribute VB_Name = "modAbbrScan"
Option Explicit
' =====================================================================
'  Сокращения ГОСТ — анализ документа
' =====================================================================

Public Type TAbbr
    Key As String
    Count As Long
    FirstPara As Long
    FirstOff As Long
    FirstAnyPara As Long
    FirstAnyOff As Long
    DecCount As Long
    FirstDecPara As Long
    FirstDecOff As Long
    DecAtFirst As Boolean
    DecText As String
    DecFromA As Boolean
    DecUncertain As Boolean
    DecVariants As String
    InList As Boolean
    ListDec As String
    ListIdx As Long
    ListSpelling As String
    DictDec As String
    Mixed As Boolean
    MixTarget As String
    AliasOf As String
    Variants As String
    VariantCount As Long
    Issues As String
    Severity As Long
    Page As Long
End Type

' зоны документа
Public Const Z_FRONT As Long = 0
Public Const Z_PRE As Long = 1
Public Const Z_LIST As Long = 2
Public Const Z_BODY As Long = 3
Public Const Z_BIB As Long = 4
Public Const Z_TOC As Long = 5
Public Const Z_HEAD As Long = 6

Public gDoc As Document
Public gA() As TAbbr
Public gN As Long
Public gIdx As Object
Public gCanon As Object
Public gDict As Object
Public gExc As Object
Public gScanned As Boolean
Public gOverride As Object

Public gPN As Long
Public gPText() As String
Public gPStart() As Long
Public gPEnd() As Long
Public gPZone() As Long
Public gPHead() As Boolean
Public gPInTbl() As Boolean
Public gPToc() As Boolean

Public gListHead As Long
Public gListIntro As Long
Public gListLast As Long
Public gListTitle As String
Public gLEN As Long
Public gLEKey() As String
Public gLEDec() As String
Public gLEPara() As Long
Public gLEPunct() As Boolean
Public gLEIsAbbr() As Boolean
Public gLESep() As Long

Public gParaIntro As Long
Public gParaTerms As Long
Public gParaConcl As Long
Public gParaBib As Long
Public gParaToc As Long

Public gFixN As Long
Public gFixP() As Long
Public gFixO() As Long
Public gFixT() As String
Public gFixNew() As String

Public gSN As Long
Public gSText() As String
Public gSSev() As Long
Public gNeeded As Long

' ---------------------------------------------------------------------
Public Function RunScan(ByVal doc As Document) As Boolean
    Set gDoc = doc
    gScanned = False
    EnsureDefaults
    Set gDict = LoadKV(DictPath())
    Set gExc = LoadKV(ExcPath())
    If gOverride Is Nothing Then
        Set gOverride = CreateObject("Scripting.Dictionary")
        gOverride.CompareMode = 0
    End If
    Application.ScreenUpdating = False
    Application.StatusBar = "Сокращения ГОСТ: читаю документ..."
    ReadParagraphs
    DetectZones
    ParseList
    Application.StatusBar = "Сокращения ГОСТ: ищу сокращения..."
    ScanTokens
    MergeSources
    Application.StatusBar = "Сокращения ГОСТ: проверяю по " & GostName() & "..."
    Evaluate
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    gScanned = True
    RunScan = True
End Function

' ---------------------------------------------------------------------
Private Sub ReadParagraphs()
    Dim p As Paragraph, r As Range, i As Long, n As Long, t As String, sn As String
    Dim tS() As Long, tE() As Long, tn As Long, tb As Table, k As Long
    n = gDoc.Paragraphs.Count
    gPN = n
    ReDim gPText(0 To n): ReDim gPStart(0 To n): ReDim gPEnd(0 To n)
    ReDim gPZone(0 To n): ReDim gPHead(0 To n): ReDim gPInTbl(0 To n): ReDim gPToc(0 To n)
    i = 0
    For Each p In gDoc.Paragraphs
        If i >= n Then Exit For
        Set r = p.Range
        t = r.Text
        Do While Len(t) > 0
            If Right$(t, 1) = vbCr Or Right$(t, 1) = Chr(7) Then
                t = Left$(t, Len(t) - 1)
            Else
                Exit Do
            End If
        Loop
        gPText(i) = t
        gPStart(i) = r.Start
        gPEnd(i) = r.End
        gPHead(i) = (p.OutlineLevel <> wdOutlineLevelBodyText)
        sn = ""
        If EndsWithDigit(t) Then
            On Error Resume Next
            sn = LCase$(p.Style.NameLocal)
            On Error GoTo 0
        End If
        gPToc(i) = (Left$(sn, 3) = "toc") Or (Left$(sn, 10) = "оглавление")
        i = i + 1
        If i Mod 300 = 0 Then
            Application.StatusBar = "Сокращения ГОСТ: читаю документ... " & i & " / " & n
            DoEvents
        End If
    Next
    gPN = i
    ' отметить абзацы в таблицах
    tn = 0
    ReDim tS(0 To 15): ReDim tE(0 To 15)
    For Each tb In gDoc.Tables
        If tn > UBound(tS) Then
            ReDim Preserve tS(0 To tn * 2): ReDim Preserve tE(0 To tn * 2)
        End If
        tS(tn) = tb.Range.Start: tE(tn) = tb.Range.End
        tn = tn + 1
    Next
    If tn > 0 Then
        k = 0
        For i = 0 To gPN - 1
            Do While k < tn
                If tE(k) <= gPStart(i) Then k = k + 1 Else Exit Do
            Loop
            If k < tn Then
                If gPStart(i) >= tS(k) And gPStart(i) < tE(k) Then gPInTbl(i) = True
            End If
        Next
    End If
End Sub

' Вид структурного заголовка: 0 — не заголовок
Private Function HeadKind(ByVal i As Long) As Long
    Dim h As String, raw As String, wc As Long, headLike As Boolean
    If gPInTbl(i) Then Exit Function
    raw = NormSpaces(gPText(i))
    If Len(raw) = 0 Or Len(raw) > 150 Then Exit Function
    wc = WordCount(raw)
    ' строки оглавления (заканчиваются номером страницы) — не заголовки
    If EndsWithDigit(raw) And Left$(NormHead(raw), 10) <> "ПРИЛОЖЕНИЕ" Then Exit Function
    headLike = gPHead(i) Or IsCapsText(raw) Or UCase$(raw) = raw Or (wc <= 6 And Right$(raw, 1) <> ".")
    If Not headLike Then Exit Function
    h = NormHead(raw)
    Select Case h
        Case "РЕФЕРАТ", "АННОТАЦИЯ", "ПРЕДИСЛОВИЕ", "СПИСОК ИСПОЛНИТЕЛЕЙ"
            HeadKind = 1
        Case "СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ"
            HeadKind = 2
        Case "ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ", "ОПРЕДЕЛЕНИЯ", "СПИСОК ТЕРМИНОВ", "СЛОВАРЬ ТЕРМИНОВ", "ТЕРМИНЫ"
            HeadKind = 3
        Case "ВВЕДЕНИЕ"
            HeadKind = 7
        Case "ЗАКЛЮЧЕНИЕ", "ВЫВОДЫ", "ОБЩИЕ ВЫВОДЫ"
            HeadKind = 8
    End Select
    If HeadKind <> 0 Then Exit Function
    If InStr(h, "СОКРАЩЕНИ") > 0 And wc <= 12 Then
        HeadKind = 4
        Exit Function
    End If
    If wc <= 8 Then
        If (InStr(h, "СПИСОК") > 0 And (InStr(h, "ИСТОЧНИК") > 0 Or InStr(h, "ЛИТЕРАТУР") > 0)) _
           Or InStr(h, "БИБЛИОГРАФ") > 0 Or InStr(h, "ССЫЛОЧНЫЕ") > 0 Then
            HeadKind = 6
            Exit Function
        End If
    End If
    If Left$(h, 10) = "ПРИЛОЖЕНИЕ" And wc <= 12 Then HeadKind = 9
End Function

Private Sub DetectZones()
    Dim i As Long, k As Long, z As Long, firstStruct As Long
    Dim kk As String, dd As String, sk As Long, t As String
    gParaIntro = -1: gParaTerms = -1: gParaConcl = -1: gParaBib = -1: gParaToc = -1
    gListHead = -1: gListTitle = ""
    firstStruct = -1
    For i = 0 To gPN - 1
        If HeadKind(i) > 0 Then firstStruct = i: Exit For
    Next
    If firstStruct <= 0 Then z = Z_BODY Else z = Z_FRONT
    For i = 0 To gPN - 1
        k = HeadKind(i)
        If k > 0 Then
            Select Case k
                Case 1: z = Z_PRE
                Case 2: z = Z_TOC: If gParaToc < 0 Then gParaToc = i
                Case 3: z = Z_PRE: If gParaTerms < 0 Then gParaTerms = i
                Case 4
                    If gListHead < 0 Then
                        z = Z_LIST
                        gListHead = i
                        gListTitle = NormSpaces(gPText(i))
                    Else
                        z = Z_BODY
                    End If
                Case 6: z = Z_BIB: If gParaBib < 0 Then gParaBib = i
                Case 7: z = Z_BODY: If gParaIntro < 0 Then gParaIntro = i
                Case 8: z = Z_BODY: gParaConcl = i
                Case 9: z = Z_BODY
            End Select
            gPZone(i) = Z_HEAD
        Else
            t = NormSpaces(gPText(i))
            If z = Z_TOC Then
                If Not gPToc(i) And Len(t) > 0 And Not EndsWithDigit(t) Then z = Z_BODY
            ElseIf z = Z_LIST Then
                If Not gPInTbl(i) And Len(t) > 0 Then
                    If gPHead(i) Then
                        z = Z_BODY
                    ElseIf LCase$(Left$(t, 11)) <> "в настоящем" Then
                        If Not SplitEntry(t, kk, dd, sk) And (IsCapsText(t) Or WordCount(t) <= 6) Then z = Z_BODY
                    End If
                End If
            ElseIf z = Z_PRE Then
                If gPHead(i) Then z = Z_BODY
            End If
            If gPToc(i) Then
                gPZone(i) = Z_TOC
            Else
                gPZone(i) = z
            End If
        End If
    Next
End Sub

' ---------------------------------------------------------------------
Private Sub AddListEntry(ByVal k As String, ByVal d As String, ByVal para As Long, ByVal sk As Long)
    If gLEN = 0 Then
        ReDim gLEKey(0 To 31): ReDim gLEDec(0 To 31): ReDim gLEPara(0 To 31)
        ReDim gLEPunct(0 To 31): ReDim gLEIsAbbr(0 To 31): ReDim gLESep(0 To 31)
    ElseIf gLEN > UBound(gLEKey) Then
        ReDim Preserve gLEKey(0 To gLEN * 2): ReDim Preserve gLEDec(0 To gLEN * 2)
        ReDim Preserve gLEPara(0 To gLEN * 2): ReDim Preserve gLEPunct(0 To gLEN * 2)
        ReDim Preserve gLEIsAbbr(0 To gLEN * 2): ReDim Preserve gLESep(0 To gLEN * 2)
    End If
    gLEKey(gLEN) = Trim$(k)
    gLEPunct(gLEN) = EndsWithPunct(d)
    gLEDec(gLEN) = StripEndPunct(d)
    gLEPara(gLEN) = para
    gLEIsAbbr(gLEN) = IsAbbrToken(Trim$(k))
    gLESep(gLEN) = sk
    gLEN = gLEN + 1
End Sub

Public Function ParaIndexOf(ByVal pos As Long) As Long
    Dim lo As Long, hi As Long, m As Long
    lo = 0: hi = gPN - 1
    Do While lo < hi
        m = (lo + hi + 1) \ 2
        If gPStart(m) <= pos Then lo = m Else hi = m - 1
    Loop
    ParaIndexOf = lo
End Function

Private Function CellText(ByVal c As Cell) As String
    Dim t As String
    t = c.Range.Text
    t = Replace(t, Chr(7), "")
    t = Replace(t, vbCr, " ")
    CellText = Trim$(NormSpaces(t))
End Function

Private Sub ParseList()
    Dim i As Long, k As String, d As String, t As String, sk As Long
    Dim tb As Table, r As Long, c1 As String, c2 As String, lastC As Long, cp As Long
    gLEN = 0: gListIntro = -1: gListLast = -1
    If gListHead < 0 Then Exit Sub
    gListLast = gListHead
    For i = gListHead + 1 To gPN - 1
        If gPZone(i) <> Z_LIST Then Exit For
        gListLast = i
        t = Trim$(Replace(gPText(i), ChrW(160), " "))
        If Len(t) > 0 And Not gPInTbl(i) Then
            If gListIntro < 0 And gLEN = 0 And LCase$(Left$(t, 11)) = "в настоящем" Then
                gListIntro = i
            ElseIf SplitEntry(t, k, d, sk) Then
                AddListEntry k, d, i, sk
            End If
        End If
    Next
    ' перечень, оформленный таблицей
    If gListLast > gListHead Then
        For Each tb In gDoc.Tables
            If tb.Range.Start >= gPStart(gListHead) And tb.Range.Start < gPEnd(gListLast) Then
                On Error Resume Next
                lastC = 0
                lastC = tb.Columns.Count
                On Error GoTo 0
                If lastC >= 2 Then
                    For r = 1 To tb.Rows.Count
                        c1 = "": c2 = "": cp = gListHead
                        On Error Resume Next
                        c1 = CellText(tb.Cell(r, 1))
                        c2 = CellText(tb.Cell(r, lastC))
                        cp = ParaIndexOf(tb.Cell(r, 1).Range.Start)
                        On Error GoTo 0
                        c2 = StripLeadDash(c2)
                        If InStr(LCase$(c1), "сокращ") > 0 Or InStr(LCase$(c1), "обознач") > 0 Then c1 = ""
                        If Len(c1) > 0 And Len(c2) > 0 And Len(c1) <= 40 Then
                            AddListEntry c1, c2, cp, 1
                        End If
                    Next
                End If
            End If
        Next
    End If
End Sub

' ---------------------------------------------------------------------
Private Function NewEntry(ByVal t As String) As Long
    If gN = 0 Then
        ReDim gA(0 To 63)
    ElseIf gN > UBound(gA) Then
        ReDim Preserve gA(0 To gN * 2)
    End If
    gA(gN).Key = t
    gA(gN).FirstPara = -1
    gA(gN).FirstAnyPara = -1
    gA(gN).FirstDecPara = -1
    gA(gN).ListIdx = -1
    gIdx.Add t, gN
    NewEntry = gN
    gN = gN + 1
End Function

Public Function GetIdx(ByVal t As String) As Long
    If gIdx.Exists(t) Then
        GetIdx = gIdx(t)
    Else
        GetIdx = NewEntry(t)
    End If
End Function

Private Sub AddFix(ByVal p As Long, ByVal o As Long, ByVal t As String, ByVal nw As String)
    If gFixN = 0 Then
        ReDim gFixP(0 To 31): ReDim gFixO(0 To 31): ReDim gFixT(0 To 31): ReDim gFixNew(0 To 31)
    ElseIf gFixN > UBound(gFixP) Then
        ReDim Preserve gFixP(0 To gFixN * 2): ReDim Preserve gFixO(0 To gFixN * 2)
        ReDim Preserve gFixT(0 To gFixN * 2): ReDim Preserve gFixNew(0 To gFixN * 2)
    End If
    gFixP(gFixN) = p: gFixO(gFixN) = o: gFixT(gFixN) = t: gFixNew(gFixN) = nw
    gFixN = gFixN + 1
End Sub

Public Sub AddToList(ByRef lst As String, ByVal v As String)
    Dim parts As Variant, i As Long, nv As String
    nv = NormSpaces(v)
    If Len(nv) = 0 Then Exit Sub
    If Len(lst) > 0 Then
        parts = Split(lst, Chr(1))
        For i = 0 To UBound(parts)
            If SameDec(CStr(parts(i)), nv) Then Exit Sub
        Next
        lst = lst & Chr(1) & v
    Else
        lst = v
    End If
End Sub

Private Function ScanZone(ByVal i As Long) As Boolean
    Dim z As Long
    z = gPZone(i)
    If z <> Z_BODY And z <> Z_PRE Then Exit Function
    If IsCapsText(gPText(i)) Then Exit Function
    ScanZone = True
End Function

' Есть ли расшифровка у вхождения tok в позиции p строки s
Public Function FindDecoding(ByVal s As String, ByVal p As Long, ByVal tok As String, _
        ByRef dec As String, ByRef fromA As Boolean, ByRef unc As Boolean) As Boolean
    Dim L As String, a As Long, b As Long, q As Long, n As Long, ch As String
    Dim seg As String, w() As String, ws() As Long, cnt As Long, idx As Long, k As Long, delim As String
    L = UpperLetters(tok)
    If Len(L) < 2 Then Exit Function
    n = Len(s)
    a = p - 1
    Do While a >= 1
        ch = Mid$(s, a, 1)
        If ch <> " " And ch <> ChrW(160) Then Exit Do
        a = a - 1
    Loop
    b = p + Len(tok)
    Do While b <= n
        ch = Mid$(s, b, 1)
        If ch <> " " And ch <> ChrW(160) Then Exit Do
        b = b + 1
    Loop
    ' Вариант Б: «полное название (АББР)»
    If a >= 1 And b <= n Then
        If Mid$(s, a, 1) = "(" And Mid$(s, b, 1) = ")" Then
            delim = ".;:()[]!?" & ChrW(&HAB) & ChrW(&HBB) & Chr(34) & vbTab & Chr(11)
            q = a - 1
            Do While q >= 1
                If InStr(delim, Mid$(s, q, 1)) > 0 Then Exit Do
                q = q - 1
            Loop
            seg = Mid$(s, q + 1, a - q - 1)
            cnt = Tokenize(seg, w, ws)
            If cnt > 0 Then
                idx = BackMatch(w, cnt - 1, L)
                If idx >= 0 Then
                    dec = Trim$(Mid$(seg, ws(idx)))
                    unc = False
                Else
                    k = cnt - Len(L)
                    If k < 0 Then k = 0
                    If k > 0 Then
                        If IsStop(w(k)) Then k = k - 1
                    End If
                    dec = Trim$(Mid$(seg, ws(k)))
                    unc = True
                End If
                dec = StripEndPunct(dec)
                Do While Right$(dec, 1) = ","
                    dec = Trim$(Left$(dec, Len(dec) - 1))
                Loop
                fromA = False
                FindDecoding = (Len(dec) > 0)
                Exit Function
            End If
        End If
    End If
    ' Вариант А: «АББР (полное название)»
    If b <= n Then
        If Mid$(s, b, 1) = "(" Then
            q = InStr(b + 1, s, ")")
            If q > 0 And q - b < 300 Then
                seg = Mid$(s, b + 1, q - b - 1)
                If InStr(seg, "(") = 0 Then
                    cnt = Tokenize(seg, w, ws)
                    If cnt >= 1 Then
                        idx = BackMatch(w, cnt - 1, L)
                        If idx >= 0 Then
                            If AllStopBefore(w, idx) Then
                                dec = StripEndPunct(seg): fromA = True: unc = False
                                FindDecoding = True
                                Exit Function
                            End If
                        End If
                        ' латинская аббревиатура с русским переводом в скобках
                        If HasLatin(tok) And Not HasCyr(tok) And cnt >= 2 And HasCyr(seg) And Not HasDigit(seg) Then
                            If Not IsRefWord(w(0)) Then
                                dec = StripEndPunct(seg): fromA = True: unc = False
                                FindDecoding = True
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Private Sub ScanTokens()
    Dim i As Long, n As Long, j As Long, t As String, toks() As String, st() As Long
    Dim ix As Long, z As Long, dec As String, fromA As Boolean, unc As Boolean, tu As String
    Dim cnt As Object, best As Object, k As Variant, mt As String
    Set gIdx = CreateObject("Scripting.Dictionary")
    gIdx.CompareMode = 0
    gN = 0
    gFixN = 0
    ' --- проход 1: сокращения и расшифровки
    For i = 0 To gPN - 1
        If ScanZone(i) Then
            z = gPZone(i)
            n = Tokenize(gPText(i), toks, st)
            For j = 0 To n - 1
                t = toks(j)
                If IsAbbrToken(t) Then
                    If Not gExc.Exists(t) Then
                        ix = GetIdx(t)
                        dec = "": fromA = False: unc = False
                        With gA(ix)
                            .Count = .Count + 1
                            If .FirstAnyPara < 0 Then .FirstAnyPara = i: .FirstAnyOff = st(j)
                            If z = Z_BODY And Not gPHead(i) And .FirstPara < 0 Then .FirstPara = i: .FirstOff = st(j)
                        End With
                        If FindDecoding(gPText(i), st(j), t, dec, fromA, unc) Then
                            With gA(ix)
                                .DecCount = .DecCount + 1
                                If .FirstDecPara < 0 Then .FirstDecPara = i: .FirstDecOff = st(j)
                                If .FirstPara = i And .FirstOff = st(j) Then .DecAtFirst = True
                                If Len(.DecText) = 0 Or (.DecUncertain And Not unc) Then
                                    .DecText = dec: .DecFromA = fromA: .DecUncertain = unc
                                End If
                                If Not unc Then AddToList .DecVariants, dec
                            End With
                        End If
                        If IsMixedScript(t) Then
                            mt = FixMixed(t)
                            gA(ix).Mixed = True
                            gA(ix).MixTarget = mt
                            If Len(mt) > 0 Then AddFix i, st(j), t, mt
                        End If
                    End If
                End If
            Next
        End If
        If i Mod 300 = 0 Then DoEvents
    Next
    ' --- каноническое написание для групп, различающихся только регистром
    Set gCanon = CreateObject("Scripting.Dictionary")
    gCanon.CompareMode = 0
    Set cnt = CreateObject("Scripting.Dictionary")
    For ix = 0 To gN - 1
        If Not gA(ix).Mixed Then
            tu = UCase$(gA(ix).Key)
            If Not gCanon.Exists(tu) Then
                gCanon.Add tu, gA(ix).Key
                cnt.Add tu, gA(ix).Count
            ElseIf gA(ix).Count > cnt(tu) Then
                gCanon(tu) = gA(ix).Key
                cnt(tu) = gA(ix).Count
            End If
        End If
    Next
    For ix = 0 To gN - 1
        If Not gA(ix).Mixed Then
            tu = UCase$(gA(ix).Key)
            If gCanon(tu) <> gA(ix).Key Then gA(ix).AliasOf = gCanon(tu)
        End If
    Next
    ' --- проход 2: разнобой написания (Плк, плк, IOT при IoT)
    For i = 0 To gPN - 1
        If ScanZone(i) Then
            n = Tokenize(gPText(i), toks, st)
            For j = 0 To n - 1
                t = toks(j)
                If Len(t) >= 2 Then
                    tu = UCase$(t)
                    If gCanon.Exists(tu) Then
                        If gCanon(tu) <> t Then
                            If IsAbbrToken(t) Or Len(t) >= 3 Then
                                ix = gIdx(gCanon(tu))
                                gA(ix).VariantCount = gA(ix).VariantCount + 1
                                gA(ix).Variants = gA(ix).Variants & Chr(1) & t
                                AddFix i, st(j), t, gCanon(tu)
                            End If
                        End If
                    End If
                End If
            Next
        End If
    Next
End Sub

Private Sub MergeSources()
    Dim i As Long, k As String, ix As Long, tu As String
    For i = 0 To gLEN - 1
        k = gLEKey(i)
        If gLEIsAbbr(i) And Not gExc.Exists(k) Then
            tu = UCase$(k)
            If gCanon.Exists(tu) Then
                ix = gIdx(gCanon(tu))
                If gCanon(tu) <> k Then gA(ix).ListSpelling = k
            Else
                ix = GetIdx(k)
            End If
            If Not gA(ix).InList Then
                gA(ix).InList = True
                gA(ix).ListDec = gLEDec(i)
                gA(ix).ListIdx = i
            End If
        End If
    Next
    For ix = 0 To gN - 1
        If gDict.Exists(gA(ix).Key) Then gA(ix).DictDec = gDict(gA(ix).Key)
    Next
End Sub

' ---------------------------------------------------------------------
Private Sub AddIssue(ByVal ix As Long, ByVal sev As Long, ByVal txt As String)
    If Len(gA(ix).Issues) > 0 Then gA(ix).Issues = gA(ix).Issues & Chr(1)
    gA(ix).Issues = gA(ix).Issues & txt
    If sev > gA(ix).Severity Then gA(ix).Severity = sev
End Sub

Private Sub AddStruct(ByVal sev As Long, ByVal txt As String)
    If gSN = 0 Then
        ReDim gSText(0 To 15): ReDim gSSev(0 To 15)
    ElseIf gSN > UBound(gSText) Then
        ReDim Preserve gSText(0 To gSN * 2): ReDim Preserve gSSev(0 To gSN * 2)
    End If
    gSText(gSN) = txt: gSSev(gSN) = sev
    gSN = gSN + 1
End Sub

Public Function OccRange(ByVal pIdx As Long, ByVal off As Long, ByVal tok As String) As Range
    Dim r As Range, k As Long, ordn As Long, pos As Long, ok As Boolean
    On Error Resume Next
    Set r = gDoc.Range(gPStart(pIdx) + off - 1, gPStart(pIdx) + off - 1 + Len(tok))
    ok = (r.Text = tok)
    If Err.Number <> 0 Then ok = False
    Err.Clear
    On Error GoTo 0
    If ok Then
        Set OccRange = r
        Exit Function
    End If
    ' смещение «уехало» из-за полей — ищем нужное по счёту вхождение в абзаце
    ordn = 1
    pos = InStr(1, gPText(pIdx), tok, vbBinaryCompare)
    Do While pos > 0 And pos < off
        ordn = ordn + 1
        pos = InStr(pos + 1, gPText(pIdx), tok, vbBinaryCompare)
    Loop
    Set r = gDoc.Range(gPStart(pIdx), gPEnd(pIdx))
    With r.Find
        .ClearFormatting
        .Text = tok
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
    End With
    k = 0
    Do While r.Find.Execute
        If r.End > gPEnd(pIdx) Then Exit Do
        k = k + 1
        If k = ordn Then
            Set OccRange = r
            Exit Function
        End If
        r.Collapse wdCollapseEnd
    Loop
    Set OccRange = gDoc.Range(gPStart(pIdx), gPStart(pIdx))
End Function

Public Function PageOf(ByVal pIdx As Long, ByVal off As Long, ByVal tok As String) As Long
    Dim r As Range
    If pIdx < 0 Then Exit Function
    On Error Resume Next
    Set r = OccRange(pIdx, off, tok)
    If r Is Nothing Then Exit Function
    PageOf = r.Information(wdActiveEndAdjustedPageNumber)
End Function

Public Function PagePara(ByVal pIdx As Long) As Long
    If pIdx < 0 Then Exit Function
    On Error Resume Next
    PagePara = gDoc.Range(gPStart(pIdx), gPStart(pIdx)).Information(wdActiveEndAdjustedPageNumber)
End Function

Private Function SpellingSummary(ByVal lst As String) As String
    Dim parts As Variant, d As Object, i As Long, k As Variant, s As String
    Set d = CreateObject("Scripting.Dictionary")
    parts = Split(Mid$(lst, 2), Chr(1))
    For i = 0 To UBound(parts)
        If d.Exists(parts(i)) Then d(parts(i)) = d(parts(i)) + 1 Else d.Add parts(i), 1
    Next
    For Each k In d.Keys
        If Len(s) > 0 Then s = s & ", "
        s = s & Q(CStr(k)) & " " & ChrW(&HD7) & d(k)
    Next
    SpellingSummary = s
End Function

Private Sub Evaluate()
    Dim ix As Long, strict As Boolean, listExists As Boolean, pFirst As Long, pDec As Long
    Dim fp As Long, fo As Long, parts As Variant, nMissing As Long, i As Long, nPunct As Long
    Dim prevKey As String, unsorted As String, nHyph As Long, seen As Object, dups As String
    Dim gost As Long, expTitle As String, symCount As Long
    strict = OptStrict()
    gost = GostIdx()
    listExists = (gListHead >= 0)
    gSN = 0
    gNeeded = 0
    For ix = 0 To gN - 1
        gA(ix).Issues = "": gA(ix).Severity = 0: gA(ix).Page = 0
        If Len(gA(ix).AliasOf) > 0 Then GoTo nextIx
        If gA(ix).Mixed Then
            If Len(gA(ix).MixTarget) > 0 Then
                AddIssue ix, 3, "В слове смешаны латинские и русские буквы. Правильно: " & Q(gA(ix).MixTarget)
            Else
                AddIssue ix, 3, "В слове смешаны латинские и русские буквы"
            End If
            GoTo nextIx
        End If
        If gA(ix).Count = 0 Then
            If gA(ix).InList Then AddIssue ix, 2, "Есть в перечне, но в тексте не встречается " & Dash() & " строку нужно удалить"
            GoTo nextIx
        End If
        gNeeded = gNeeded + 1
        With gA(ix)
            If .DecCount = 0 And Not .InList Then
                If Len(.DictDec) > 0 Then
                    AddIssue ix, 3, "Не расшифровано ни в тексте, ни в перечне (в словаре есть: " & Q(.DictDec) & ")"
                Else
                    AddIssue ix, 3, "Не расшифровано ни в тексте, ни в перечне"
                End If
            ElseIf Not .DecAtFirst And .FirstPara >= 0 Then
                If .DecCount > 0 Then
                    If .InList Then
                        AddIssue ix, 2, "Расшифровка дана не при первом упоминании, а позже"
                    Else
                        AddIssue ix, 3, "Расшифровка дана не при первом упоминании, а позже"
                    End If
                ElseIf strict Then
                    AddIssue ix, 2, "Нет расшифровки при первом упоминании (есть только в перечне). Её требует большинство нормоконтролёров"
                Else
                    AddIssue ix, 1, "Расшифровка есть только в перечне; при первом упоминании в тексте её можно добавить"
                End If
            End If
            If .DecCount > 1 Then AddIssue ix, 1, "Расшифровано повторно (" & .DecCount & " раза) " & Dash() & " достаточно одного раза при первом упоминании"
            If InStr(.DecVariants, Chr(1)) > 0 Then
                parts = Split(.DecVariants, Chr(1))
                AddIssue ix, 2, "В тексте разные расшифровки: " & Q(CStr(parts(0))) & " и " & Q(CStr(parts(1)))
            End If
            If .VariantCount > 0 Then AddIssue ix, 2, "Разное написание: " & SpellingSummary(.Variants) & " " & Dash() & " должно быть везде " & Q(.Key)
            If Len(.ListSpelling) > 0 Then AddIssue ix, 2, "В перечне написано " & Q(.ListSpelling) & ", а в тексте " & Q(.Key)
            If listExists And Not .InList Then
                AddIssue ix, 3, "Отсутствует в перечне сокращений"
                nMissing = nMissing + 1
            End If
            If .InList And .ListIdx >= 0 Then
                If gLEPunct(.ListIdx) Then AddIssue ix, 2, "В перечне в конце строки стоит знак препинания (ГОСТ 7.32-2017, п. 6.15)"
            End If
            If .InList And .DecFromA And Not .DecUncertain And Len(.ListDec) > 0 Then
                If Not SameDec(.DecText, .ListDec) Then
                    AddIssue ix, 1, "Расшифровка в тексте (" & Q(.DecText) & ") отличается от перечня"
                End If
            End If
            If .Count = 1 Then AddIssue ix, 1, "Употреблено один раз " & Dash() & " возможно, проще написать полностью"
        End With
nextIx:
    Next
    ' обозначения (не аббревиатуры) в перечне тоже требуют пояснения
    For i = 0 To gLEN - 1
        If Not gLEIsAbbr(i) Then symCount = symCount + 1
    Next
    gNeeded = gNeeded + symCount

    ' --- проверки перечня как структурного элемента
    expTitle = ListTitleFor()
    If Not listExists Then
        If gost = 0 Then
            If gNeeded > 3 Then
                AddStruct 3, "Нет структурного элемента " & Q(expTitle) & ", а сокращений и обозначений, требующих пояснения, " & gNeeded & ". По ГОСТ 7.32-2017 (п. 5.6.2) при числе больше трёх перечень обязателен. Кнопка " & Q("Сформировать перечень") & " создаст его"
            ElseIf gNeeded > 0 Then
                AddStruct 1, "Сокращений немного (" & gNeeded & "). По ГОСТ 7.32-2017 (п. 5.6.3) при числе меньше трёх перечень не составляют, а поясняют их в тексте при первом упоминании"
            End If
        ElseIf gost = 1 Then
            If gNeeded > 0 Then AddStruct 2, "Нет элемента " & Q(expTitle) & " (ГОСТ Р 7.0.11-2011, раздел 4). Сокращений, требующих пояснения: " & gNeeded
        Else
            If gNeeded > 0 Then AddStruct 2, "Нет элемента " & Q(expTitle) & " (ГОСТ Р 2.105-2019, п. 6.15.2). Сокращений, требующих пояснения: " & gNeeded
        End If
    Else
        If NormHead(gListTitle) <> NormHead(expTitle) Then
            If gost = 0 And NormHead(gListTitle) = "ОПРЕДЕЛЕНИЯ, ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ" Then
                ' допустимо по п. 5.6.2
            ElseIf gost = 0 Then
                AddStruct 2, "Заголовок перечня " & Q(gListTitle) & ". По ГОСТ 7.32-2017 (п. 6.2.1) он должен быть " & Q(expTitle)
            Else
                AddStruct 1, "Заголовок перечня " & Q(gListTitle) & ". Для " & GostName() & " обычно пишут " & Q(expTitle)
            End If
        End If
        If gost = 0 And gListIntro < 0 Then
            AddStruct 2, "Перечень должен начинаться со слов " & Q("В настоящем отчете о НИР применяют следующие сокращения и обозначения") & " (ГОСТ 7.32-2017, п. 5.6.1)"
        End If
        If gost = 0 And gNeeded < 3 Then
            AddStruct 1, "Сокращений меньше трёх " & Dash() & " по п. 5.6.3 ГОСТ 7.32-2017 перечень можно не составлять"
        End If
        ' алфавитный порядок, знаки препинания, тире, повторы
        Set seen = CreateObject("Scripting.Dictionary")
        prevKey = ""
        For i = 0 To gLEN - 1
            If gLEIsAbbr(i) Then
                If Len(prevKey) > 0 And Len(unsorted) = 0 Then
                    If StrComp(SortKey(gLEKey(i)), SortKey(prevKey), vbBinaryCompare) < 0 Then
                        unsorted = Q(prevKey) & " стоит перед " & Q(gLEKey(i))
                    End If
                End If
                prevKey = gLEKey(i)
            End If
            If gLEPunct(i) Then nPunct = nPunct + 1
            If gLESep(i) = 2 Then nHyph = nHyph + 1
            If seen.Exists(gLEKey(i)) Then
                If Len(dups) > 0 Then dups = dups & ", "
                dups = dups & Q(gLEKey(i))
            Else
                seen.Add gLEKey(i), 1
            End If
        Next
        If Len(unsorted) > 0 Then AddStruct 2, "Перечень не в алфавитном порядке: " & unsorted & " (ГОСТ 7.32-2017, п. 6.15)"
        If nPunct > 0 Then AddStruct 2, "В " & nPunct & " строк(ах) перечня в конце стоит знак препинания " & Dash() & " по п. 6.15 ГОСТ 7.32-2017 их не ставят"
        If nHyph > 0 Then AddStruct 1, "В " & nHyph & " строк(ах) вместо тире стоит дефис"
        If Len(dups) > 0 Then AddStruct 2, "Повторяются строки перечня: " & dups
        If nMissing > 0 Then AddStruct 3, "В перечне не хватает " & nMissing & " сокращ. из текста"
        ' положение в документе
        If gost = 0 Then
            If gParaIntro >= 0 And gListHead > gParaIntro Then AddStruct 2, "Перечень стоит после введения. По ГОСТ 7.32-2017 (раздел 4) он идёт перед введением, после " & Q("Термины и определения")
            If gParaTerms >= 0 And gListHead < gParaTerms Then AddStruct 1, "По ГОСТ 7.32-2017 (раздел 4) перечень идёт после элемента " & Q("Термины и определения")
        ElseIf gost = 1 Then
            If gParaConcl >= 0 And gListHead < gParaConcl Then AddStruct 2, "По ГОСТ Р 7.0.11-2011 (раздел 4) список сокращений располагают после заключения"
        Else
            If gParaTerms >= 0 And gListHead > gParaTerms Then AddStruct 2, "По ГОСТ Р 2.105-2019 (п. 6.15.2) элемент " & Q("Обозначения и сокращения") & " помещают перед " & Q("Термины и определения")
        End If
    End If
End Sub

Public Function CountSev(ByVal sev As Long) As Long
    Dim ix As Long, i As Long, c As Long
    For ix = 0 To gN - 1
        If Len(gA(ix).AliasOf) = 0 Then
            If gA(ix).Severity = sev And (gA(ix).Count > 0 Or gA(ix).InList) Then c = c + 1
        End If
    Next
    For i = 0 To gSN - 1
        If gSSev(i) = sev Then c = c + 1
    Next
    CountSev = c
End Function

Public Function VisibleCount() As Long
    Dim ix As Long, c As Long
    For ix = 0 To gN - 1
        If Len(gA(ix).AliasOf) = 0 And (gA(ix).Count > 0 Or gA(ix).InList) Then c = c + 1
    Next
    VisibleCount = c
End Function

' Лучшая расшифровка для перечня и исправлений. review = нужна проверка человеком
Public Function BestDec(ByVal ix As Long, ByRef review As Boolean, ByRef src As String) As String
    review = False
    If Not gOverride Is Nothing Then
        If gOverride.Exists(gA(ix).Key) Then
            BestDec = gOverride(gA(ix).Key): src = "ваша правка"
            Exit Function
        End If
    End If
    With gA(ix)
        If Len(.ListDec) > 0 Then
            BestDec = .ListDec: src = "перечень"
        ElseIf Len(.DictDec) > 0 Then
            BestDec = .DictDec: src = "словарь"
        ElseIf Len(.DecText) > 0 And .DecFromA And Not .DecUncertain Then
            BestDec = .DecText: src = "текст"
        ElseIf Len(.DecText) > 0 Then
            BestDec = .DecText: src = "текст (?)": review = True
        Else
            BestDec = "": src = Dash(): review = True
        End If
    End With
End Function

Public Function SevLabel(ByVal sev As Long) As String
    Select Case sev
        Case 3: SevLabel = "Ошибка"
        Case 2: SevLabel = "Внимание"
        Case 1: SevLabel = "Совет"
        Case Else: SevLabel = "ОК"
    End Select
End Function

Public Function SevMark(ByVal sev As Long) As String
    Select Case sev
        Case 3: SevMark = ChrW(&H2716)
        Case 2: SevMark = ChrW(&H25B2)
        Case 1: SevMark = ChrW(&H2022)
        Case Else: SevMark = ChrW(&H2714)
    End Select
End Function

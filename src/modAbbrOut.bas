Attribute VB_Name = "modAbbrOut"
Option Explicit
' =====================================================================
'  Сокращения ГОСТ — примечания, отчёт, исправления, перечень
' =====================================================================

' ---------- примечания в документе ----------
Public Sub ClearOurComments(ByVal doc As Document)
    Dim i As Long, t As String
    For i = doc.Comments.Count To 1 Step -1
        t = ""
        On Error Resume Next
        t = doc.Comments(i).Range.Text
        On Error GoTo 0
        If Left$(t, Len(TagC())) = TagC() Then doc.Comments(i).Delete
    Next
End Sub

Private Function IssuesText(ByVal s As String) As String
    IssuesText = Replace(s, Chr(1), vbCr & ChrW(&H2022) & " ")
End Function

Public Function AddComments() As Long
    Dim ix As Long, n As Long, i As Long, j As Long, pos() As Long, txt() As String
    Dim rs() As Range, tmpP As Long, tmpT As String, tmpR As Range, fp As Long, fo As Long
    Dim r As Range, s As String, anchor As Long
    ReDim pos(0 To gN + 2): ReDim txt(0 To gN + 2): ReDim rs(0 To gN + 2)
    n = 0
    For ix = 0 To gN - 1
        With gA(ix)
            If Len(.AliasOf) = 0 And .Severity >= 2 And .Count > 0 Then
                fp = .FirstPara: fo = .FirstOff
                If fp < 0 Then fp = .FirstAnyPara: fo = .FirstAnyOff
                If fp >= 0 Then
                    Set rs(n) = OccRange(fp, fo, .Key)
                    pos(n) = rs(n).Start
                    txt(n) = TagC() & Q(.Key) & ": " & IssuesText(.Issues)
                    n = n + 1
                End If
            ElseIf Len(.AliasOf) = 0 And .Severity >= 2 And .InList And .ListIdx >= 0 Then
                Set rs(n) = gDoc.Range(gPStart(gLEPara(.ListIdx)), gPEnd(gLEPara(.ListIdx)) - 1)
                pos(n) = rs(n).Start
                txt(n) = TagC() & Q(.Key) & ": " & IssuesText(.Issues)
                n = n + 1
            End If
        End With
    Next
    ' замечания к перечню целиком
    s = ""
    For i = 0 To gSN - 1
        If gSSev(i) >= 2 Then s = s & vbCr & ChrW(&H2022) & " " & gSText(i)
    Next
    If Len(s) > 0 Then
        anchor = gListHead
        If anchor < 0 Then anchor = ListAnchor()
        If anchor < 0 Then anchor = 0
        Set rs(n) = gDoc.Range(gPStart(anchor), gPEnd(anchor) - 1)
        pos(n) = rs(n).Start
        txt(n) = TagC() & "Перечень сокращений:" & s
        n = n + 1
    End If
    ' добавляем с конца документа, чтобы позиции не сдвигались
    For i = 1 To n - 1
        tmpP = pos(i): tmpT = txt(i): Set tmpR = rs(i)
        j = i - 1
        Do While j >= 0
            If pos(j) >= tmpP Then Exit Do
            pos(j + 1) = pos(j): txt(j + 1) = txt(j): Set rs(j + 1) = rs(j)
            j = j - 1
        Loop
        pos(j + 1) = tmpP: txt(j + 1) = tmpT: Set rs(j + 1) = tmpR
    Next
    For i = 0 To n - 1
        On Error Resume Next
        gDoc.Comments.Add rs(i), txt(i)
        On Error GoTo 0
    Next
    AddComments = n
End Function

' ---------- отчёт ----------
Private Function AddPara(ByVal d As Document, ByVal t As String, Optional ByVal sz As Single = 11, _
        Optional ByVal bold As Boolean = False, Optional ByVal col As Long = -1) As Range
    Dim p As Range
    Set p = d.Content.Paragraphs.Last.Range
    If Len(p.Text) > 1 Then
        d.Content.InsertParagraphAfter
        Set p = d.Content.Paragraphs.Last.Range
    End If
    p.InsertBefore t
    p.Font.Size = sz
    p.Font.Bold = bold
    If col >= 0 Then p.Font.Color = col Else p.Font.Color = wdColorAutomatic
    p.ParagraphFormat.SpaceAfter = 4
    p.ParagraphFormat.SpaceBefore = 0
    p.ParagraphFormat.FirstLineIndent = 0
    p.ParagraphFormat.LeftIndent = 0
    p.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Set AddPara = p
End Function

Private Function SevColor(ByVal sev As Long) As Long
    Select Case sev
        Case 3: SevColor = RGB(255, 199, 206)
        Case 2: SevColor = RGB(255, 235, 156)
        Case 1: SevColor = RGB(221, 235, 247)
        Case Else: SevColor = RGB(198, 239, 206)
    End Select
End Function

Private Function Cln(ByVal s As String) As String
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Cln = s
End Function

Public Sub SortedIdx(ByRef ord() As Long, ByRef n As Long)
    Dim ix As Long, i As Long, j As Long, t As Long, kt As String, st As Long
    Dim keys() As String
    n = 0
    ReDim ord(0 To gN): ReDim keys(0 To gN)
    For ix = 0 To gN - 1
        If Len(gA(ix).AliasOf) = 0 And (gA(ix).Count > 0 Or gA(ix).InList) Then
            ord(n) = ix
            keys(n) = CStr(9 - gA(ix).Severity) & SortKey(gA(ix).Key)
            n = n + 1
        End If
    Next
    For i = 1 To n - 1
        t = ord(i): kt = keys(i)
        j = i - 1
        Do While j >= 0
            If StrComp(keys(j), kt, vbBinaryCompare) <= 0 Then Exit Do
            ord(j + 1) = ord(j): keys(j + 1) = keys(j)
            j = j - 1
        Loop
        ord(j + 1) = t: keys(j + 1) = kt
    Next
End Sub

Public Sub BuildReport()
    Dim rep As Document, ord() As Long, n As Long, i As Long, ix As Long, block As String
    Dim review As Boolean, src As String, dec As String, tb As Table, p As Range, k As Long
    Dim nE As Long, nW As Long, nI As Long, srcName As String, sevs() As Long
    SortedIdx ord, n
    nE = CountSev(3): nW = CountSev(2): nI = CountSev(1)
    srcName = gDoc.Name
    Application.ScreenUpdating = False
    Set rep = Documents.Add
    On Error Resume Next
    rep.Variables.Add "AbbrGOST_Report", "1"
    On Error GoTo 0
    With rep.PageSetup
        .Orientation = wdOrientLandscape
        .LeftMargin = CentimetersToPoints(1.5)
        .RightMargin = CentimetersToPoints(1.5)
        .TopMargin = CentimetersToPoints(1.5)
        .BottomMargin = CentimetersToPoints(1.5)
    End With
    AddPara rep, "Сокращения и обозначения: отчёт о проверке", 18, True, RGB(31, 56, 100)
    AddPara rep, "Документ: " & srcName & "    Стандарт: " & GostName() & "    " & Format$(Now, "dd.mm.yyyy hh:nn"), 10, False, RGB(89, 89, 89)
    AddPara rep, "Найдено сокращений: " & VisibleCount() & "    " & ChrW(&H2716) & " ошибок: " & nE & "    " & ChrW(&H25B2) & " внимание: " & nW & "    " & ChrW(&H2022) & " советов: " & nI, 12, True

    AddPara rep, "Перечень сокращений как элемент документа", 13, True, RGB(31, 56, 100)
    If gListHead >= 0 Then
        AddPara rep, "Найден: " & Q(gListTitle) & ", стр. " & PagePara(gListHead) & ", строк: " & gLEN, 11
    Else
        AddPara rep, "Перечень в документе не найден.", 11
    End If
    If gSN = 0 Then
        AddPara rep, ChrW(&H2714) & " Замечаний к перечню нет.", 11
    Else
        For i = 0 To gSN - 1
            Set p = AddPara(rep, SevMark(gSSev(i)) & " " & SevLabel(gSSev(i)) & ": " & gSText(i), 11)
            p.Shading.BackgroundPatternColor = SevColor(gSSev(i))
        Next
    End If

    AddPara rep, "Сокращения в тексте", 13, True, RGB(31, 56, 100)
    If n = 0 Then
        AddPara rep, "Сокращений, требующих расшифровки, не найдено.", 11
    Else
        block = "Статус" & vbTab & "Сокращение" & vbTab & "Расшифровка" & vbTab & "Откуда" & vbTab & "Употр." & vbTab & "Стр." & vbTab & "Замечания"
        ReDim sevs(0 To n)
        For i = 0 To n - 1
            ix = ord(i)
            If gA(ix).Page = 0 Then
                If gA(ix).FirstPara >= 0 Then
                    gA(ix).Page = PageOf(gA(ix).FirstPara, gA(ix).FirstOff, gA(ix).Key)
                ElseIf gA(ix).FirstAnyPara >= 0 Then
                    gA(ix).Page = PageOf(gA(ix).FirstAnyPara, gA(ix).FirstAnyOff, gA(ix).Key)
                ElseIf gA(ix).ListIdx >= 0 Then
                    gA(ix).Page = PagePara(gLEPara(gA(ix).ListIdx))
                End If
            End If
            sevs(i) = gA(ix).Severity
            dec = BestDec(ix, review, src)
            If gA(ix).Mixed Then dec = "": src = Dash()
            block = block & vbCr & SevMark(gA(ix).Severity) & " " & SevLabel(gA(ix).Severity) & vbTab & _
                Cln(gA(ix).Key) & vbTab & Cln(dec) & vbTab & src & vbTab & gA(ix).Count & vbTab & _
                IIf(gA(ix).Page > 0, CStr(gA(ix).Page), Dash()) & vbTab & Cln(Replace(gA(ix).Issues, Chr(1), "; "))
        Next
        Set p = AddPara(rep, block, 9)
        Set tb = p.ConvertToTable(Separator:=wdSeparateByTabs, NumColumns:=7)
        tb.Borders.Enable = True
        tb.Range.Font.Size = 9
        tb.Range.ParagraphFormat.SpaceAfter = 0
        tb.Rows(1).Range.Font.Bold = True
        tb.Rows(1).HeadingFormat = True
        tb.Rows(1).Shading.BackgroundPatternColor = RGB(217, 226, 243)
        tb.PreferredWidthType = wdPreferredWidthPercent
        tb.PreferredWidth = 100
        On Error Resume Next
        tb.Columns(1).PreferredWidthType = wdPreferredWidthPercent: tb.Columns(1).PreferredWidth = 10
        tb.Columns(2).PreferredWidthType = wdPreferredWidthPercent: tb.Columns(2).PreferredWidth = 10
        tb.Columns(3).PreferredWidthType = wdPreferredWidthPercent: tb.Columns(3).PreferredWidth = 27
        tb.Columns(4).PreferredWidthType = wdPreferredWidthPercent: tb.Columns(4).PreferredWidth = 8
        tb.Columns(5).PreferredWidthType = wdPreferredWidthPercent: tb.Columns(5).PreferredWidth = 6
        tb.Columns(6).PreferredWidthType = wdPreferredWidthPercent: tb.Columns(6).PreferredWidth = 5
        tb.Columns(7).PreferredWidthType = wdPreferredWidthPercent: tb.Columns(7).PreferredWidth = 34
        For k = 2 To tb.Rows.Count
            tb.Cell(k, 1).Shading.BackgroundPatternColor = SevColor(sevs(k - 2))
        Next
        On Error GoTo 0
    End If

    AddPara rep, "Что дальше", 13, True, RGB(31, 56, 100)
    AddPara rep, "1. Примечания " & Q("[Сокращения ГОСТ]") & " стоят в документе у первого упоминания каждого проблемного сокращения (вкладка Рецензирование " & ChrW(&H2192) & " Следующее примечание).", 10
    AddPara rep, "2. Кнопка " & Q("Исправить") & " вставит расшифровки при первом упоминании, исправит разнобой написания и смешение латиницы с кириллицей. Все правки вносятся в режиме исправлений " & Dash() & " их можно принять или отклонить.", 10
    AddPara rep, "3. Кнопка " & Q("Сформировать перечень") & " соберёт перечень по алфавиту в нужном месте документа. Строки, которые нужно проверить (падеж, расшифровка не найдена), будут выделены жёлтым.", 10
    AddPara rep, "4. Нет расшифровки? Выделите сокращение и нажмите " & Q("В словарь") & " " & Dash() & " в следующих документах она подставится сама.", 10
    AddPara rep, "Откуда: перечень " & Dash() & " из перечня документа; текст " & Dash() & " найдена в тексте; текст (?) " & Dash() & " распознана неуверенно, проверьте; словарь " & Dash() & " из личного словаря.", 9, False, RGB(89, 89, 89)
    rep.Saved = True
    Application.ScreenUpdating = True
    rep.Activate
End Sub

' ---------- место для перечня по выбранному ГОСТ ----------
Public Function ListAnchor() As Long
    Dim i As Long
    ListAnchor = -1
    Select Case GostIdx()
        Case 1
            If gParaConcl >= 0 Then
                For i = gParaConcl + 1 To gPN - 1
                    If gPZone(i) = Z_HEAD Then ListAnchor = i: Exit Function
                Next
            End If
            If gParaBib >= 0 Then ListAnchor = gParaBib
        Case 2
            If gParaTerms >= 0 Then ListAnchor = gParaTerms: Exit Function
            If gParaIntro >= 0 Then ListAnchor = gParaIntro: Exit Function
            For i = 0 To gPN - 1
                If gPZone(i) = Z_BODY And gPHead(i) Then ListAnchor = i: Exit Function
            Next
        Case Else
            If gParaIntro >= 0 Then ListAnchor = gParaIntro
    End Select
End Function

' ---------- исправления ----------
' Все функции возвращают текст для строки состояния панели; окон не показывают.

Private Sub SortDesc(sS() As Long, sE() As Long, sTxt() As String, sKind() As Long, ByVal n As Long)
    Dim i As Long, j As Long, tS As Long, tE As Long, tT As String, tK As Long
    For i = 1 To n - 1
        tS = sS(i): tE = sE(i): tT = sTxt(i): tK = sKind(i)
        j = i - 1
        Do While j >= 0
            If sS(j) >= tS Then Exit Do
            sS(j + 1) = sS(j): sE(j + 1) = sE(j): sTxt(j + 1) = sTxt(j): sKind(j + 1) = sKind(j)
            j = j - 1
        Loop
        sS(j + 1) = tS: sE(j + 1) = tE: sTxt(j + 1) = tT: sKind(j + 1) = tK
    Next
End Sub

Private Sub ApplyOps(sS() As Long, sE() As Long, sTxt() As String, sKind() As Long, ByVal n As Long, ByVal undoName As String)
    Dim i As Long, oldTrack As Boolean
    SortDesc sS, sE, sTxt, sKind, n
    BeginUndo undoName
    oldTrack = gDoc.TrackRevisions
    gDoc.TrackRevisions = True
    Application.ScreenUpdating = False
    For i = 0 To n - 1
        If sKind(i) = 1 Then
            gDoc.Range(sS(i), sS(i)).InsertAfter sTxt(i)
        Else
            gDoc.Range(sS(i), sE(i)).Text = sTxt(i)
        End If
    Next
    gDoc.TrackRevisions = oldTrack
    Application.ScreenUpdating = True
    EndUndo
End Sub

' Исправить всё: расшифровки при первом упоминании + написание
Public Function FixAll() As String
    Dim ix As Long, n As Long, i As Long, r As Range, review As Boolean, src As String
    Dim dec As String, nIns As Long, nRep As Long, noDec As String, strict As Boolean
    Dim sS() As Long, sE() As Long, sTxt() As String, sKind() As Long
    strict = OptStrict()
    ReDim sS(0 To gN + gFixN + 1): ReDim sE(0 To gN + gFixN + 1)
    ReDim sTxt(0 To gN + gFixN + 1): ReDim sKind(0 To gN + gFixN + 1)
    For ix = 0 To gN - 1
        With gA(ix)
            If Len(.AliasOf) = 0 And Not .Mixed And .Count > 0 And .DecCount = 0 And .FirstPara >= 0 Then
                If Not .InList Or strict Then
                    dec = BestDec(ix, review, src)
                    If Len(dec) > 0 And Not review Then
                        Set r = OccRange(.FirstPara, .FirstOff, .Key)
                        sS(n) = r.End: sE(n) = r.End
                        sTxt(n) = " (" & dec & ")": sKind(n) = 1
                        n = n + 1: nIns = nIns + 1
                    Else
                        If Len(noDec) > 0 Then noDec = noDec & ", "
                        noDec = noDec & .Key
                    End If
                End If
            End If
        End With
    Next
    For i = 0 To gFixN - 1
        Set r = OccRange(gFixP(i), gFixO(i), gFixT(i))
        If r.Text = gFixT(i) Then
            sS(n) = r.Start: sE(n) = r.End
            sTxt(n) = gFixNew(i): sKind(n) = 2
            n = n + 1: nRep = nRep + 1
        End If
    Next
    If n = 0 Then
        FixAll = "Автоматически исправлять нечего."
    Else
        ApplyOps sS, sE, sTxt, sKind, n, "Сокращения ГОСТ: исправить всё"
        FixAll = "Внесено в режиме исправлений: расшифровок " & nIns & ", исправлений написания " & nRep & ". Отменить: Ctrl+Z."
    End If
    If Len(noDec) > 0 Then FixAll = FixAll & " Нет надёжной расшифровки: " & noDec & " " & Dash() & " впишите её в карточке."
End Function

' Вставить расшифровку после первого упоминания одного сокращения
Public Function InsertDecoding(ByVal key As String, ByVal dec As String) As String
    Dim ix As Long, r As Range, fp As Long, fo As Long
    Dim sS(0) As Long, sE(0) As Long, sTxt(0) As String, sKind(0) As Long
    dec = StripEndPunct(dec)
    If Len(dec) = 0 Then InsertDecoding = "Впишите расшифровку в поле над кнопками.": Exit Function
    If Not gIdx.Exists(key) Then InsertDecoding = Q(key) & " больше не встречается в тексте.": Exit Function
    ix = gIdx(key)
    fp = gA(ix).FirstPara: fo = gA(ix).FirstOff
    If fp < 0 Then fp = gA(ix).FirstAnyPara: fo = gA(ix).FirstAnyOff
    If fp < 0 Then InsertDecoding = Q(key) & " не найдено в тексте.": Exit Function
    Set r = OccRange(fp, fo, key)
    sS(0) = r.End: sE(0) = r.End: sTxt(0) = " (" & dec & ")": sKind(0) = 1
    ApplyOps sS, sE, sTxt, sKind, 1, "Сокращения ГОСТ: расшифровка " & key
    InsertDecoding = "Вставлено после первого упоминания " & Q(key) & ": (" & dec & "). Отменить: Ctrl+Z."
End Function

' Исправить написание: для канонического ключа — все варианты, для «смешанного» — его вхождения
Public Function FixSpelling(ByVal key As String) As String
    Dim i As Long, n As Long, r As Range
    Dim sS() As Long, sE() As Long, sTxt() As String, sKind() As Long
    ReDim sS(0 To gFixN + 1): ReDim sE(0 To gFixN + 1): ReDim sTxt(0 To gFixN + 1): ReDim sKind(0 To gFixN + 1)
    For i = 0 To gFixN - 1
        If gFixNew(i) = key Or gFixT(i) = key Then
            Set r = OccRange(gFixP(i), gFixO(i), gFixT(i))
            If r.Text = gFixT(i) Then
                sS(n) = r.Start: sE(n) = r.End: sTxt(n) = gFixNew(i): sKind(n) = 2
                n = n + 1
            End If
        End If
    Next
    If n = 0 Then FixSpelling = "Исправлять в написании " & Q(key) & " нечего.": Exit Function
    ApplyOps sS, sE, sTxt, sKind, n, "Сокращения ГОСТ: написание " & key
    FixSpelling = "Исправлено вхождений: " & n & ". Отменить: Ctrl+Z."
End Function

' Где окажется перечень
Public Function ListPlaceText() As String
    Dim a As Long
    If gListHead >= 0 Then
        ListPlaceText = "Найден " & Q(gListTitle) & " (стр. " & PagePara(gListHead) & "), строк: " & gLEN & ". Кнопка ниже пересоберёт его на месте."
    Else
        a = ListAnchor()
        If a >= 0 Then
            ListPlaceText = "Перечня нет. По " & GostName() & " он будет вставлен перед " & Q(NormSpaces(gPText(a))) & " (стр. " & PagePara(a) & ")."
        Else
            ListPlaceText = "Перечня нет, и место по " & GostName() & " не найдено " & Dash() & " он будет вставлен в позицию курсора."
        End If
    End If
End Function

Public Function StructText() As String
    Dim i As Long, s As String
    If gSN = 0 Then StructText = ChrW(&H2714) & " Замечаний к перечню нет.": Exit Function
    For i = 0 To gSN - 1
        If Len(s) > 0 Then s = s & vbLf
        s = s & ChrW(&H2022) & " " & gSText(i)
    Next
    StructText = s
End Function

Private Sub BeginUndo(ByVal t As String)
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord t
End Sub

Private Sub EndUndo()
    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
End Sub

' ---------- обозначения, хранящиеся в документе ----------
Public Function GetSymbols(ByVal doc As Document) As Object
    Dim d As Object, s As String, lines As Variant, i As Long, p As Long
    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    s = doc.Variables("AbbrGOST_Symbols").Value
    On Error GoTo 0
    If Len(s) > 0 Then
        lines = Split(s, vbLf)
        For i = 0 To UBound(lines)
            p = InStr(lines(i), vbTab)
            If p > 1 Then
                If Not d.Exists(Left$(lines(i), p - 1)) Then d.Add Left$(lines(i), p - 1), Mid$(lines(i), p + 1)
            End If
        Next
    End If
    Set GetSymbols = d
End Function

Public Sub SaveSymbols(ByVal doc As Document, ByVal d As Object)
    Dim k As Variant, s As String
    For Each k In d.Keys
        If Len(s) > 0 Then s = s & vbLf
        s = s & k & vbTab & d(k)
    Next
    On Error Resume Next
    doc.Variables("AbbrGOST_Symbols").Delete
    If Len(s) > 0 Then doc.Variables.Add "AbbrGOST_Symbols", s
End Sub

' ---------- перечень ----------
' Собрать строки будущего перечня. dropUnused — убрать сокращения, которых нет в тексте.
Public Function BuildListEntries(keys() As String, decs() As String, rev() As Boolean, ByVal dropUnused As Boolean) As Long
    Dim n As Long, ix As Long, i As Long, j As Long, review As Boolean, src As String, dec As String
    Dim syms As Object, k As Variant, have As Object, sk() As String, d1 As String
    Dim tk As String, td As String, tr As Boolean, ts As String
    ReDim keys(0 To gN + gLEN + 64): ReDim decs(0 To gN + gLEN + 64): ReDim rev(0 To gN + gLEN + 64)
    Set have = CreateObject("Scripting.Dictionary")
    For ix = 0 To gN - 1
        With gA(ix)
            If Len(.AliasOf) = 0 And Not .Mixed Then
                If .Count > 0 Or (.InList And Not dropUnused) Then
                    dec = BestDec(ix, review, src)
                    If Len(dec) = 0 Then dec = "[расшифровка]"
                    If review And Len(dec) > 1 Then
                        d1 = Mid$(dec, 2, 1)
                        If IsLowerL(AscW(d1)) Then dec = LCase$(Left$(dec, 1)) & Mid$(dec, 2)
                    End If
                    keys(n) = .Key: decs(n) = dec: rev(n) = review
                    have.Add .Key, 1
                    n = n + 1
                End If
            End If
        End With
    Next
    For i = 0 To gLEN - 1
        If Not gLEIsAbbr(i) And Not have.Exists(gLEKey(i)) Then
            keys(n) = gLEKey(i): decs(n) = gLEDec(i): rev(n) = False
            have.Add gLEKey(i), 1
            n = n + 1
        End If
    Next
    Set syms = GetSymbols(gDoc)
    For Each k In syms.Keys
        If Not have.Exists(CStr(k)) And n <= UBound(keys) Then
            keys(n) = CStr(k): decs(n) = syms(k): rev(n) = False
            have.Add CStr(k), 1
            n = n + 1
        End If
    Next
    If n = 0 Then BuildListEntries = 0: Exit Function
    ReDim sk(0 To n)
    For i = 0 To n - 1
        sk(i) = SortKey(keys(i))
        If Not IsAbbrToken(keys(i)) Then sk(i) = "5" & sk(i)
    Next
    For i = 1 To n - 1
        tk = keys(i): td = decs(i): tr = rev(i): ts = sk(i)
        j = i - 1
        Do While j >= 0
            If StrComp(sk(j), ts, vbBinaryCompare) <= 0 Then Exit Do
            keys(j + 1) = keys(j): decs(j + 1) = decs(j): rev(j + 1) = rev(j): sk(j + 1) = sk(j)
            j = j - 1
        Loop
        keys(j + 1) = tk: decs(j + 1) = td: rev(j + 1) = tr: sk(j + 1) = ts
    Next
    BuildListEntries = n
End Function

Public Function UnusedInList() As Long
    Dim ix As Long, c As Long
    For ix = 0 To gN - 1
        If Len(gA(ix).AliasOf) = 0 And Not gA(ix).Mixed And gA(ix).Count = 0 And gA(ix).InList Then c = c + 1
    Next
    UnusedInList = c
End Function

' Записать перечень в документ: на место старого или по ГОСТ / в позицию курсора
Public Function WriteList(ByVal dropUnused As Boolean) As String
    Dim keys() As String, decs() As String, rev() As Boolean, n As Long, i As Long, nRev As Long, where As String
    n = BuildListEntries(keys, decs, rev, dropUnused)
    If n = 0 Then WriteList = "В документе нет сокращений и обозначений для перечня.": Exit Function
    For i = 0 To n - 1
        If rev(i) Then nRev = nRev + 1
    Next
    If gListHead >= 0 Then
        WriteExistingList keys, decs, rev, n
        where = "Перечень пересобран на месте"
    Else
        where = WriteNewList(keys, decs, rev, n)
        If Len(where) = 0 Then WriteList = "Не удалось вставить перечень: поставьте курсор в документ.": Exit Function
    End If
    WriteList = where & ": " & n & " строк."
    If nRev > 0 Then WriteList = WriteList & " Жёлтым выделены " & nRev & " " & Dash() & " проверьте падеж."
    WriteList = WriteList & " Отменить: Ctrl+Z."
End Function

Private Function ListLines(keys() As String, decs() As String, ByVal n As Long) As String
    Dim i As Long, s As String, sep As String
    If OptTabStyle() Then sep = vbTab & Dash() & " " Else sep = " " & Dash() & " "
    For i = 0 To n - 1
        s = s & keys(i) & sep & StripEndPunct(decs(i)) & vbCr
    Next
    ListLines = s
End Function

Private Sub FormatList(ByVal r As Range, ByVal hasIntro As Boolean, keys() As String, rev() As Boolean, ByVal n As Long, ByVal bodyStyle As Variant)
    Dim p As Paragraph, i As Long, w As Single, maxLen As Long, k As Long
    For i = 0 To n - 1
        If Len(keys(i)) > maxLen Then maxLen = Len(keys(i))
    Next
    w = 0.32 * maxLen + 0.9
    If w < 2 Then w = 2
    If w > 5 Then w = 5
    w = CentimetersToPoints(w)
    On Error Resume Next
    r.Style = bodyStyle
    On Error GoTo 0
    r.Font.Reset
    r.ParagraphFormat.Reset
    r.HighlightColorIndex = wdNoHighlight
    i = 0
    k = 0
    For Each p In r.Paragraphs
        i = i + 1
        p.Format.PageBreakBefore = False
        p.Format.KeepWithNext = False
        On Error Resume Next
        p.Format.OutlineLevel = wdOutlineLevelBodyText
        p.Range.ListFormat.RemoveNumbers
        On Error GoTo 0
        If i = 1 And hasIntro Then
            p.Format.Alignment = wdAlignParagraphJustify
        Else
            p.Format.Alignment = wdAlignParagraphLeft
            If OptTabStyle() Then
                p.Format.LeftIndent = w
                p.Format.FirstLineIndent = -w
                p.TabStops.ClearAll
                p.TabStops.Add w
            Else
                p.Format.LeftIndent = 0
                p.Format.FirstLineIndent = 0
            End If
            If k < n Then
                If rev(k) Then p.Range.HighlightColorIndex = wdYellow
            End If
            k = k + 1
        End If
    Next
End Sub

Private Function BodyStyleRef() As Variant
    Dim i As Long
    If gListIntro >= 0 Then
        On Error Resume Next
        Set BodyStyleRef = gDoc.Range(gPStart(gListIntro), gPStart(gListIntro)).Paragraphs(1).Style
        If Err.Number = 0 Then Exit Function
        Err.Clear
    End If
    If gParaIntro >= 0 Then
        For i = gParaIntro + 1 To gPN - 1
            If gPZone(i) = Z_BODY And Not gPHead(i) And Not gPInTbl(i) And Len(gPText(i)) > 40 Then
                On Error Resume Next
                Set BodyStyleRef = gDoc.Range(gPStart(i), gPStart(i)).Paragraphs(1).Style
                If Err.Number = 0 Then Exit Function
                Err.Clear
                Exit For
            End If
        Next
    End If
    Set BodyStyleRef = gDoc.Styles(wdStyleNormal)
End Function

Private Sub WriteExistingList(keys() As String, decs() As String, rev() As Boolean, ByVal n As Long)
    Dim afterPara As Long, s As Long, e As Long, lastKeep As Long, r As Range, txt As String
    Dim needIntro As Boolean, bs As Variant, rr As Range
    Set bs = BodyStyleRef()
    afterPara = gListHead
    If gListIntro >= 0 Then afterPara = gListIntro
    needIntro = (GostIdx() = 0 And gListIntro < 0)
    s = gPEnd(afterPara)
    e = s
    lastKeep = gListLast
    Do While lastKeep > afterPara
        If InStr(gPText(lastKeep), Chr(12)) > 0 Or Len(Trim$(gPText(lastKeep))) = 0 Then
            lastKeep = lastKeep - 1
        Else
            Exit Do
        End If
    Loop
    If lastKeep > afterPara Then
        e = gPEnd(lastKeep)
        If gPInTbl(lastKeep) Then
            On Error Resume Next
            Set rr = gDoc.Range(gPStart(lastKeep), gPStart(lastKeep))
            If rr.Tables.Count > 0 Then e = rr.Tables(1).Range.End
            On Error GoTo 0
        End If
    End If
    BeginUndo "Сокращения ГОСТ: перечень"
    Application.ScreenUpdating = False
    If e > s Then gDoc.Range(s, e).Delete
    txt = ""
    If needIntro Then txt = "В настоящем отчете о НИР применяют следующие сокращения и обозначения:" & vbCr
    txt = txt & ListLines(keys, decs, n)
    Set r = gDoc.Range(s, s)
    r.InsertAfter txt
    FormatList r, needIntro, keys, rev, n, bs
    Application.ScreenUpdating = True
    EndUndo
End Sub

' Возвращает описание места вставки или "" при неудаче
Private Function WriteNewList(keys() As String, decs() As String, rev() As Boolean, ByVal n As Long) As String
    Dim anchor As Long, pos As Long, r As Range, txt As String, title As String
    Dim hasIntro As Boolean, hs As Variant, bs As Variant, tp As Paragraph, body As Range, nxt As Paragraph
    Dim where As String
    anchor = ListAnchor()
    Set bs = BodyStyleRef()
    If anchor >= 0 Then
        pos = gPStart(anchor)
        Set hs = gDoc.Range(pos, pos).Paragraphs(1).Style
        where = "Перечень вставлен перед " & Q(NormSpaces(gPText(anchor)))
    Else
        If Not Selection.Document Is gDoc Then Exit Function
        pos = Selection.Range.Start
        If pos <> Selection.Range.Paragraphs(1).Range.Start Then
            gDoc.Range(pos, pos).InsertParagraphAfter
            pos = pos + 1
        End If
        Set hs = gDoc.Styles(wdStyleHeading1)
        where = "Перечень вставлен в позицию курсора"
    End If
    title = ListTitleFor()
    hasIntro = (GostIdx() = 0)
    txt = title & vbCr
    If hasIntro Then txt = txt & "В настоящем отчете о НИР применяют следующие сокращения и обозначения:" & vbCr
    txt = txt & ListLines(keys, decs, n)
    BeginUndo "Сокращения ГОСТ: перечень"
    Application.ScreenUpdating = False
    Set r = gDoc.Range(pos, pos)
    r.InsertAfter txt
    Set tp = r.Paragraphs(1)
    Set body = gDoc.Range(tp.Range.End, r.End)
    FormatList body, hasIntro, keys, rev, n, bs
    On Error Resume Next
    tp.Style = hs
    tp.Range.ListFormat.RemoveNumbers
    tp.Format.PageBreakBefore = True
    If GostIdx() <> 2 Then tp.Format.Alignment = wdAlignParagraphCenter
    If anchor >= 0 Then
        Set nxt = gDoc.Range(r.End, r.End).Paragraphs(1)
        If Not nxt.Format.PageBreakBefore Then nxt.Format.PageBreakBefore = True
    End If
    On Error GoTo 0
    Application.ScreenUpdating = True
    EndUndo
    WriteNewList = where
End Function

' ---------- пополнение словаря ----------
Public Function LearnFromDoc() As String
    Dim ix As Long, dec As String, n As Long
    For ix = 0 To gN - 1
        With gA(ix)
            If Len(.AliasOf) = 0 And Not .Mixed And Not gDict.Exists(.Key) Then
                dec = ""
                If Len(.ListDec) > 0 Then
                    dec = .ListDec
                ElseIf .DecFromA And Not .DecUncertain Then
                    dec = .DecText
                End If
                If Len(dec) > 0 Then
                    UpsertKV DictPath(), .Key, dec
                    n = n + 1
                End If
            End If
        End With
    Next
    If n = 0 Then
        LearnFromDoc = "Новых расшифровок для словаря нет."
    Else
        LearnFromDoc = "В словарь добавлено расшифровок: " & n & "."
    End If
End Function

' ---------- справка ----------
Public Sub ShowHelp()
    Dim d As Document
    Set d = Documents.Add
    On Error Resume Next
    d.Variables.Add "AbbrGOST_Report", "1"
    On Error GoTo 0
    AddPara d, "Сокращения ГОСТ " & Dash() & " справка", 18, True, RGB(31, 56, 100)
    AddPara d, "Надстройка находит в документе сокращения и обозначения, проверяет их по выбранному стандарту и помогает привести в порядок перечень.", 11
    AddPara d, "Как работать", 13, True, RGB(31, 56, 100)
    AddPara d, "Проверить " & Dash() & " открывает панель рядом с текстом. Слева статус, затем сокращение и суть замечания. Щелчок по строке переносит курсор к первому упоминанию. Внизу карточка: все замечания, поле расшифровки и кнопки действий.", 11
    AddPara d, "Вставить расшифровку в текст " & Dash() & " добавляет её в скобках после первого упоминания (в режиме исправлений). Исправить написание " & Dash() & " приводит Плк, плк, ПЛK с латинской K к единому виду. В словарь / Не проверять " & Dash() & " запоминает решение для всех документов.", 11
    AddPara d, "Вкладка Перечень показывает перечень до записи в документ: можно поправить любую расшифровку, затем нажать Записать перечень в документ. Существующий перечень пересобирается на месте, новый вставляется туда, где его требует стандарт.", 11
    AddPara d, "Вкладка Обозначения " & Dash() & " символы величин (F, " & ChrW(&H3C9) & ", Vmax): их нельзя найти в формулах автоматически, поэтому они вносятся здесь и хранятся в документе. Вкладка Словарь " & Dash() & " ваши расшифровки и исключения с поиском.", 11
    AddPara d, "Исправить всё " & Dash() & " всё сразу для тех, кто доверяет. Любое действие отменяется через Ctrl+Z.", 11
    AddPara d, "Что проверяется по стандартам", 13, True, RGB(31, 56, 100)
    AddPara d, "ГОСТ 7.32-2017: заголовок " & Q("ПЕРЕЧЕНЬ СОКРАЩЕНИЙ И ОБОЗНАЧЕНИЙ") & " (п. 6.2.1), вводная фраза (п. 5.6.1), обязательность перечня при числе обозначений больше трёх (пп. 5.6.2, 5.6.3), оформление столбцом по алфавиту через тире без знаков препинания в конце (п. 6.15), место перед введением (раздел 4).", 11
    AddPara d, "ГОСТ Р 7.0.11-2011: список сокращений и условных обозначений после заключения (раздел 4).", 11
    AddPara d, "ГОСТ Р 2.105-2019: элемент " & Q("Обозначения и сокращения") & " перед " & Q("Термины и определения") & " (п. 6.15.2), оформление по ГОСТ 7.32 (п. 6.1.2).", 11
    AddPara d, "Расшифровка при первом упоминании и единообразие написания проверяются для всех стандартов. Строгий режим требует расшифровку в тексте даже при наличии перечня " & Dash() & " так делает большинство нормоконтролёров.", 11
    AddPara d, "Чего надстройка не делает", 13, True, RGB(31, 56, 100)
    AddPara d, "Не ставит расшифровку в нужный падеж: в тексте она вставляется в скобках в именительном падеже (это допустимо), а в перечне строки из текста подсвечиваются для проверки. Не видит символы внутри формул " & Dash() & " для них есть кнопка " & Q("Обозначение") & ". Не придумывает расшифровок: берёт их только из документа или словаря.", 11
    AddPara d, "Поддержать автора", 13, True, RGB(31, 56, 100)
    AddPara d, "Надстройка бесплатна и распространяется с открытым кодом (лицензия MIT). Если она сэкономила вам время, можно угостить автора кофе: " & TipUrl(), 11
    AddPara d, "Исходный код и новые версии: https://github.com/shevchenko-robotics/abbr-gost-word", 11
    AddPara d, "Где лежат файлы", 13, True, RGB(31, 56, 100)
    AddPara d, "Словарь: " & DictPath(), 10
    AddPara d, "Исключения: " & ExcPath(), 10
    AddPara d, "Оба файла можно править в Блокноте: одна строка " & Dash() & " одно сокращение, расшифровка через табуляцию или тире.", 10
    d.Saved = True
End Sub

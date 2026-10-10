Attribute VB_Name = "modAbbrMain"
Option Explicit
' =====================================================================
'  Сокращения ГОСТ — лента и команды
'  Надстройка для Microsoft Word (Windows). Версия 2.0
' =====================================================================

Public gRibbon As Object

' ---------- лента ----------
Public Sub AbbrRibbonLoad(ribbon As Object)
    Set gRibbon = ribbon
End Sub

Public Sub RibbonRefresh()
    On Error Resume Next
    If Not gRibbon Is Nothing Then gRibbon.Invalidate
End Sub

Public Sub AbbrOnAction(control As Object)
    Select Case control.ID
        Case "btnCheck": CmdCheck
        Case "btnFix": CmdFix
        Case "btnList": CmdPanelPage 1
        Case "btnAddSym": CmdAddSymbol
        Case "btnAddDict": CmdAddDict
        Case "btnReport": CmdReport
        Case "btnClear": CmdClear
        Case "btnComments": CmdComments
        Case "btnHelp": ShowHelp
        Case "btnTip": OpenTipPage
    End Select
End Sub

Public Sub AbbrGostSel(control As Object, ByRef returnedVal)
    returnedVal = GostIdx()
End Sub

Public Sub AbbrGostSelect(control As Object, id As String, index As Integer)
    SetOpt "Gost", CStr(index)
    If PanelVisible() Then
        If RunScan(gDoc) Then UserForm1.RefreshAll
        UserForm1.SetStatus "Стандарт: " & GostName() & "."
    End If
End Sub

Public Sub AbbrPressed(control As Object, ByRef returnedVal)
    Select Case control.ID
        Case "chkStrict": returnedVal = OptStrict()
        Case Else: returnedVal = False
    End Select
End Sub

Public Sub AbbrCheck(control As Object, pressed As Boolean)
    Select Case control.ID
        Case "chkStrict"
            SetOpt "Strict", IIf(pressed, "1", "0")
            If PanelVisible() Then
                If RunScan(gDoc) Then UserForm1.RefreshAll
            End If
    End Select
End Sub

' ---------- общие ----------
Private Function TargetDoc(ByRef doc As Document) As Boolean
    Dim v As String
    If Documents.Count = 0 Then
        MsgBox "Откройте документ, который нужно проверить.", vbExclamation, AppTitle()
        Exit Function
    End If
    Set doc = ActiveDocument
    v = ""
    On Error Resume Next
    v = doc.Variables("AbbrGOST_Report").Value
    On Error GoTo 0
    If v = "1" Then
        MsgBox "Это окно отчёта или справки. Перейдите в проверяемый документ и нажмите кнопку ещё раз.", vbExclamation, AppTitle()
        Exit Function
    End If
    TargetDoc = True
End Function

Private Sub ReportError()
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    MsgBox "Что-то пошло не так: " & Err.Description & " (код " & Err.Number & ")." & vbCr & vbCr & _
        "Документ не повреждён. Если ошибка повторяется, пришлите разработчику этот текст.", _
        vbExclamation, AppTitle()
End Sub

Private Function ScanActive(ByRef doc As Document) As Boolean
    If Not TargetDoc(doc) Then Exit Function
    ScanActive = RunScan(doc)
End Function

Public Function PanelVisible() As Boolean
    Dim f As Object
    For Each f In VBA.UserForms
        If TypeName(f) = "UserForm1" Then
            PanelVisible = f.Visible
            Exit Function
        End If
    Next
End Function

Public Function PanelDocAlive() As Boolean
    Dim s As String
    If gDoc Is Nothing Or Not gScanned Then Exit Function
    On Error Resume Next
    s = gDoc.Name
    PanelDocAlive = (Err.Number = 0)
    Err.Clear
End Function

' Открыть панель на нужной вкладке (после проверки документа)
Private Sub ShowPanel(ByVal page As Long, ByVal msg As String)
    If PanelVisible() Then
        UserForm1.RefreshAll
    Else
        UserForm1.Show vbModeless
        UserForm1.RefreshAll
    End If
    UserForm1.GoPage page
    If Len(msg) > 0 Then UserForm1.SetStatus msg
End Sub

' ---------- команды ----------
Public Sub CmdCheck()
    Dim doc As Document
    On Error GoTo eh
    If Not ScanActive(doc) Then Exit Sub
    ShowPanel 0, "Проверено по " & GostName() & ": " & VisibleCount() & " сокращений. Щёлкните по строке " & Dash() & " курсор перейдёт к первому упоминанию."
    Exit Sub
eh:
    ReportError
End Sub

Public Sub CmdPanelPage(ByVal page As Long)
    Dim doc As Document
    On Error GoTo eh
    If Not ScanActive(doc) Then Exit Sub
    ShowPanel page, ""
    Exit Sub
eh:
    ReportError
End Sub

Public Sub CmdFix()
    Dim doc As Document, msg As String
    On Error GoTo eh
    If Not ScanActive(doc) Then Exit Sub
    msg = FixAll()
    If Not RunScan(doc) Then Exit Sub
    ShowPanel 0, msg
    Exit Sub
eh:
    ReportError
End Sub

Public Sub CmdReport()
    Dim doc As Document
    On Error GoTo eh
    If Not ScanActive(doc) Then Exit Sub
    BuildReport
    Exit Sub
eh:
    ReportError
End Sub

Public Sub CmdComments()
    Dim doc As Document, n As Long
    On Error GoTo eh
    If Not ScanActive(doc) Then Exit Sub
    ClearOurComments doc
    n = AddComments()
    If PanelVisible() Then UserForm1.SetStatus "Поставлено примечаний: " & n & ". Убрать их: кнопка " & Q("Убрать примечания") & "."
    Exit Sub
eh:
    ReportError
End Sub

Public Sub CmdClear()
    Dim doc As Document
    On Error GoTo eh
    If Not TargetDoc(doc) Then Exit Sub
    ClearOurComments doc
    If PanelVisible() Then UserForm1.SetStatus "Примечания надстройки удалены."
    Exit Sub
eh:
    ReportError
End Sub

Public Function SelWordPublic() As String
    Dim s As String, punct As String
    punct = ".,;:()[]" & ChrW(&HAB) & ChrW(&HBB) & Chr(34)
    On Error Resume Next
    s = Selection.Text
    On Error GoTo 0
    s = NormSpaces(Replace(s, vbCr, ""))
    Do While Len(s) > 0
        If InStr(punct, Right$(s, 1)) > 0 Then s = Left$(s, Len(s) - 1) Else Exit Do
    Loop
    Do While Len(s) > 0
        If InStr(punct, Left$(s, 1)) > 0 Then s = Mid$(s, 2) Else Exit Do
    Loop
    If Len(s) > 40 Then s = ""
    SelWordPublic = s
End Function

Public Sub CmdAddDict()
    Dim doc As Document, k As String
    On Error GoTo eh
    k = SelWordPublic()
    If Not ScanActive(doc) Then Exit Sub
    If Len(k) > 0 And gIdx.Exists(k) Then
        ShowPanel 0, "Впишите расшифровку " & Q(k) & " и нажмите " & Q("В словарь") & "."
        UserForm1.SelectKey k
    Else
        ShowPanel 3, "Введите сокращение и расшифровку, затем " & Q("Сохранить") & "."
        UserForm1.PrefillDict k
    End If
    Exit Sub
eh:
    ReportError
End Sub

Public Sub CmdAddSymbol()
    Dim doc As Document, k As String
    On Error GoTo eh
    k = SelWordPublic()
    If Not ScanActive(doc) Then Exit Sub
    ShowPanel 2, "Впишите пояснение к обозначению и нажмите " & Q("Сохранить") & "."
    UserForm1.PrefillSym k
    Exit Sub
eh:
    ReportError
End Sub

Public Sub CmdOpenFile(ByVal path As String)
    On Error GoTo eh
    EnsureDefaults
    Shell "notepad.exe """ & path & """", vbNormalFocus
    Exit Sub
eh:
    ReportError
End Sub

' ---------- переход по тексту ----------
Private Function FindFrom(ByVal startPos As Long, ByVal key As String, ByVal wrap As Boolean) As Boolean
    Dim r As Range
    Set r = gDoc.Range(startPos, gDoc.Content.End)
    With r.Find
        .ClearFormatting
        .Text = key
        .MatchCase = True
        .MatchWholeWord = True
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
    End With
    If r.Find.Execute Then
        r.Select
        ActiveWindow.ScrollIntoView Selection.Range, True
        FindFrom = True
    ElseIf wrap Then
        FindFrom = FindFrom(0, key, False)
    End If
End Function

Public Function GoToFirstKey(ByVal key As String) As Boolean
    Dim r As Range, fp As Long, fo As Long, ix As Long
    If Not PanelDocAlive() Then Exit Function
    On Error Resume Next
    gDoc.Activate
    If gIdx.Exists(key) Then
        ix = gIdx(key)
        fp = gA(ix).FirstPara: fo = gA(ix).FirstOff
        If fp < 0 Then fp = gA(ix).FirstAnyPara: fo = gA(ix).FirstAnyOff
        If fp >= 0 Then
            Set r = OccRange(fp, fo, key)
            If Not r Is Nothing Then
                If r.Text = key Then
                    r.Select
                    ActiveWindow.ScrollIntoView Selection.Range, True
                    GoToFirstKey = True
                    Exit Function
                End If
            End If
        End If
    End If
    On Error GoTo 0
    GoToFirstKey = FindFrom(0, key, False)
End Function

Public Function GoToNextKey(ByVal key As String) As Boolean
    If Not PanelDocAlive() Then Exit Function
    gDoc.Activate
    GoToNextKey = FindFrom(Selection.End, key, True)
End Function

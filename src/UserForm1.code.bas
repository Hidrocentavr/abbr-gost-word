Option Explicit
' =====================================================================
'  Сокращения ГОСТ — панель (немодальное окно рядом с текстом)
'  Все элементы создаются при открытии в UserForm_Initialize.
' =====================================================================

Private WithEvents tbs As MSForms.TabStrip
Private WithEvents cmbGost As MSForms.ComboBox
Private WithEvents btnRecheck As MSForms.CommandButton
Private fr(0 To 3) As MSForms.Frame
Private lblStatus As MSForms.Label

' --- вкладка «Сокращения»
Private WithEvents cmbFilter As MSForms.ComboBox
Private lblCounts As MSForms.Label
Private WithEvents lstAbbr As MSForms.ListBox
Private lblKey As MSForms.Label
Private lblInfo As MSForms.Label
Private lblIssues As MSForms.Label
Private lblDecCap As MSForms.Label
Private txtDec As MSForms.TextBox
Private chkAutoDict As MSForms.CheckBox
Private WithEvents btnGo As MSForms.CommandButton
Private WithEvents btnNext As MSForms.CommandButton
Private WithEvents btnInsert As MSForms.CommandButton
Private WithEvents btnSpell As MSForms.CommandButton
Private WithEvents btnToDict As MSForms.CommandButton
Private WithEvents btnIgnore As MSForms.CommandButton
Private WithEvents btnFixAll As MSForms.CommandButton

' --- вкладка «Перечень»
Private lblPlace As MSForms.Label
Private lblStruct As MSForms.Label
Private WithEvents lstPrev As MSForms.ListBox
Private txtPrevDec As MSForms.TextBox
Private WithEvents btnPrevSave As MSForms.CommandButton
Private WithEvents chkDrop As MSForms.CheckBox
Private WithEvents cmbStyle As MSForms.ComboBox
Private WithEvents btnWrite As MSForms.CommandButton

' --- вкладка «Обозначения»
Private WithEvents lstSym As MSForms.ListBox
Private txtSymKey As MSForms.TextBox
Private txtSymDec As MSForms.TextBox
Private WithEvents btnSymSel As MSForms.CommandButton
Private WithEvents btnSymSave As MSForms.CommandButton
Private WithEvents btnSymDel As MSForms.CommandButton

' --- вкладка «Словарь»
Private WithEvents cmbDictKind As MSForms.ComboBox
Private WithEvents txtSearch As MSForms.TextBox
Private WithEvents lstDict As MSForms.ListBox
Private txtDKey As MSForms.TextBox
Private txtDDec As MSForms.TextBox
Private WithEvents btnDSave As MSForms.CommandButton
Private WithEvents btnDDel As MSForms.CommandButton
Private WithEvents btnDOpen As MSForms.CommandButton

Private mRows() As Long
Private mRowN As Long
Private mCurKey As String
Private mBusy As Boolean
Private mPrevKeys() As String
Private mPrevN As Long
Private mDictKeys() As String
Private mDictN As Long

' ---------------------------------------------------------------------
Private Function Ctl(ByVal parent As Object, ByVal kind As String, ByVal l As Single, ByVal t As Single, _
        ByVal w As Single, ByVal h As Single, Optional ByVal cap As String = "") As Object
    Dim c As Object
    Set c = parent.Controls.Add("Forms." & kind & ".1")
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    On Error Resume Next
    c.Font.Name = "Segoe UI"
    c.Font.Size = 9
    If Len(cap) > 0 Then c.Caption = cap
    On Error GoTo 0
    Set Ctl = c
End Function

Private Function Lbl(ByVal parent As Object, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, Optional ByVal cap As String = "") As MSForms.Label
    Dim c As MSForms.Label
    Set c = Ctl(parent, "Label", l, t, w, h)
    c.Caption = cap
    c.WordWrap = True
    Set Lbl = c
End Function

Private Sub UserForm_Initialize()
    Dim f As MSForms.Frame, i As Long
    Me.Caption = "Сокращения ГОСТ"
    Me.Width = 342
    Me.Height = 528
    Me.StartUpPosition = 0
    On Error Resume Next
    Me.Left = Application.Left + Application.Width - Me.Width - 24
    Me.Top = Application.Top + 120
    If Me.Left < 0 Then Me.Left = 20
    On Error GoTo 0

    Set cmbGost = Ctl(Me, "ComboBox", 6, 6, 212, 18)
    cmbGost.Style = fmStyleDropDownList
    cmbGost.AddItem "ГОСТ 7.32-2017 (отчёт о НИР)"
    cmbGost.AddItem "ГОСТ Р 7.0.11-2011 (диссертация)"
    cmbGost.AddItem "ГОСТ Р 2.105-2019 (ЕСКД)"
    Set btnRecheck = Ctl(Me, "CommandButton", 222, 5, 108, 20, "Проверить заново")

    Set tbs = Ctl(Me, "TabStrip", 6, 30, 324, 22)
    tbs.Tabs.Clear
    tbs.Tabs.Add "t0", "Сокращения"
    tbs.Tabs.Add "t1", "Перечень"
    tbs.Tabs.Add "t2", "Обозначения"
    tbs.Tabs.Add "t3", "Словарь"

    For i = 0 To 3
        Set f = Ctl(Me, "Frame", 6, 54, 324, 414)
        f.Caption = ""
        f.BorderStyle = fmBorderStyleNone
        f.SpecialEffect = fmSpecialEffectFlat
        f.Visible = (i = 0)
        Set fr(i) = f
    Next
    Set lblStatus = Lbl(Me, 6, 470, 324, 30)
    lblStatus.ForeColor = RGB(31, 56, 100)

    BuildPageAbbr fr(0)
    BuildPageList fr(1)
    BuildPageSym fr(2)
    BuildPageDict fr(3)

    mBusy = True
    cmbGost.ListIndex = GostIdx()
    cmbFilter.ListIndex = 0
    cmbStyle.ListIndex = IIf(OptTabStyle(), 0, 1)
    cmbDictKind.ListIndex = 0
    chkDrop.Value = True
    chkAutoDict.Value = True
    mBusy = False
End Sub

Private Sub BuildPageAbbr(ByVal f As MSForms.Frame)
    Set cmbFilter = Ctl(f, "ComboBox", 0, 2, 190, 18)
    cmbFilter.Style = fmStyleDropDownList
    cmbFilter.AddItem "Требуют действий"
    cmbFilter.AddItem "Только ошибки"
    cmbFilter.AddItem "Все сокращения"
    Set lblCounts = Lbl(f, 196, 4, 128, 16)
    Set lstAbbr = Ctl(f, "ListBox", 0, 24, 324, 160)
    lstAbbr.ColumnCount = 3
    lstAbbr.ColumnWidths = "54 pt;58 pt;200 pt"
    Set lblKey = Lbl(f, 0, 190, 324, 18)
    lblKey.Font.Size = 12
    lblKey.Font.Bold = True
    Set lblInfo = Lbl(f, 0, 208, 324, 26)
    lblInfo.ForeColor = RGB(89, 89, 89)
    Set lblIssues = Lbl(f, 0, 234, 324, 62)
    Set lblDecCap = Lbl(f, 0, 298, 324, 13, "Расшифровка (в именительном падеже):")
    Set txtDec = Ctl(f, "TextBox", 0, 312, 324, 19)
    Set chkAutoDict = Ctl(f, "CheckBox", 0, 333, 324, 15, "Запоминать мои расшифровки в словаре")
    Set btnGo = Ctl(f, "CommandButton", 0, 350, 76, 21, "К первому")
    Set btnNext = Ctl(f, "CommandButton", 80, 350, 76, 21, "Следующее")
    Set btnInsert = Ctl(f, "CommandButton", 160, 350, 164, 21, "Вставить расшифровку в текст")
    Set btnSpell = Ctl(f, "CommandButton", 0, 374, 120, 21, "Исправить написание")
    Set btnToDict = Ctl(f, "CommandButton", 124, 374, 96, 21, "В словарь")
    Set btnIgnore = Ctl(f, "CommandButton", 224, 374, 100, 21, "Не проверять")
    Set btnFixAll = Ctl(f, "CommandButton", 0, 396, 324, 18, "Исправить всё автоматически")
End Sub

Private Sub BuildPageList(ByVal f As MSForms.Frame)
    Set lblPlace = Lbl(f, 0, 2, 324, 30)
    Set lblStruct = Lbl(f, 0, 34, 324, 96)
    lblStruct.ForeColor = RGB(156, 87, 0)
    Lbl f, 0, 134, 324, 13, "Каким будет перечень (жёлтое в документе = проверить):"
    Set lstPrev = Ctl(f, "ListBox", 0, 148, 324, 156)
    lstPrev.ColumnCount = 3
    lstPrev.ColumnWidths = "14 pt;64 pt;230 pt"
    Set txtPrevDec = Ctl(f, "TextBox", 0, 308, 240, 19)
    Set btnPrevSave = Ctl(f, "CommandButton", 244, 307, 80, 21, "Изменить")
    Set chkDrop = Ctl(f, "CheckBox", 0, 332, 324, 15, "Убрать из перечня сокращения, которых нет в тексте")
    Lbl f, 0, 354, 70, 14, "Оформление:"
    Set cmbStyle = Ctl(f, "ComboBox", 72, 351, 150, 18)
    cmbStyle.Style = fmStyleDropDownList
    cmbStyle.AddItem "С выравниванием"
    cmbStyle.AddItem "В строку"
    Set btnWrite = Ctl(f, "CommandButton", 0, 378, 324, 26, "Записать перечень в документ")
    btnWrite.Font.Bold = True
End Sub

Private Sub BuildPageSym(ByVal f As MSForms.Frame)
    Lbl f, 0, 2, 324, 40, "Обозначения величин (F, " & ChrW(&H3C9) & ", Vmax...) надстройка в формулах не видит " & Dash() & " добавьте их здесь, и они войдут в перечень. Хранятся в самом документе."
    Set lstSym = Ctl(f, "ListBox", 0, 44, 324, 230)
    lstSym.ColumnCount = 2
    lstSym.ColumnWidths = "70 pt;240 pt"
    Lbl f, 0, 280, 70, 13, "Обозначение"
    Lbl f, 76, 280, 248, 13, "Пояснение с единицей, например: сила, Н"
    Set txtSymKey = Ctl(f, "TextBox", 0, 294, 70, 19)
    Set txtSymDec = Ctl(f, "TextBox", 76, 294, 248, 19)
    Set btnSymSel = Ctl(f, "CommandButton", 0, 318, 104, 21, "Из выделения")
    Set btnSymSave = Ctl(f, "CommandButton", 110, 318, 104, 21, "Сохранить")
    Set btnSymDel = Ctl(f, "CommandButton", 220, 318, 104, 21, "Удалить")
End Sub

Private Sub BuildPageDict(ByVal f As MSForms.Frame)
    Set cmbDictKind = Ctl(f, "ComboBox", 0, 2, 190, 18)
    cmbDictKind.Style = fmStyleDropDownList
    cmbDictKind.AddItem "Расшифровки (словарь)"
    cmbDictKind.AddItem "Не проверять (исключения)"
    Set txtSearch = Ctl(f, "TextBox", 196, 2, 128, 18)
    txtSearch.ControlTipText = "Поиск"
    Set lstDict = Ctl(f, "ListBox", 0, 24, 324, 250)
    lstDict.ColumnCount = 2
    lstDict.ColumnWidths = "70 pt;240 pt"
    Lbl f, 0, 280, 70, 13, "Сокращение"
    Lbl f, 76, 280, 248, 13, "Расшифровка"
    Set txtDKey = Ctl(f, "TextBox", 0, 294, 70, 19)
    Set txtDDec = Ctl(f, "TextBox", 76, 294, 248, 19)
    Set btnDSave = Ctl(f, "CommandButton", 0, 318, 104, 21, "Сохранить")
    Set btnDDel = Ctl(f, "CommandButton", 110, 318, 104, 21, "Удалить")
    Set btnDOpen = Ctl(f, "CommandButton", 220, 318, 104, 21, "Открыть файл")
End Sub

' ---------------------------------------------------------------------
'  Внешний интерфейс
' ---------------------------------------------------------------------
Public Sub GoPage(ByVal p As Long)
    If p < 0 Or p > 3 Then p = 0
    tbs.Value = p
    ShowPage p
End Sub

Public Sub SetStatus(ByVal s As String)
    lblStatus.Caption = s
End Sub

Public Sub RefreshAll()
    mBusy = True
    cmbGost.ListIndex = GostIdx()
    mBusy = False
    FillAbbr mCurKey
    FillListPage
    FillSym
    FillDict
End Sub

Public Sub PrefillDict(ByVal k As String)
    txtDKey.Text = k
    txtDDec.Text = ""
    If gDict Is Nothing Then Exit Sub
    If gDict.Exists(k) Then txtDDec.Text = gDict(k)
End Sub

Public Sub PrefillSym(ByVal k As String)
    txtSymKey.Text = k
    txtSymDec.Text = ""
End Sub

Public Sub SelectKey(ByVal k As String)
    Dim i As Long
    For i = 0 To mRowN - 1
        If gA(mRows(i)).Key = k Then
            mBusy = True
            lstAbbr.ListIndex = i
            mBusy = False
            ShowDetail mRows(i)
            Exit Sub
        End If
    Next
End Sub

' ---------------------------------------------------------------------
'  Вкладки
' ---------------------------------------------------------------------
Private Sub ShowPage(ByVal p As Long)
    Dim i As Long
    For i = 0 To 3
        fr(i).Visible = (i = p)
    Next
End Sub

Private Sub tbs_Change()
    ShowPage tbs.Value
End Sub

Private Sub cmbGost_Change()
    If mBusy Then Exit Sub
    SetOpt "Gost", CStr(cmbGost.ListIndex)
    RibbonRefresh
    DoRescan "Стандарт: " & GostName() & "."
End Sub

Private Sub btnRecheck_Click()
    DoRescan ""
End Sub

Private Function Rescan() As Boolean
    If Not PanelDocAlive() Then
        SetStatus "Документ закрыт. Откройте его и нажмите " & Q("Проверить") & " на вкладке."
        Exit Function
    End If
    SetStatus "Проверяю..."
    DoEvents
    On Error GoTo eh
    Rescan = RunScan(gDoc)
    Exit Function
eh:
    Application.ScreenUpdating = True
    SetStatus "Ошибка при проверке: " & Err.Description
End Function

Private Sub DoRescan(ByVal msg As String)
    If Not Rescan() Then Exit Sub
    RefreshAll
    If Len(msg) = 0 Then msg = "Проверено: " & VisibleCount() & " сокращений."
    SetStatus msg
End Sub

' ---------------------------------------------------------------------
'  Вкладка «Сокращения»
' ---------------------------------------------------------------------
Private Function FirstIssue(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, Chr(1))
    If p > 0 Then s = Left$(s, p - 1)
    If Len(s) = 0 Then s = "Всё в порядке"
    FirstIssue = s
End Function

Private Sub FillAbbr(ByVal keepKey As String)
    Dim ord() As Long, n As Long, i As Long, ix As Long, filt As Long, sel As Long
    Dim nE As Long, nW As Long
    filt = cmbFilter.ListIndex
    SortedIdx ord, n
    mBusy = True
    lstAbbr.Clear
    mRowN = 0
    If n > 0 Then ReDim mRows(0 To n - 1)
    sel = -1
    For i = 0 To n - 1
        ix = ord(i)
        If gA(ix).Severity = 3 Then nE = nE + 1
        If gA(ix).Severity = 2 Then nW = nW + 1
        If filt = 2 Or (filt = 0 And gA(ix).Severity >= 2) Or (filt = 1 And gA(ix).Severity = 3) Then
            lstAbbr.AddItem SevLabel(gA(ix).Severity)
            lstAbbr.List(mRowN, 1) = gA(ix).Key
            lstAbbr.List(mRowN, 2) = FirstIssue(gA(ix).Issues)
            mRows(mRowN) = ix
            If gA(ix).Key = keepKey Then sel = mRowN
            mRowN = mRowN + 1
        End If
    Next
    lblCounts.Caption = "ошибок " & nE & ", внимание " & nW
    If mRowN = 0 Then
        ClearDetail
        If n = 0 Then
            lblKey.Caption = "Сокращений не найдено"
        Else
            lblKey.Caption = ChrW(&H2714) & " Здесь всё исправлено"
            lblInfo.Caption = "Выберите " & Q("Все сокращения") & ", чтобы увидеть остальные."
        End If
    Else
        If sel < 0 Then sel = 0
        lstAbbr.ListIndex = sel
        ShowDetail mRows(sel)
    End If
    mBusy = False
End Sub

Private Sub ClearDetail()
    mCurKey = ""
    lblKey.Caption = ""
    lblInfo.Caption = ""
    lblIssues.Caption = ""
    txtDec.Text = ""
    btnGo.Enabled = False: btnNext.Enabled = False: btnInsert.Enabled = False
    btnSpell.Enabled = False: btnToDict.Enabled = False: btnIgnore.Enabled = False
End Sub

Private Function CurIx() As Long
    CurIx = -1
    If Len(mCurKey) = 0 Or gIdx Is Nothing Then Exit Function
    If gIdx.Exists(mCurKey) Then CurIx = gIdx(mCurKey)
End Function

Private Sub ShowDetail(ByVal ix As Long)
    Dim review As Boolean, src As String, dec As String, info As String, pg As Long, s As String
    mCurKey = gA(ix).Key
    lblKey.Caption = gA(ix).Key & "   " & Dash() & "   " & SevLabel(gA(ix).Severity)
    Select Case gA(ix).Severity
        Case 3: lblKey.ForeColor = RGB(192, 0, 0)
        Case 2: lblKey.ForeColor = RGB(191, 113, 0)
        Case 1: lblKey.ForeColor = RGB(31, 78, 121)
        Case Else: lblKey.ForeColor = RGB(56, 118, 29)
    End Select
    dec = BestDec(ix, review, src)
    If gA(ix).FirstPara >= 0 Then
        pg = PageOf(gA(ix).FirstPara, gA(ix).FirstOff, gA(ix).Key)
    ElseIf gA(ix).FirstAnyPara >= 0 Then
        pg = PageOf(gA(ix).FirstAnyPara, gA(ix).FirstAnyOff, gA(ix).Key)
    End If
    info = "Употреблено: " & gA(ix).Count
    If pg > 0 Then info = info & "   " & ChrW(&H2022) & "   первое упоминание: стр. " & pg
    info = info & "   " & ChrW(&H2022) & "   в перечне: " & IIf(gA(ix).InList, "да", "нет")
    If Len(dec) > 0 Then info = info & vbLf & "Расшифровка взята: " & src
    lblInfo.Caption = info
    s = gA(ix).Issues
    If Len(s) = 0 Then
        lblIssues.Caption = ChrW(&H2714) & " Замечаний нет."
    Else
        lblIssues.Caption = ChrW(&H2022) & " " & Replace(s, Chr(1), vbLf & ChrW(&H2022) & " ")
    End If
    txtDec.Text = dec
    btnGo.Enabled = (gA(ix).Count > 0)
    btnNext.Enabled = (gA(ix).Count > 0)
    btnInsert.Enabled = (gA(ix).Count > 0 And Not gA(ix).DecAtFirst And Not gA(ix).Mixed)
    btnSpell.Enabled = (gA(ix).VariantCount > 0 Or (gA(ix).Mixed And Len(gA(ix).MixTarget) > 0))
    btnToDict.Enabled = Not gA(ix).Mixed
    btnIgnore.Enabled = Not gA(ix).Mixed
End Sub

Private Sub cmbFilter_Change()
    If mBusy Then Exit Sub
    If gIdx Is Nothing Then Exit Sub
    FillAbbr mCurKey
End Sub

Private Sub lstAbbr_Click()
    If mBusy Then Exit Sub
    If lstAbbr.ListIndex < 0 Or lstAbbr.ListIndex >= mRowN Then Exit Sub
    ShowDetail mRows(lstAbbr.ListIndex)
    If gA(mRows(lstAbbr.ListIndex)).Count > 0 Then
        If GoToFirstKey(mCurKey) Then SetStatus "Курсор у первого упоминания " & Q(mCurKey) & "."
    End If
End Sub

Private Sub btnGo_Click()
    If Len(mCurKey) = 0 Then Exit Sub
    If GoToFirstKey(mCurKey) Then
        SetStatus "Курсор у первого упоминания " & Q(mCurKey) & "."
    Else
        SetStatus Q(mCurKey) & " в тексте не найдено."
    End If
End Sub

Private Sub btnNext_Click()
    If Len(mCurKey) = 0 Then Exit Sub
    If GoToNextKey(mCurKey) Then
        SetStatus "Следующее вхождение " & Q(mCurKey) & "."
    Else
        SetStatus Q(mCurKey) & " в тексте не найдено."
    End If
End Sub

Private Sub RememberDec(ByVal key As String, ByVal dec As String)
    Dim review As Boolean, src As String, ix As Long
    dec = StripEndPunct(dec)
    If Len(dec) = 0 Then Exit Sub
    If gIdx.Exists(key) Then
        ix = gIdx(key)
        If dec <> BestDec(ix, review, src) Then
            If gOverride.Exists(key) Then gOverride(key) = dec Else gOverride.Add key, dec
        End If
    End If
    If chkAutoDict.Value Then
        If Not gDict.Exists(key) Then
            UpsertKV DictPath(), key, dec
        ElseIf gDict(key) <> dec Then
            UpsertKV DictPath(), key, dec
        End If
    End If
End Sub

Private Sub btnInsert_Click()
    Dim k As String, d As String, msg As String
    k = mCurKey: d = Trim$(txtDec.Text)
    If Len(k) = 0 Then Exit Sub
    If Len(d) = 0 Then SetStatus "Впишите расшифровку в поле над кнопками.": Exit Sub
    On Error GoTo eh
    RememberDec k, d
    If Not Rescan() Then Exit Sub
    msg = InsertDecoding(k, d)
    DoRescan msg
    Exit Sub
eh:
    Application.ScreenUpdating = True
    SetStatus "Не получилось: " & Err.Description
End Sub

Private Sub btnSpell_Click()
    Dim k As String, msg As String
    k = mCurKey
    If Len(k) = 0 Then Exit Sub
    On Error GoTo eh
    If Not Rescan() Then Exit Sub
    msg = FixSpelling(k)
    If gIdx.Exists(k) Then
        If gA(gIdx(k)).Mixed Then mCurKey = gA(gIdx(k)).MixTarget
    End If
    DoRescan msg
    Exit Sub
eh:
    Application.ScreenUpdating = True
    SetStatus "Не получилось: " & Err.Description
End Sub

Private Sub btnToDict_Click()
    Dim k As String, d As String
    k = mCurKey: d = StripEndPunct(Trim$(txtDec.Text))
    If Len(k) = 0 Then Exit Sub
    If Len(d) = 0 Then SetStatus "Впишите расшифровку, затем нажмите " & Q("В словарь") & ".": Exit Sub
    UpsertKV DictPath(), k, d
    DoRescan Q(k) & " " & Dash() & " " & d & ": сохранено в словаре, будет подставляться во всех документах."
End Sub

Private Sub btnIgnore_Click()
    Dim k As String
    k = mCurKey
    If Len(k) = 0 Then Exit Sub
    UpsertKV ExcPath(), k, ""
    mCurKey = ""
    DoRescan Q(k) & " больше не проверяется. Вернуть: вкладка " & Q("Словарь") & " " & ChrW(&H2192) & " " & Q("Не проверять") & "."
End Sub

Private Sub btnFixAll_Click()
    Dim msg As String
    On Error GoTo eh
    If Not Rescan() Then Exit Sub
    msg = FixAll()
    DoRescan msg
    Exit Sub
eh:
    Application.ScreenUpdating = True
    SetStatus "Не получилось: " & Err.Description
End Sub

' ---------------------------------------------------------------------
'  Вкладка «Перечень»
' ---------------------------------------------------------------------
Private Sub FillListPage()
    Dim keys() As String, decs() As String, rev() As Boolean, n As Long, i As Long, u As Long
    lblPlace.Caption = ListPlaceText()
    lblStruct.Caption = StructText()
    n = BuildListEntries(keys, decs, rev, chkDrop.Value)
    lstPrev.Clear
    mPrevN = n
    If n > 0 Then ReDim mPrevKeys(0 To n - 1)
    For i = 0 To n - 1
        lstPrev.AddItem IIf(rev(i), "?", "")
        lstPrev.List(i, 1) = keys(i)
        lstPrev.List(i, 2) = decs(i)
        mPrevKeys(i) = keys(i)
    Next
    u = UnusedInList()
    If u > 0 Then
        chkDrop.Caption = "Убрать из перечня сокращения, которых нет в тексте (" & u & ")"
    Else
        chkDrop.Caption = "Убрать из перечня сокращения, которых нет в тексте"
    End If
    txtPrevDec.Text = ""
End Sub

Private Sub lstPrev_Click()
    If lstPrev.ListIndex < 0 Then Exit Sub
    txtPrevDec.Text = lstPrev.List(lstPrev.ListIndex, 2)
    If IsAbbrToken(mPrevKeys(lstPrev.ListIndex)) Then GoToFirstKey mPrevKeys(lstPrev.ListIndex)
End Sub

Private Sub btnPrevSave_Click()
    Dim k As String, d As String, syms As Object
    If lstPrev.ListIndex < 0 Then SetStatus "Выберите строку перечня.": Exit Sub
    k = mPrevKeys(lstPrev.ListIndex)
    d = StripEndPunct(Trim$(txtPrevDec.Text))
    If Len(d) = 0 Then Exit Sub
    If gIdx.Exists(k) Then
        If gOverride.Exists(k) Then gOverride(k) = d Else gOverride.Add k, d
        If chkAutoDict.Value Then UpsertKV DictPath(), k, d
    Else
        Set syms = GetSymbols(gDoc)
        If syms.Exists(k) Then syms(k) = d Else syms.Add k, d
        SaveSymbols gDoc, syms
    End If
    FillListPage
    SetStatus "Расшифровка " & Q(k) & " изменена. В документ попадёт при записи перечня."
End Sub

Private Sub chkDrop_Click()
    If mBusy Then Exit Sub
    If gIdx Is Nothing Then Exit Sub
    FillListPage
End Sub

Private Sub cmbStyle_Change()
    If mBusy Then Exit Sub
    SetOpt "ListStyle", CStr(cmbStyle.ListIndex)
    RibbonRefresh
End Sub

Private Sub btnWrite_Click()
    Dim msg As String
    On Error GoTo eh
    If Not Rescan() Then Exit Sub
    msg = WriteList(chkDrop.Value)
    DoRescan msg
    GoPage 1
    Exit Sub
eh:
    Application.ScreenUpdating = True
    SetStatus "Не получилось: " & Err.Description
End Sub

' ---------------------------------------------------------------------
'  Вкладка «Обозначения»
' ---------------------------------------------------------------------
Private Sub FillSym()
    Dim d As Object, k As Variant, i As Long
    lstSym.Clear
    If Not PanelDocAlive() Then Exit Sub
    Set d = GetSymbols(gDoc)
    For Each k In d.Keys
        lstSym.AddItem CStr(k)
        lstSym.List(i, 1) = d(k)
        i = i + 1
    Next
End Sub

Private Sub lstSym_Click()
    If lstSym.ListIndex < 0 Then Exit Sub
    txtSymKey.Text = lstSym.List(lstSym.ListIndex, 0)
    txtSymDec.Text = lstSym.List(lstSym.ListIndex, 1)
End Sub

Private Sub btnSymSel_Click()
    txtSymKey.Text = SelWordPublic()
    txtSymDec.SetFocus
End Sub

Private Sub btnSymSave_Click()
    Dim d As Object, k As String, v As String
    k = Trim$(txtSymKey.Text): v = StripEndPunct(Trim$(txtSymDec.Text))
    If Len(k) = 0 Or Len(v) = 0 Then SetStatus "Заполните обозначение и пояснение.": Exit Sub
    If Not PanelDocAlive() Then Exit Sub
    Set d = GetSymbols(gDoc)
    If d.Exists(k) Then d(k) = v Else d.Add k, v
    SaveSymbols gDoc, d
    FillSym
    FillListPage
    txtSymKey.Text = "": txtSymDec.Text = ""
    SetStatus Q(k) & " " & Dash() & " " & v & ": сохранено. Обозначений: " & d.Count & "."
End Sub

Private Sub btnSymDel_Click()
    Dim d As Object, k As String
    k = Trim$(txtSymKey.Text)
    If Len(k) = 0 Or Not PanelDocAlive() Then Exit Sub
    Set d = GetSymbols(gDoc)
    If d.Exists(k) Then
        d.Remove k
        SaveSymbols gDoc, d
        FillSym
        FillListPage
        SetStatus Q(k) & " удалено."
    End If
End Sub

' ---------------------------------------------------------------------
'  Вкладка «Словарь»
' ---------------------------------------------------------------------
Private Function DictFile() As String
    If cmbDictKind.ListIndex = 1 Then DictFile = ExcPath() Else DictFile = DictPath()
End Function

Private Sub FillDict()
    Dim d As Object, k As Variant, keys() As String, n As Long, i As Long, j As Long, t As String, q2 As String
    EnsureDefaults
    Set d = LoadKV(DictFile())
    q2 = LCase$(Trim$(txtSearch.Text))
    ReDim keys(0 To d.Count)
    For Each k In d.Keys
        If Len(q2) = 0 Or InStr(LCase$(CStr(k)), q2) > 0 Or InStr(LCase$(CStr(d(k))), q2) > 0 Then
            keys(n) = CStr(k): n = n + 1
        End If
    Next
    For i = 1 To n - 1
        t = keys(i): j = i - 1
        Do While j >= 0
            If StrComp(SortKey(keys(j)), SortKey(t), vbBinaryCompare) <= 0 Then Exit Do
            keys(j + 1) = keys(j): j = j - 1
        Loop
        keys(j + 1) = t
    Next
    lstDict.Clear
    mDictN = n
    If n > 0 Then ReDim mDictKeys(0 To n - 1)
    For i = 0 To n - 1
        lstDict.AddItem keys(i)
        lstDict.List(i, 1) = d(keys(i))
        mDictKeys(i) = keys(i)
    Next
End Sub

Private Sub cmbDictKind_Change()
    If mBusy Then Exit Sub
    txtDKey.Text = "": txtDDec.Text = ""
    txtDDec.Enabled = (cmbDictKind.ListIndex = 0)
    FillDict
End Sub

Private Sub txtSearch_Change()
    FillDict
End Sub

Private Sub lstDict_Click()
    If lstDict.ListIndex < 0 Then Exit Sub
    txtDKey.Text = lstDict.List(lstDict.ListIndex, 0)
    txtDDec.Text = lstDict.List(lstDict.ListIndex, 1)
End Sub

Private Sub btnDSave_Click()
    Dim k As String, v As String
    k = Trim$(txtDKey.Text): v = StripEndPunct(Trim$(txtDDec.Text))
    If Len(k) = 0 Then SetStatus "Введите сокращение.": Exit Sub
    If cmbDictKind.ListIndex = 0 And Len(v) = 0 Then SetStatus "Введите расшифровку.": Exit Sub
    If cmbDictKind.ListIndex = 1 Then v = ""
    UpsertKV DictFile(), k, v
    FillDict
    If PanelDocAlive() Then
        DoRescan Q(k) & " сохранено."
    Else
        SetStatus Q(k) & " сохранено."
    End If
End Sub

Private Sub btnDDel_Click()
    Dim k As String
    k = Trim$(txtDKey.Text)
    If Len(k) = 0 Then Exit Sub
    DeleteKV DictFile(), k
    txtDKey.Text = "": txtDDec.Text = ""
    FillDict
    If PanelDocAlive() Then
        DoRescan Q(k) & " удалено."
    Else
        SetStatus Q(k) & " удалено."
    End If
End Sub

Private Sub btnDOpen_Click()
    CmdOpenFile DictFile()
End Sub

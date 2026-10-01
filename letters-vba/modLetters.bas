Attribute VB_Name = "modLetters"
Option Explicit

' ============================================================
' Letters register - logic, form builder and sheet button.
' Run SetupLetters ONCE. Arabic text is stored as Unicode code
' points (function U) so the file is pure ASCII and imports
' correctly regardless of the Windows code page.
' ============================================================

Private Const TBL_NAME As String = "Table2"
Private Const FORM_NAME As String = "frmLetters"
Private Const BTN_NAME As String = "btnLettersForm"
Private Const MB_RTL As Long = 1572864          ' right-to-left reading + right aligned

' brand colours (BGR long values), taken from the logo's dark teal
Private Const CLR_TEAL As Long = &H3E3500          ' RGB(0,53,62)
Private Const CLR_TEAL_MID As Long = &H756B0F      ' RGB(15,107,117)
Private Const CLR_BG As Long = &HF5F4EE            ' RGB(238,244,245)
Private Const CLR_PALE As Long = &HE6E4D6          ' RGB(214,228,230)
Private Const CLR_LINE As Long = &HD5D2BE          ' RGB(190,210,213)
Private Const CLR_DANGER As Long = &H403DA6        ' RGB(166,61,64)
Private Const CLR_WHITE As Long = &HFFFFFF

' column positions inside Table2
Private Const C_SEQ As Long = 1
Private Const C_NO As Long = 2
Private Const C_COMP As Long = 3
Private Const C_TITLE As Long = 4
Private Const C_TYPE As Long = 5
Private Const C_DATE As Long = 6
Private Const C_NOTES As Long = 7

' state of the row currently loaded for editing
Private mRow As Long
Private mOrigNo As String
Private mOrigDate As String
Private mOrigComp As String
Private mOrigTitle As String
Private mOrigType As String
Private mStep As String                 ' last setup step, shown if setup fails

' ---------------------------------------------------------------
' helpers
' ---------------------------------------------------------------
Private Function U(ByVal codes As String) As String
    Dim p() As String, i As Long
    p = Split(codes, " ")
    For i = 0 To UBound(p)
        U = U & ChrW(CLng("&H" & p(i)))
    Next i
End Function

Private Function GetTable() As ListObject
    On Error Resume Next
    Set GetTable = ThisWorkbook.Worksheets(U("0627 0644 0639 0642 0627 0631 064A 0629")).ListObjects(TBL_NAME)
    On Error GoTo 0
    If GetTable Is Nothing Then
        MsgBox U("062A 0639 0630 0631 0020 0627 0644 0639 062B 0648 0631 0020 0639 0644 0649 0020 0627 0644 062C 062F 0648 0644 0020") & TBL_NAME & U("0020 0641 064A 0020 0648 0631 0642 0629 0020 0627 0644 0639 0642 0627 0631 064A 0629"), vbCritical + MB_RTL
    End If
End Function

Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        Txt = Format$(v, "dd/mm/yyyy")
    Else
        Txt = Trim$(CStr(v))
    End If
End Function

' lower-case, strip diacritics, unify alef/teh marbuta/alef maksura, Arabic digits -> 0-9
Private Function Norm(ByVal s As String) As String
    Dim i As Long, ch As Long, out As String
    s = LCase$(s)
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If ch < 0 Then ch = ch + 65536
        Select Case ch
            Case &H64B To &H652, &H640
                ' diacritics / tatweel: dropped
            Case &H622, &H623, &H625
                out = out & ChrW(&H627)
            Case &H629
                out = out & ChrW(&H647)
            Case &H649
                out = out & ChrW(&H64A)
            Case &H660 To &H669
                out = out & ChrW(ch - &H660 + 48)
            Case &H6F0 To &H6F9
                out = out & ChrW(ch - &H6F0 + 48)
            Case Else
                out = out & Mid$(s, i, 1)
        End Select
    Next i
    Norm = out
End Function

' accepts d/m/yyyy, d-m-yyyy, d.m.yyyy and yyyy/m/d (Arabic digits allowed)
Private Function ParseDate(ByVal s As String, ByRef d As Date) As Boolean
    Dim p() As String, y As Long, m As Long, dd As Long
    s = Norm(Trim$(s))
    s = Replace(Replace(Replace(s, "-", "/"), ".", "/"), "\", "/")
    p = Split(s, "/")
    If UBound(p) <> 2 Then Exit Function
    If Not IsNumeric(p(0)) Then Exit Function
    If Not IsNumeric(p(1)) Then Exit Function
    If Not IsNumeric(p(2)) Then Exit Function
    If Len(Trim$(p(0))) = 4 Then
        y = CLng(p(0)): m = CLng(p(1)): dd = CLng(p(2))
    Else
        dd = CLng(p(0)): m = CLng(p(1)): y = CLng(p(2))
    End If
    If y < 100 Then y = y + 2000
    If m < 1 Or m > 12 Or dd < 1 Or dd > 31 Or y < 1900 Or y > 2200 Then Exit Function
    d = DateSerial(y, m, dd)
    ParseDate = (Day(d) = dd And Month(d) = m)
End Function

' index (inside the table body) of the first row whose company cell is empty; 0 = none
Private Function FirstEmptyRow(lo As ListObject) As Long
    Dim v As Variant, r As Long
    If lo.DataBodyRange Is Nothing Then Exit Function
    v = lo.DataBodyRange.Value
    For r = 1 To UBound(v, 1)
        If Len(Txt(v(r, C_COMP))) = 0 Then FirstEmptyRow = r: Exit Function
    Next r
End Function

' R1C1 formula of the first row (other than skipRow) that still holds a formula in this column
Private Function RefFormula(lo As ListObject, ByVal col As Long, ByVal skipRow As Long) As String
    Dim r As Long, c As Range
    If lo.DataBodyRange Is Nothing Then Exit Function
    For r = 1 To lo.ListRows.Count
        If r <> skipRow Then
            Set c = lo.DataBodyRange.Cells(r, col)
            If c.HasFormula Then RefFormula = c.FormulaR1C1: Exit Function
        End If
    Next r
End Function

' put the column formula back into a cell that was overwritten by a manual value
Private Function RestoreFormula(lo As ListObject, ByVal r As Long, ByVal col As Long) As Boolean
    Dim c As Range, fml As String, oldAuto As Boolean
    Set c = lo.DataBodyRange.Cells(r, col)
    If c.HasFormula Then RestoreFormula = True: Exit Function
    fml = RefFormula(lo, col, r)
    If Len(fml) = 0 Then Exit Function
    oldAuto = Application.AutoCorrect.AutoFillFormulasInLists
    Application.AutoCorrect.AutoFillFormulasInLists = False   ' never spread to other rows
    c.FormulaR1C1 = fml
    Application.AutoCorrect.AutoFillFormulasInLists = oldAuto
    RestoreFormula = True
End Function

Private Function Validate(f As Object) As Boolean
    If Len(Trim$(f.cboCompany.Value)) = 0 Then
        MsgBox U("0627 0644 0634 0631 0643 0629 0020 002F 0020 0627 0644 062C 0647 0629 0020 0645 0637 0644 0648 0628 0629 002E"), vbExclamation + MB_RTL
        f.cboCompany.SetFocus
    ElseIf Len(Trim$(f.txtTitle.Value)) = 0 Then
        MsgBox U("0639 0646 0648 0627 0646 0020 0627 0644 062E 0637 0627 0628 0020 0645 0637 0644 0648 0628 002E"), vbExclamation + MB_RTL
        f.txtTitle.SetFocus
    ElseIf Len(Trim$(f.cboType.Value)) = 0 Then
        MsgBox U("062D 062F 062F 0020 0646 0648 0639 0020 0627 0644 062E 0637 0627 0628 0020 0028 0635 0627 062F 0631 0020 002F 0020 0648 0627 0631 062F 0029 002E"), vbExclamation + MB_RTL
        f.cboType.SetFocus
    Else
        Validate = True
    End If
End Function

Private Function IsIncoming(f As Object) As Boolean
    IsIncoming = (Norm(Trim$(f.cboType.Value)) = Norm(U("0648 0627 0631 062F")))
End Function

Private Sub SetType(f As Object, ByVal t As String)
    Dim i As Long
    f.cboType.ListIndex = -1
    For i = 0 To f.cboType.ListCount - 1
        If Norm(f.cboType.List(i)) = Norm(t) Then f.cboType.ListIndex = i: Exit For
    Next i
End Sub

Private Sub SetCount(f As Object, ByVal n As Long)
    f.lblCount.Caption = U("0639 062F 062F 0020 0627 0644 0646 062A 0627 0626 062C 003A 0020") & n
End Sub

' the loaded row must still be the one the user picked
Private Function RowStillValid(lo As ListObject) As Boolean
    If mRow < 1 Or mRow > lo.ListRows.Count Then
        MsgBox U("0627 062E 062A 0631 0020 062E 0637 0627 0628 0627 0020 0645 0646 0020 0627 0644 0642 0627 0626 0645 0629 0020 0623 0648 0644 0627 002E"), vbExclamation + MB_RTL
        Exit Function
    End If
    With lo.DataBodyRange
        If Txt(.Cells(mRow, C_COMP).Value) <> mOrigComp Or Txt(.Cells(mRow, C_TITLE).Value) <> mOrigTitle Then
            MsgBox U("062A 063A 064A 0631 0020 0645 062D 062A 0648 0649 0020 0647 0630 0627 0020 0627 0644 0635 0641 0020 0645 0646 0630 0020 0627 062E 062A 064A 0627 0631 0647 002E 0020 0623 0639 062F 0020 0627 0644 0628 062D 062B 0020 062B 0645 0020 0627 062E 062A 0631 0647 0627 0020 0645 0646 0020 062C 062F 064A 062F 002E"), vbExclamation + MB_RTL
            Exit Function
        End If
    End With
    RowStillValid = True
End Function

' ---------------------------------------------------------------
' launcher (sheet button)
' ---------------------------------------------------------------
Private Function HasCtl(f As Object, ByVal nm As String) As Boolean
    Dim c As Object
    On Error Resume Next
    Set c = f.Controls(nm)
    On Error GoTo 0
    HasCtl = Not c Is Nothing
End Function

' the form may carry a different name if VBA refused to rename it, so look it up by its content
Public Sub ShowLettersForm()
    Dim f As Object, k As Long, cand As String
    For k = 0 To 20
        Select Case k
            Case 0: cand = FORM_NAME
            Case 1 To 9: cand = FORM_NAME & (k + 1)
            Case Else: cand = "UserForm" & (k - 9)
        End Select
        Set f = Nothing
        On Error Resume Next
        Set f = VBA.UserForms.Add(cand)
        On Error GoTo 0
        If Not f Is Nothing Then
            If HasCtl(f, "lstResults") Then
                f.Show
                Exit Sub
            End If
            Unload f
        End If
    Next k
    MsgBox U("0627 0644 0646 0627 0641 0630 0629 0020 063A 064A 0631 0020 0645 0648 062C 0648 062F 0629 002E 0020 0634 063A 0644 0020 0053 0065 0074 0075 0070 004C 0065 0074 0074 0065 0072 0073 0020 0623 0648 0644 0627 002E"), vbExclamation + MB_RTL
End Sub

' ---------------------------------------------------------------
' form logic (called from the thin event stubs inside the form)
' ---------------------------------------------------------------
Public Sub FormInit(f As Object)
    f.cboType.List = Array(U("0635 0627 062F 0631"), U("0648 0627 0631 062F"))
    f.cboFilter.List = Array(U("0627 0644 0643 0644"), U("0635 0627 062F 0631"), U("0648 0627 0631 062F"))
    f.cboFilter.Value = U("0627 0644 0643 0644")
    LoadCompanies f
    FormClear f
    FormSearch f
End Sub

Public Sub FormTypeChanged(f As Object)
    Dim show As Boolean
    show = IsIncoming(f)
    f.lblNo.Visible = show
    f.txtNo.Visible = show
End Sub

Public Sub FormClear(f As Object)
    mRow = 0: mOrigNo = "": mOrigDate = "": mOrigComp = "": mOrigTitle = "": mOrigType = ""
    f.cboCompany.Value = ""
    f.txtTitle.Value = ""
    f.cboType.ListIndex = -1
    f.txtNo.Value = ""
    f.lstResults.ListIndex = -1
    FormTypeChanged f
End Sub

Private Sub LoadCompanies(f As Object)
    Dim lo As ListObject, v As Variant, r As Long, s As String
    Dim col As New Collection, arr() As String, n As Long, i As Long, j As Long, t As String
    Dim keep As String
    Set lo = GetTable()
    If lo Is Nothing Then Exit Sub
    keep = f.cboCompany.Value
    If Not lo.DataBodyRange Is Nothing Then
        v = lo.DataBodyRange.Value
        For r = 1 To UBound(v, 1)
            s = Txt(v(r, C_COMP))
            If Len(s) > 0 Then
                On Error Resume Next
                col.Add s, Norm(s)          ' duplicate key -> skipped
                On Error GoTo 0
            End If
        Next r
    End If
    n = col.Count
    If n = 0 Then
        f.cboCompany.Clear
    Else
        ReDim arr(0 To n - 1)
        For i = 1 To n
            t = col(i)
            j = i - 2
            Do While j >= 0
                If StrComp(arr(j), t, vbTextCompare) <= 0 Then Exit Do
                arr(j + 1) = arr(j)
                j = j - 1
            Loop
            arr(j + 1) = t
        Next i
        f.cboCompany.List = arr
    End If
    f.cboCompany.Value = keep
End Sub

Public Sub FormSearch(f As Object)
    Dim lo As ListObject, v As Variant, r As Long, k As Long, j As Long
    Dim q As String, flt As String, typ As String, comp As String, hay As String
    Dim sSeq As String, sNo As String, sTitle As String, sNotes As String, sDate As String
    Dim res() As Variant, out() As Variant
    Set lo = GetTable()
    If lo Is Nothing Then Exit Sub
    q = Trim$(f.txtSearch.Value)
    If Len(q) = 0 And mRow = 0 Then q = Trim$(f.cboCompany.Value)   ' search by the company box too
    q = Norm(q)
    flt = Norm(Trim$(f.cboFilter.Value))
    If lo.DataBodyRange Is Nothing Then
        f.lstResults.Clear: SetCount f, 0: Exit Sub
    End If
    v = lo.DataBodyRange.Value
    ReDim res(0 To UBound(v, 1) - 1, 0 To 6)
    For r = UBound(v, 1) To 1 Step -1              ' newest first
        comp = Txt(v(r, C_COMP))
        If Len(comp) > 0 Then
            typ = Txt(v(r, C_TYPE))
            If Len(flt) = 0 Or flt = Norm(U("0627 0644 0643 0644")) Or Norm(typ) = flt Then
                sSeq = Txt(v(r, C_SEQ))
                sNo = Txt(v(r, C_NO))
                sTitle = Txt(v(r, C_TITLE))
                sNotes = Txt(v(r, C_NOTES))
                sDate = Txt(v(r, C_DATE))
                hay = Norm(sSeq & "|" & sNo & "|" & comp & "|" & sTitle & "|" & sNotes & "|" & sDate)
                If Len(q) = 0 Or InStr(1, hay, q, vbBinaryCompare) > 0 Then
                    ' list is right-to-left: first column is the rightmost
                    res(k, 0) = sSeq
                    res(k, 1) = sNo
                    res(k, 2) = comp
                    res(k, 3) = sTitle
                    res(k, 4) = typ
                    res(k, 5) = sDate
                    res(k, 6) = r
                    k = k + 1
                End If
            End If
        End If
    Next r
    If k = 0 Then
        f.lstResults.Clear
    Else
        ReDim out(0 To k - 1, 0 To 6)
        For r = 0 To k - 1
            For j = 0 To 6
                out(r, j) = res(r, j)
            Next j
        Next r
        f.lstResults.List = out
    End If
    SetCount f, k
End Sub

Public Sub FormPick(f As Object)
    Dim lo As ListObject, i As Long, r As Long, rw As Range
    i = f.lstResults.ListIndex
    If i < 0 Then Exit Sub
    Set lo = GetTable()
    If lo Is Nothing Then Exit Sub
    r = CLng(f.lstResults.List(i, 6))
    If r < 1 Or r > lo.ListRows.Count Then Exit Sub
    Set rw = lo.ListRows(r).Range
    SetType f, Txt(rw.Cells(1, C_TYPE).Value)
    f.cboCompany.Value = Txt(rw.Cells(1, C_COMP).Value)
    f.txtTitle.Value = Txt(rw.Cells(1, C_TITLE).Value)
    mRow = r
    mOrigComp = Txt(rw.Cells(1, C_COMP).Value)
    mOrigTitle = Txt(rw.Cells(1, C_TITLE).Value)
    mOrigType = Txt(rw.Cells(1, C_TYPE).Value)
    mOrigNo = Txt(rw.Cells(1, C_NO).Value)
    If mOrigNo = U("0628 062F 0648 0646 0020 0631 0642 0645") Then mOrigNo = ""       ' the formula's placeholder for incoming letters
    f.txtNo.Value = mOrigNo
End Sub

Public Sub FormAdd(f As Object)
    Dim lo As ListObject, r As Long, rw As Range
    Dim sNo As String, isNew As Boolean, oldAuto As Boolean, oldScr As Boolean
    If Not Validate(f) Then Exit Sub
    Set lo = GetTable()
    If lo Is Nothing Then Exit Sub
    oldScr = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False
    r = FirstEmptyRow(lo)
    If r = 0 Then
        lo.ListRows.Add
        r = lo.ListRows.Count
        isNew = True
    End If
    Set rw = lo.ListRows(r).Range
    rw.Cells(1, C_COMP).Value = Trim$(f.cboCompany.Value)
    rw.Cells(1, C_TITLE).Value = Trim$(f.txtTitle.Value)
    rw.Cells(1, C_TYPE).Value = Trim$(f.cboType.Value)
    If IsIncoming(f) And Len(Trim$(f.txtNo.Value)) > 0 Then rw.Cells(1, C_NO).Value = Trim$(f.txtNo.Value)
    If isNew Then                                  ' a freshly added row may lack the formulas
        If Not rw.Cells(1, C_SEQ).HasFormula Then RestoreFormula lo, r, C_SEQ
        If Not rw.Cells(1, C_NO).HasFormula And Len(Txt(rw.Cells(1, C_NO).Value)) = 0 Then RestoreFormula lo, r, C_NO
        If Not rw.Cells(1, C_DATE).HasFormula And Len(Txt(rw.Cells(1, C_DATE).Value)) = 0 Then RestoreFormula lo, r, C_DATE
    End If
    rw.Calculate
    sNo = Txt(rw.Cells(1, C_NO).Value)
    Application.ScreenUpdating = oldScr
    MsgBox U("062A 0645 062A 0020 0627 0644 0625 0636 0627 0641 0629 0020 0628 0646 062C 0627 062D 002E") & vbCrLf & U("0631 0642 0645 0020 0627 0644 062E 0637 0627 0628 003A 0020") & sNo & vbCrLf & U("0627 0644 0635 0641 003A 0020") & rw.Row, vbInformation + MB_RTL
    LoadCompanies f
    FormClear f
    FormSearch f
    Exit Sub
Fail:
    Application.ScreenUpdating = oldScr
    MsgBox U("062A 0639 0630 0631 062A 0020 0627 0644 0625 0636 0627 0641 0629 003A 0020") & Err.Description, vbCritical + MB_RTL
End Sub

Public Sub FormSave(f As Object)
    Dim lo As ListObject, rw As Range, sNo As String
    Dim oldScr As Boolean, warn As String
    If mRow = 0 Then
        MsgBox U("0627 062E 062A 0631 0020 062E 0637 0627 0628 0627 0020 0645 0646 0020 0627 0644 0642 0627 0626 0645 0629 0020 0623 0648 0644 0627 002E"), vbExclamation + MB_RTL
        Exit Sub
    End If
    If Not Validate(f) Then Exit Sub
    Set lo = GetTable()
    If lo Is Nothing Then Exit Sub
    If Not RowStillValid(lo) Then Exit Sub
    oldScr = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Set rw = lo.ListRows(mRow).Range
    rw.Cells(1, C_COMP).Value = Trim$(f.cboCompany.Value)
    rw.Cells(1, C_TITLE).Value = Trim$(f.txtTitle.Value)
    rw.Cells(1, C_TYPE).Value = Trim$(f.cboType.Value)
    ' letter number (column B): only touched if the user changed it, or the type is now outgoing
    sNo = Trim$(f.txtNo.Value)
    If IsIncoming(f) Then
        If sNo <> mOrigNo Then
            If Len(sNo) = 0 Then
                If Not RestoreFormula(lo, mRow, C_NO) Then warn = warn & U("062A 0639 0630 0631 062A 0020 0625 0639 0627 062F 0629 0020 0645 0639 0627 062F 0644 0629 0020 0631 0642 0645 0020 0627 0644 062E 0637 0627 0628 002E") & vbCrLf
            Else
                rw.Cells(1, C_NO).Value = sNo
            End If
        End If
    ElseIf Norm(mOrigType) <> Norm(Trim$(f.cboType.Value)) And Not rw.Cells(1, C_NO).HasFormula Then   ' type just changed to outgoing
        If Not RestoreFormula(lo, mRow, C_NO) Then warn = warn & U("062A 0639 0630 0631 062A 0020 0625 0639 0627 062F 0629 0020 0645 0639 0627 062F 0644 0629 0020 0631 0642 0645 0020 0627 0644 062E 0637 0627 0628 002E") & vbCrLf
    End If
    rw.Calculate
    Application.ScreenUpdating = oldScr
    MsgBox U("062A 0645 0020 062D 0641 0638 0020 0627 0644 062A 0639 062F 064A 0644 002E") & IIf(Len(warn) > 0, vbCrLf & warn, ""), vbInformation + MB_RTL
    LoadCompanies f
    FormClear f
    FormSearch f
    Exit Sub
Fail:
    Application.ScreenUpdating = oldScr
    MsgBox U("062A 0639 0630 0631 0020 0627 0644 062D 0641 0638 003A 0020") & Err.Description, vbCritical + MB_RTL
End Sub

Private Function HasRowsBelow(lo As ListObject, ByVal r As Long) As Boolean
    Dim i As Long
    For i = r + 1 To lo.ListRows.Count
        If Len(Txt(lo.DataBodyRange.Cells(i, C_COMP).Value)) > 0 Then HasRowsBelow = True: Exit Function
    Next i
End Function

Public Sub FormDelete(f As Object)
    Dim lo As ListObject, rw As Range, oldScr As Boolean, warn As String, note As String
    If mRow = 0 Then
        MsgBox U("0627 062E 062A 0631 0020 062E 0637 0627 0628 0627 0020 0645 0646 0020 0627 0644 0642 0627 0626 0645 0629 0020 0623 0648 0644 0627 002E"), vbExclamation + MB_RTL
        Exit Sub
    End If
    Set lo = GetTable()
    If lo Is Nothing Then Exit Sub
    If Not RowStillValid(lo) Then Exit Sub
    note = ""
    If HasRowsBelow(lo, mRow) Then note = vbCrLf & vbCrLf & U("062A 0646 0628 064A 0647 003A 0020 0623 0631 0642 0627 0645 0020 0627 0644 062E 0637 0627 0628 0627 062A 0020 0627 0644 062A 0644 0642 0627 0626 064A 0629 0020 0627 0644 0644 0627 062D 0642 0629 0020 0642 062F 0020 062A 062A 063A 064A 0631 0020 0644 0623 0646 0020 0627 0644 062A 0631 0642 064A 0645 0020 064A 0639 062A 0645 062F 0020 0639 0644 0649 0020 062A 0631 062A 064A 0628 0020 0627 0644 0635 0641 0648 0641 002E")
    If MsgBox(U("0647 0644 0020 062A 0631 064A 062F 0020 062D 0630 0641 0020 0628 064A 0627 0646 0627 062A 0020 0647 0630 0627 0020 0627 0644 062E 0637 0627 0628 061F") & vbCrLf & vbCrLf & mOrigComp & vbCrLf & mOrigTitle & vbCrLf & vbCrLf & U("0633 064A 062A 0645 0020 0645 0633 062D 0020 0627 0644 0634 0631 0643 0629 0020 0648 0627 0644 0639 0646 0648 0627 0646 0020 0648 0627 0644 0646 0648 0639 0020 0648 0627 0644 0645 0644 0627 062D 0638 0627 062A 0020 0641 0642 0637 060C 0020 0648 0644 0646 0020 064A 062D 0630 0641 0020 0627 0644 0635 0641 002E") & note, _
              vbYesNo + vbQuestion + vbDefaultButton2 + MB_RTL) <> vbYes Then Exit Sub
    oldScr = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Set rw = lo.ListRows(mRow).Range
    rw.Cells(1, C_COMP).ClearContents
    rw.Cells(1, C_TITLE).ClearContents
    rw.Cells(1, C_TYPE).ClearContents
    rw.Cells(1, C_NOTES).ClearContents
    If Not RestoreFormula(lo, mRow, C_NO) Then warn = warn & U("062A 0639 0630 0631 062A 0020 0625 0639 0627 062F 0629 0020 0645 0639 0627 062F 0644 0629 0020 0631 0642 0645 0020 0627 0644 062E 0637 0627 0628 002E") & vbCrLf
    If Not RestoreFormula(lo, mRow, C_DATE) Then warn = warn & U("062A 0639 0630 0631 062A 0020 0625 0639 0627 062F 0629 0020 0645 0639 0627 062F 0644 0629 0020 0627 0644 062A 0627 0631 064A 062E 002E") & vbCrLf
    rw.Calculate
    Application.ScreenUpdating = oldScr
    MsgBox U("062A 0645 0020 0627 0644 062D 0630 0641 002E") & IIf(Len(warn) > 0, vbCrLf & warn, ""), vbInformation + MB_RTL
    LoadCompanies f
    FormClear f
    FormSearch f
    Exit Sub
Fail:
    Application.ScreenUpdating = oldScr
    MsgBox U("062A 0639 0630 0631 0020 0627 0644 062D 0630 0641 003A 0020") & Err.Description, vbCritical + MB_RTL
End Sub

' ---------------------------------------------------------------
' one-time setup: builds the UserForm and the sheet button
' ---------------------------------------------------------------
Public Sub SetupLetters()
    Dim vbp As Object
    On Error Resume Next
    Set vbp = ThisWorkbook.VBProject
    If Err.Number <> 0 Then Set vbp = Nothing
    On Error GoTo 0
    If vbp Is Nothing Then
        MsgBox U("0641 0639 0644 0020 0627 0644 062E 064A 0627 0631 003A 0020 0645 0644 0641 0020 003E 0020 062E 064A 0627 0631 0627 062A 0020 003E 0020 0645 0631 0643 0632 0020 0627 0644 062A 0648 062B 064A 0642 0020 003E 0020 0625 0639 062F 0627 062F 0627 062A 0020 0645 0631 0643 0632 0020 0627 0644 062A 0648 062B 064A 0642 0020 003E 0020 0625 0639 062F 0627 062F 0627 062A 0020 0627 0644 0645 0627 0643 0631 0648 0020 003E 0020 0627 0644 062B 0642 0629 0020 0641 064A 0020 0627 0644 0648 0635 0648 0644 0020 0625 0644 0649 0020 0646 0645 0648 0630 062C 0020 0643 0627 0626 0646 0020 0645 0634 0631 0648 0639 0020 0056 0042 0041 060C 0020 062B 0645 0020 0623 0639 062F 0020 062A 0634 063A 064A 0644 0020 0053 0065 0074 0075 0070 004C 0065 0074 0074 0065 0072 0073 002E"), vbExclamation + MB_RTL
        Exit Sub
    End If
    If GetTable() Is Nothing Then Exit Sub

    On Error GoTo FormFail
    BuildForm vbp

    On Error GoTo BtnFail
    mStep = "button"
    AddSheetButton
    MsgBox U("062A 0645 0020 0625 0646 0634 0627 0621 0020 0627 0644 0646 0627 0641 0630 0629 0020 0648 0627 0644 0632 0631 002E 0020 0627 0636 063A 0637 0020 0632 0631 0020 0625 062F 062E 0627 0644 0020 002F 0020 0628 062D 062B 0020 0641 064A 0020 0627 0644 0635 0641 0020 0627 0644 0623 0648 0644 002E"), vbInformation + MB_RTL
    Exit Sub

FormFail:
    MsgBox U("062A 0639 0630 0631 0020 0625 0646 0634 0627 0621 0020 0627 0644 0646 0627 0641 0630 0629 002E") & vbCrLf & U("0627 0644 062E 0637 0648 0629 003A 0020") & mStep & vbCrLf & U("0631 0642 0645 0020 0627 0644 062E 0637 0623 003A 0020") & Err.Number & vbCrLf & Err.Description, vbCritical + MB_RTL
    Exit Sub

BtnFail:
    MsgBox U("062A 0645 0020 0625 0646 0634 0627 0621 0020 0627 0644 0646 0627 0641 0630 0629 060C 0020 0644 0643 0646 0020 062A 0639 0630 0631 0020 0625 0646 0634 0627 0621 0020 0627 0644 0632 0631 002E") & vbCrLf & U("0627 0644 062E 0637 0648 0629 003A 0020") & mStep & vbCrLf & U("0631 0642 0645 0020 0627 0644 062E 0637 0623 003A 0020") & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & U("0623 0646 0634 0626 0020 0632 0631 0627 0020 064A 062F 0648 064A 0627 003A 0020 0645 0637 0648 0631 0020 003E 0020 0625 062F 0631 0627 062C 0020 003E 0020 0632 0631 060C 0020 0648 0627 0631 0628 0637 0647 0020 0628 0627 0644 0645 0627 0643 0631 0648 0020 0053 0068 006F 0077 004C 0065 0074 0074 0065 0072 0073 0046 006F 0072 006D 002E"), vbExclamation + MB_RTL
End Sub

Private Function AddCtl(d As Object, ByVal progId As String, ByVal nm As String, _
                        ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                        Optional ByVal cap As String = "") As Object
    Dim c As Object
    mStep = "control " & nm
    Set c = d.Controls.Add(progId, nm, True)
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    On Error Resume Next                          ' not every property exists on every control type
    c.Font.Name = "Tahoma"
    c.Font.Size = 10
    c.RightToLeft = True
    c.TextAlign = 3                               ' right
    If Len(cap) > 0 Then c.Caption = cap
    On Error GoTo 0
    Set AddCtl = c
End Function

' our form from an earlier setup (any name), or an empty leftover UserForm
Private Function IsLettersForm(comp As Object) As Boolean
    Dim n As Long
    If comp.Name Like FORM_NAME & "*" Or comp.Name Like "zzOldForm*" Then IsLettersForm = True: Exit Function
    On Error Resume Next
    n = comp.Designer.Controls.Count
    If Err.Number <> 0 Then Exit Function
    If HasCtl(comp.Designer, "lstResults") Then IsLettersForm = True: Exit Function
    If n = 0 And comp.Name Like "UserForm#*" Then IsLettersForm = True
End Function

Private Sub Paint(c As Object, ByVal back As Long, ByVal fore As Long, Optional ByVal bold As Boolean = False)
    On Error Resume Next
    c.BackColor = back
    c.ForeColor = fore
    If bold Then c.Font.Bold = True
End Sub

Private Sub Ink(c As Object)                          ' transparent label in brand colour
    On Error Resume Next
    c.BackStyle = 0
    c.ForeColor = CLR_TEAL
    c.Font.Bold = True
End Sub

Private Sub BuildForm(vbp As Object)
    Dim comp As Object, d As Object, c As Object, i As Long, k As Long, x As Single
    Dim heads As Variant, widths As Variant

    mStep = "remove old form"
    For i = vbp.VBComponents.Count To 1 Step -1
        Set comp = vbp.VBComponents(i)
        If comp.Type = 3 Then                       ' MSForm
            If IsLettersForm(comp) Then vbp.VBComponents.Remove comp   ' carried out when the macro ends
        End If
    Next i

    mStep = "add form"
    Set comp = vbp.VBComponents.Add(3)             ' vbext_ct_MSForm
    mStep = "rename form"
    ' A name freed by Remove may still be reserved, so try a few names.
    ' If none is accepted the automatic name stays; ShowLettersForm finds the form by content.
    On Error Resume Next
    For k = 0 To 9
        Err.Clear
        If k = 0 Then comp.Name = FORM_NAME Else comp.Name = FORM_NAME & (k + 1)
        If Err.Number = 0 Then Exit For
    Next k
    Err.Clear
    On Error GoTo 0
    On Error Resume Next
    comp.Properties("Caption").Value = U("0625 062F 062E 0627 0644 0020 0648 0627 0644 0628 062D 062B 0020 0641 064A 0020 0627 0644 062E 0637 0627 0628 0627 062A")
    comp.Properties("Width").Value = 910
    comp.Properties("Height").Value = 555
    comp.Properties("StartUpPosition").Value = 1
    comp.Properties("RightToLeft").Value = True
    comp.Properties("BackColor").Value = CLR_BG
    On Error GoTo 0
    Set d = comp.Designer

    ' --- header band: white with the logo, teal rule underneath ---
    Set c = AddCtl(d, "Forms.Label.1", "lblBand", 0, 0, 910, 72)
    Paint c, CLR_WHITE, CLR_WHITE
    Set c = AddCtl(d, "Forms.Label.1", "lblBandLine", 0, 72, 910, 4)
    Paint c, CLR_TEAL, CLR_TEAL
    Set c = AddCtl(d, "Forms.Label.1", "lblBrand", 440, 22, 440, 28, U("0634 0631 0643 0629 0020 0639 0644 064A 0020 0627 0628 0631 0627 0647 064A 0645 0020 0627 0644 0631 0628 064A 0634 064A 0020 0627 0644 0639 0642 0627 0631 064A 0629"))
    Paint c, CLR_WHITE, CLR_TEAL, True
    On Error Resume Next
    c.BackStyle = 0: c.Font.Size = 16
    On Error GoTo 0
    Set c = AddCtl(d, "Forms.Label.1", "lblSub", 15, 24, 380, 26, U("0627 0644 062E 0637 0627 0628 0627 062A 0020 0627 0644 0635 0627 062F 0631 0629 0020 0648 0627 0644 0648 0627 0631 062F 0629"))
    Paint c, CLR_WHITE, CLR_TEAL_MID, True
    On Error Resume Next
    c.BackStyle = 0: c.TextAlign = 1: c.Font.Size = 16
    On Error GoTo 0

    ' --- logo: embedded into the form now, so the file is not needed afterwards ---
    mStep = "logo"
    Dim fso As Object, logoPath As String, logoOk As Boolean
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    logoPath = ThisWorkbook.Path & "\Logo.png"
    If Not fso.FileExists(logoPath) Then logoPath = U("0043 003A 005C 0055 0073 0065 0072 0073 005C 004D 006F 0068 0061 006D 006D 0065 0064 0046 0061 0079 0079 0061 0064 005C 004F 006E 0065 0044 0072 0069 0076 0065 0020 002D 0020 0041 006C 0072 0075 0062 0061 0069 0073 0068 0069 0020 0048 006F 006C 0064 0069 006E 0067 0020 0043 006F 006D 0070 0061 006E 0079 005C 0627 0644 0631 0628 064A 0634 064A 0020 0627 0644 0639 0642 0627 0631 064A 0629 005C 0627 0644 0634 0624 0648 0646 0020 0627 0644 0625 062F 0627 0631 064A 0629 005C 062E 0637 0627 0628 0627 062A 005C 004C 006F 0067 006F 002E 0070 006E 0067")
    If fso.FileExists(logoPath) Then
        Err.Clear
        Set c = AddCtl(d, "Forms.Image.1", "imgLogo", 745, 4, 135, 64)
        c.Picture = LoadPicture(logoPath)
        c.PictureSizeMode = 1                       ' zoom, keeps the proportions
        c.BackStyle = 0
        c.BorderStyle = 0
        logoOk = (Err.Number = 0)
        If Not logoOk Then c.Visible = False
    End If
    Err.Clear
    d.Controls("lblBrand").Visible = Not logoOk     ' company name text only when there is no logo
    On Error GoTo 0

    ' --- inputs: right pair (label L=775, input L=485) / left pair (label L=310, input L=15) ---
    Ink AddCtl(d, "Forms.Label.1", "lblCompany", 775, 90, 105, 18, U("0627 0644 0634 0631 0643 0629 0020 002F 0020 0627 0644 062C 0647 0629"))
    Set c = AddCtl(d, "Forms.ComboBox.1", "cboCompany", 485, 88, 290, 22)
    On Error Resume Next
    c.Style = 0: c.MatchEntry = 1                  ' free text + auto-complete from previous companies
    On Error GoTo 0

    Ink AddCtl(d, "Forms.Label.1", "lblType", 310, 90, 105, 18, U("0635 0627 062F 0631 0020 002F 0020 0648 0627 0631 062F"))
    Set c = AddCtl(d, "Forms.ComboBox.1", "cboType", 15, 88, 290, 22)
    On Error Resume Next
    c.Style = 2                                    ' drop-down list only
    On Error GoTo 0

    Ink AddCtl(d, "Forms.Label.1", "lblTitle", 775, 124, 105, 18, U("0639 0646 0648 0627 0646 0020 0627 0644 062E 0637 0627 0628"))
    AddCtl d, "Forms.TextBox.1", "txtTitle", 485, 122, 290, 22

    Ink AddCtl(d, "Forms.Label.1", "lblNo", 310, 124, 105, 18, U("0631 0642 0645 0020 0627 0644 062E 0637 0627 0628"))
    AddCtl d, "Forms.TextBox.1", "txtNo", 15, 122, 290, 22
    d.Controls("lblNo").Visible = False
    d.Controls("txtNo").Visible = False
    d.Controls("txtNo").ControlTipText = U("0644 0644 062E 0637 0627 0628 0020 0627 0644 0648 0627 0631 062F 0020 0641 0642 0637 0020 002D 0020 0627 062E 062A 064A 0627 0631 064A 002E 0020 0627 062A 0631 0643 0647 0020 0641 0627 0631 063A 0627 0020 0644 064A 0628 0642 0649 0020 0627 0644 0631 0642 0645 0020 0627 0644 062A 0644 0642 0627 0626 064A 002E")

    ' --- search row ---
    Set c = AddCtl(d, "Forms.Label.1", "lblRule", 15, 158, 865, 1)
    Paint c, CLR_LINE, CLR_LINE
    Ink AddCtl(d, "Forms.Label.1", "lblSearch", 775, 174, 105, 18, U("0628 062D 062B 0020 0639 0646"))
    AddCtl d, "Forms.TextBox.1", "txtSearch", 485, 172, 290, 22
    d.Controls("txtSearch").ControlTipText = U("062C 0632 0621 0020 0645 0646 0020 0627 0644 0634 0631 0643 0629 0020 0623 0648 0020 0627 0644 0631 0642 0645 0020 0623 0648 0020 0627 0644 0639 0646 0648 0627 0646 0020 0623 0648 0020 0627 0644 0645 0644 0627 062D 0638 0627 062A 0020 0623 0648 0020 0627 0644 062A 0627 0631 064A 062E")
    Ink AddCtl(d, "Forms.Label.1", "lblFilter", 310, 174, 105, 18, U("062A 0635 0641 064A 0629 0020 0627 0644 0646 0648 0639"))
    AddCtl d, "Forms.ComboBox.1", "cboFilter", 15, 172, 290, 22
    On Error Resume Next
    d.Controls("cboFilter").Style = 2
    On Error GoTo 0

    ' --- buttons, right to left ---
    Dim caps As Variant, names As Variant, backs As Variant, fores As Variant
    names = Array("cmdAdd", "cmdSearch", "cmdSave", "cmdClear", "cmdDelete", "cmdClose")
    caps = Array(U("0625 0636 0627 0641 0629"), U("0628 062D 062B"), U("062D 0641 0638 0020 0627 0644 062A 0639 062F 064A 0644"), U("0645 0633 062D 0020 0627 0644 062E 0627 0646 0627 062A"), U("062D 0630 0641"), U("0625 063A 0644 0627 0642"))
    backs = Array(CLR_TEAL, CLR_TEAL_MID, CLR_TEAL, CLR_PALE, CLR_DANGER, CLR_PALE)
    fores = Array(CLR_WHITE, CLR_WHITE, CLR_WHITE, CLR_TEAL, CLR_WHITE, CLR_TEAL)
    For i = 0 To 5
        Set c = AddCtl(d, "Forms.CommandButton.1", CStr(names(i)), 755 - i * 135, 210, 125, 32, CStr(caps(i)))
        Paint c, CLng(backs(i)), CLng(fores(i)), True
    Next i

    ' --- column headers: the list is right-to-left, so the first column sits at the right edge ---
    heads = Array(U("0023"), U("0631 0642 0645 0020 0627 0644 062E 0637 0627 0628"), U("0627 0644 0634 0631 0643 0629 0020 002F 0020 0627 0644 062C 0647 0629"), U("0639 0646 0648 0627 0646 0020 0627 0644 062E 0637 0627 0628"), U("0627 0644 0646 0648 0639"), U("0627 0644 062A 0627 0631 064A 062E"))
    widths = Array(45, 135, 170, 330, 60, 90)
    x = 877                                        ' right edge of the list's inner area
    For i = 0 To 5
        x = x - widths(i)
        Set c = AddCtl(d, "Forms.Label.1", "lblHead" & i, x, 256, widths(i), 20, CStr(heads(i)))
        Paint c, CLR_TEAL, CLR_WHITE, True
        On Error Resume Next
        c.TextAlign = 2: c.BorderStyle = 0
        On Error GoTo 0
    Next i

    Set c = AddCtl(d, "Forms.ListBox.1", "lstResults", 15, 276, 865, 210)
    On Error Resume Next
    c.RightToLeft = True
    c.ColumnCount = 7
    c.ColumnWidths = "45;135;170;330;60;90;0"
    On Error GoTo 0

    ' --- vertical grid lines between the columns (a ListBox cannot draw them itself) ---
    x = 877
    For i = 0 To 4
        x = x - widths(i)
        mStep = "grid line " & i
        Set c = d.Controls.Add("Forms.Label.1", "lnSep" & i, True)
        c.Left = x: c.Top = 278: c.Width = 1: c.Height = 206
        On Error Resume Next
        c.Caption = "": c.BackColor = CLR_LINE: c.BorderStyle = 0: c.SpecialEffect = 0
        On Error GoTo 0
    Next i

    Set c = AddCtl(d, "Forms.Label.1", "lblCount", 15, 492, 865, 16, U("0639 062F 062F 0020 0627 0644 0646 062A 0627 0626 062C 003A 0020 0030"))
    Ink c

    mStep = "form code"
    comp.CodeModule.AddFromString FormCode()
End Sub

Private Function FormCode() As String
    Dim s As String
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    FormInit Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cmdAdd_Click()" & vbCrLf
    s = s & "    FormAdd Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cmdSearch_Click()" & vbCrLf
    s = s & "    FormSearch Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cmdSave_Click()" & vbCrLf
    s = s & "    FormSave Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cmdClear_Click()" & vbCrLf
    s = s & "    FormClear Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cmdDelete_Click()" & vbCrLf
    s = s & "    FormDelete Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cmdClose_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub cboType_Change()" & vbCrLf
    s = s & "    FormTypeChanged Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub lstResults_Click()" & vbCrLf
    s = s & "    FormPick Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" & vbCrLf
    s = s & "    If KeyCode = 13 Then FormSearch Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FormCode = s
End Function

Private Sub AddSheetButton()
    Dim ws As Worksheet, lastCol As Long, ma As Range, nxt As Range, shp As Shape, h As Single
    mStep = "button: sheet"
    Set ws = ThisWorkbook.Worksheets(U("0627 0644 0639 0642 0627 0631 064A 0629"))
    mStep = "button: delete old"
    On Error Resume Next
    ws.Shapes(BTN_NAME).Delete
    On Error GoTo 0
    mStep = "button: position"
    On Error Resume Next
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set ma = ws.Cells(1, lastCol).MergeArea
    Set nxt = ws.Cells(1, ma.Column + ma.Columns.Count)
    On Error GoTo 0
    If nxt Is Nothing Then Set nxt = ws.Range("H1")
    h = ws.Rows(1).Height - 2
    If h < 18 Then h = 18
    If h > 26 Then h = 26
    mStep = "button: create"
    Set shp = ws.Shapes.AddFormControl(0, nxt.Left + 2, nxt.Top + 1, 110, h)    ' 0 = xlButtonControl
    shp.Name = BTN_NAME
    mStep = "button: macro link"
    ' qualify with the workbook name: another workbook may be the active one
    On Error Resume Next
    shp.OnAction = "'" & ThisWorkbook.Name & "'!ShowLettersForm"
    If Err.Number <> 0 Then
        Err.Clear
        shp.OnAction = "ShowLettersForm"
    End If
    If Err.Number <> 0 Then
        Dim oaNum As Long, oaMsg As String
        oaNum = Err.Number: oaMsg = Err.Description
        On Error GoTo 0
        Err.Raise oaNum, , oaMsg
    End If
    On Error GoTo 0
    On Error Resume Next
    shp.Placement = 2                               ' xlMove
    shp.TextFrame.Characters.Text = U("0625 062F 062E 0627 0644 0020 002F 0020 0628 062D 062B")
    shp.TextFrame.Characters.Font.Size = 10
    On Error GoTo 0
End Sub

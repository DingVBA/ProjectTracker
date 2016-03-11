Attribute VB_Name = "Main"
Public Const STARTROW = 5
Public Const ENROLLLEN = 12
Public Const LSHEETNAME = "LGE Service Center Project List"
Public Const SHEETNAME = "Tracking"

Sub unique()
  Dim Enrollments As New Collection
  Dim newE As New Collection
  
  Dim wb As Workbook
  Dim wbl As Workbook
  Dim ws As Worksheet
  Dim wsl As Worksheet
  Dim w As String
  
Application.EnableEvents = False
Application.DisplayAlerts = False
Application.ScreenUpdating = False
  On Error Resume Next
  Set wb = ThisWorkbook
  Set ws = wb.Worksheets(SHEETNAME)
  lastrow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
  oldlastrow = lastrow
  
  Call GetExistingEnrollments(Enrollments)
  
  w = Application.GetOpenFilename(Title:="Select the Project List file", MultiSelect:=False)
  If w = "False" Then Exit Sub
  Set wbl = Workbooks.Open(Filename:=w)
  Set wsl = wbl.Worksheets(LSHEETNAME)
  llastrow = wsl.Range("A" & wsl.Rows.Count).End(xlUp).Row
  
  Call GetProjectList(newE, wbl)
  
  For Each a In newE
    If Not Contains(Enrollments, a) Then
        Enrollments.Add a, CStr(a)
        ws.Cells(lastrow + 1, TRCol1.Enrollment_ID).Value = GetEnrollmentString(a)
        ws.Cells(lastrow + 1, TRCol1.Enrollment_ID).Interior.ColorIndex = 37
        ApptDate = Application.WorksheetFunction.VLookup(a, wsl.Range("B2:BQ" & llastrow), PLCol1.Schedule_date - 1, False)
        Auditor = Application.WorksheetFunction.VLookup(a, wsl.Range("B2:BQ" & llastrow), PLCol1.First_and_last_name_of_main_auditor - 1, False)
        CustomerName = Application.WorksheetFunction.VLookup(a, wsl.Range("B2:BQ" & llastrow), PLCol1.Remit_to_contact_name - 1, False)
        Address = Application.WorksheetFunction.VLookup(a, wsl.Range("B2:BQ" & llastrow), PLCol1.Remit_to_contact_street_address - 1, False)
        sStatus = Application.WorksheetFunction.VLookup(a, wsl.Range("B2:BQ" & llastrow), PLCol1.Status - 1, False)
        Select Case sStatus
            Case "SUSPENSE", "COMPLETE"
                FAStatus = "Closed"
            Case "CANCELLED"
                FAStatus = "CANCELLED"
            Case "SCHEDULED"
                FAStatus = "HOLD"
        End Select
        ws.Cells(lastrow + 1, TRCol1.F_A_Status).Value = FAStatus
        ws.Cells(lastrow + 1, TRCol1.Analyst).Value = Auditor
        ws.Cells(lastrow + 1, TRCol1.Appt_Date).Value = GetJudiDateFormat(ApptDate)
        ws.Cells(lastrow + 1, TRCol1.Customer_Name).Value = CustomerName
        ws.Cells(lastrow + 1, TRCol1.Street_Address).Value = Address
        ws.Cells(lastrow + 1, TRCol1.Enrollment_Status).Value = sStatus
        ws.Cells(lastrow, TRCol1.End_Date).Copy
        ws.Cells(lastrow + 1, TRCol1.End_Date).PasteSpecial Paste:=xlPasteFormulas
        ws.Cells(lastrow, TRCol1.Enrollment_ID_Duplicate).Copy
        ws.Cells(lastrow + 1, TRCol1.Enrollment_ID_Duplicate).PasteSpecial Paste:=xlPasteFormulas
        ws.Cells(lastrow, TRCol1.Nexant_Project_ID).Copy
        ws.Cells(lastrow + 1, TRCol1.Nexant_Project_ID).PasteSpecial Paste:=xlPasteFormulas
        ws.Cells(lastrow, TRCol1.Project_ID).AutoFill _
        Destination:=ws.Range(ws.Cells(lastrow, TRCol1.Project_ID), ws.Cells(lastrow + 1, TRCol1.Project_ID)), Type:=xlFillDefault
        lastrow = lastrow + 1
    Else
        Debug.Print a
    End If
  Next
  Call SetProjectListFormat(wsl)
  
  wsl.Range("$A$1:$BQ$" & llastrow).AutoFilter Field:=PLCol1.short_program_name, Criteria1:="HEAP"
  wsl.Range("$A$1:$BQ$" & llastrow).AutoFilter Field:=PLCol1.Status, Criteria1:=Array( _
        "COMPLETE", "SCHEDULED", "SITE WORK COMPLETE"), Operator:=xlFilterValues
        
  wbl.SaveAs wbl.Path & "\" & Replace(wbl.Name, ".csv", ""), 51
  wbl.Close SaveChanges:=False
  Set wbl = Nothing
Application.EnableEvents = True
Application.DisplayAlerts = True
Application.ScreenUpdating = True

enrollmentnum = lastrow - oldlastrow
If enrollmentnum > 0 Then
    MsgBox (CStr(enrollmentnum) + " enrollments are loaded.")
Else
    MsgBox ("No enrollment is loaded.")
End If
End Sub

Public Sub SetProjectListFormat(ByRef wsl As Worksheet)
    wsl.Columns("A:BQ").EntireColumn.Hidden = True
    wsl.Columns("A:B").EntireColumn.Hidden = False
    wsl.Columns("G").EntireColumn.Hidden = False
    wsl.Columns("R").EntireColumn.Hidden = False
    wsl.Columns("AD").EntireColumn.Hidden = False
    wsl.Columns("AO").EntireColumn.Hidden = False
    wsl.Columns("AT").EntireColumn.Hidden = False
    wsl.Columns("AU").EntireColumn.Hidden = False
    
    wsl.Columns("G").ColumnWidth = 9
    wsl.Columns("G").ColumnWidth = 13
    wsl.Columns("R").ColumnWidth = 18
    wsl.Columns("AD").ColumnWidth = 18
    wsl.Columns("AO").ColumnWidth = 10
    wsl.Columns("AT").ColumnWidth = 26
    wsl.Columns("AU").ColumnWidth = 36
    ActiveWindow.LargeScroll ToRight:=-1
End Sub
Public Function GetJudiDateFormat(ByVal str As String) As String
    sYear = Mid(str, 1, 4)
    sMonth = Mid(str, 5, 2)
    sDay = Mid(str, 7, 2)
    GetJudiDateFormat = sYear + "-" + sMonth + "-" + sDay
End Function
Public Function GetEnrollmentString(ByVal lng As Long) As String
    Dim str As String
    Dim i As Integer
    
    str = CStr(lng)
    
    If Len(str) < ENROLLLEN Then
        For i = Len(str) + 1 To ENROLLLEN
            str = "0" + str
        Next i
    End If
    GetEnrollmentString = str
End Function
Public Sub GetExistingEnrollments(ByRef colMyCol As Collection)
    Dim wbLoad As Workbook
    Dim ws As Worksheet
    Set wbLoad = ThisWorkbook
    Set ws = wbLoad.Worksheets(SHEETNAME)

    Dim rRange As Range
    Dim rCell As Range
    Set rRange = ws.Range("K" & STARTROW)
    
    If Len(rRange.Value) = 0 Then GoTo BeforeExit
    If Len(rRange.Offset(1, 0).Value) > 0 Then
       Set rRange = Range(rRange, rRange.End(xlDown))
    End If
        
    On Error Resume Next
    For Each rCell In rRange
        c = CLng(rCell.Value)
        colMyCol.Add c, CStr(c)
    Next
    
BeforeExit:
    Exit Sub
ErrorHandle:
    MsgBox (err.Description & " Error in procedure GetExistingEnrollments ")
    Resume BeforeExit
End Sub

Public Sub GetProjectList(ByRef arr As Collection, ByRef wbLoad As Workbook)
    Dim i As Long
    Set ETab = wbLoad.Worksheets(LSHEETNAME)
    lastrow = ETab.Range("A" & ETab.Rows.Count).End(xlUp).Row

    FrmProgress.Show vbModeless
    For i = 2 To lastrow
        Call UpdateProgressBar(i * 1# / lastrow, "processing...")
        On Error Resume Next
        pstatus = ETab.Cells(i, PLCol1.Status).Value
        progname = ETab.Cells(i, PLCol1.short_program_name).Value
        If (pstatus = "SITE WORK COMPLETE" Or pstatus = "SCHEDULED" Or pstatus = "COMPLETE") And progname = "HEAP" Then
            a = CLng(ETab.Cells(i, PLCol1.Enrollment_ID).Value)
            If Not Contains(arr, a) Then
                arr.Add a, CStr(a)
            End If
        End If
    Next i
    FrmProgress.Hide
End Sub

Public Sub UpdateProgressBar(FilePctDone As Single, ProgressCaption As String)
    With FrmProgress
        ' Update the Caption property of the Frame control.
        .FilesProgressLabel.Caption = Format(FilePctDone, "0%") & " - " & ProgressCaption
        ' Widen the Label control.
        .FilesProgressBar.Width = FilePctDone * (.FilesProgressLabel.Width - 10)
    End With
    ' The DoEvents allows the UserForm to update.
    DoEvents
End Sub
Public Function Contains(col As Collection, key As Variant) As Boolean
Dim obj As Variant
On Error GoTo err
    Contains = True
    obj = col(CStr(key))
    Exit Function
err:
    Contains = False
End Function

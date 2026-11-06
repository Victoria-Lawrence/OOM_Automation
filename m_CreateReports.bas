Attribute VB_Name = "m_CreateReports"
Sub Create_And_Save_Reports()

    Dim Currentbook As Workbook: Dim Newbook As Workbook
    Dim StopPath As String, LesserPath As String, LargeBPath As String, FlaggedPath As String, ZeroPath As String
    Dim PriorStopPath As String, PriorLesserPath As String
    Dim l As Long: Dim DataSetLength As Long: Dim lastRow As Long: Dim i As Long
    Dim pt As PivotTable
    Dim Rng As Range: Dim DateStamp As Range
    Dim MyCol As Long: Dim MyRow As Long
    Dim LastCol As Long: Dim CurColReserve As Long, CurColCases As Long
 

'12.2.25 - VRL - Looks for the info in the cell to the right of the  (Sheet, Column, Text)
    StopPath = PullPath("Home", "G", "Stoploss Path:")
    LesserPath = PullPath("Home", "G", "Lesser Of Path:")
    LargeBPath = PullPath("Home", "G", "Large Balance Path:")
    FlaggedPath = PullPath("Home", "G", "Flagged Path:")
    ZeroPath = PullPath("Home", "G", "Zero Contractuals Path:")
    PriorStopPath = PullPath("Home", "G", "Prior Month Stoploss Path:")
    PriorLesserPath = PullPath("Home", "G", "Prior Lesser Of Path:")
    PriorLargeBPath = PullPath("Home", "G", "Prior Large Balance Path:")
    PriorFlaggedPath = PullPath("Home", "G", "Prior Flagged Path:")

    Set Currentbook = ActiveWorkbook
    Set DateStamp = Sheets("Home").Range("P3")

    'Creating Large Balance Report
    
    Sheets("LARGE BALANCE").Copy
    Set Newbook = ActiveWorkbook
    Range("A1").EntireRow.Delete
    
    Newbook.SaveAs (LargeBPath)
    Newbook.Close
    Currentbook.Activate
    
    'Creating Flagged Accounts Report
    
    Sheets("FLAGGED").Copy
    Set Newbook = ActiveWorkbook
    Range("A1").EntireRow.Delete
    
    Newbook.SaveAs (FlaggedPath)
    Newbook.Close
    Currentbook.Activate
    
    'Creating Zero Contractuals Report
    
    Sheets("ZERO CONTRACTUALS").Copy
    Set Newbook = ActiveWorkbook
    Range("A1").EntireRow.Delete
    
    Newbook.SaveAs (ZeroPath)
    Newbook.Close
    Currentbook.Activate
    
    '12.3.25 - VRL - Separating out Lesser of Report
        Set Newbook = Workbooks.Open(PriorLesserPath)
        Set ws = Workbooks.Open(PriorLesserPath).Sheets("LO_PATIENT_DETAIL_TABLE")
        'Sheets("LO_PATIENT_DETAIL_TABLE").Select
            'l = Range("A1000000").End(xlUp).Row
            
        'Find last row based on Column L
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ws.Select
    Range("A2:V" & lastRow).ClearContents
    
    Currentbook.Activate
    Sheets("LESSER OF").Select
    
    'Changes the My Blue from OON to IN

    DataSetLength = Range("A1000000").End(xlUp).Row
    Range("A3:V" & DataSetLength).Copy
    Newbook.Activate
    Range("A2").PasteSpecial (xlPasteAll)
    Application.CutCopyMode = False
    
    Range("A1").Select
    
    DataSetLength = DataSetLength - 1
    
        For i = 2 To lastRow 'Assuming row 1 is header
        If ws.Cells(i, "L").Value = 421564 Or ws.Cells(i, "L").Value = 136105 Then
            ws.Cells(i, "V").Value = "IN"
        End If
    Next i
    
    'Setting PivotTable source and refreshing Pivot
    Set Rng = Range("A1:V" & DataSetLength)

    Sheets("Summary").Select
    
    Range("A2").Select
    
    Set pt = ActiveCell.PivotTable
    pt.ChangePivotCache ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Rng)
    
       
    Sheets("Trending").Select
    
    Application.Calculation = xlCalculationAutomatic
    
    'Updating Lesser of Dollars
    MyCol = Range("B6").End(xlToRight).Column
    
    Cells(6, MyCol).EntireColumn.Insert
    Cells(6, MyCol).Value = DateStamp
    Cells(7, MyCol + 1).Select
    Range(Selection, Selection.End(xlDown)).Copy
    Cells(7, MyCol).PasteSpecial (xlPasteValues)
    Application.CutCopyMode = False
    
        'Updating Lesser Of Counts
    MyCol = Range("XFD6").End(xlToLeft).Column
    
    Cells(6, MyCol).EntireColumn.Insert
    Cells(6, MyCol).Value = DateStamp
    Cells(7, MyCol + 1).Select
    Range(Selection, Selection.End(xlDown)).Copy
    Cells(7, MyCol).PasteSpecial (xlPasteValues)
    Application.CutCopyMode = False
    
    '11.10.25 - VL - Update
' Update to hide everything but the needed columns and the most recent 6 months for both the count and the amount.


'Find last used column in row 6
LastCol = Cells(6, Columns.Count).End(xlToLeft).Column

'Find positions of "Cur" for Reserve and Cases
CurColReserve = 0
CurColCases = 0

For i = 1 To LastCol
    If UCase(Cells(6, i).Value) = "CUR" Then
        If CurColReserve = 0 Then
            CurColReserve = i 'First Cur for Reserve
        Else
            CurColCases = i 'Second Cur for Cases
            Exit For
        End If
    End If
Next i


'Unhide all first
Cells.EntireColumn.Hidden = False

'Always show A:B
Columns("A:B").Hidden = False

'--- Handle Reserve (Dollars) ---
If CurColReserve > 0 Then
    'Hide everything between column B and (CurColReserve - 6)
    For i = 3 To CurColReserve - 7
        Columns(i).Hidden = True
    Next i
End If

'--- Handle Cases (Counts) ---
If CurColCases > 0 Then
    'Hide everything before (CurColCases - 6) except A:B
    For i = CurColReserve + 2 To CurColCases - 7
        Columns(i).Hidden = True
    Next i
    
    'Hide everything after (CurColCases + 2)
    For i = CurColCases + 3 To LastCol
        Columns(i).Hidden = True
    Next i
End If
'End 11.10.25 Update
    
    Range("A1").Select

    Newbook.SaveAs (LesserPath)
    Newbook.Close

    'Creating StopLoss Lesser Of Report

    Set Newbook = Workbooks.Open(PriorStopPath)
    

    'Pulling Stop Loss Information into model

    Sheets("SL_PATIENT_DETAIL_TABLE").Select
    
    l = Range("A1000000").End(xlUp).Row
    
    Range("A2:AD" & l).ClearContents

    Currentbook.Activate
    Sheets("STOPLOSS").Select
    
    DataSetLength = Range("A1000000").End(xlUp).Row
    Range("A3:AD" & DataSetLength).Copy
    Newbook.Activate
    Range("A2").PasteSpecial (xlPasteAll)
    Application.CutCopyMode = False

    Range("A1").Select
    
    DataSetLength = DataSetLength - 1

    'Setting PivotTable source and refreshing Pivot
    Set Rng = Range("A1:AD" & DataSetLength)

    Sheets("Summary").Select
    
    Range("A4").Select
    
    Set pt = ActiveCell.PivotTable
    pt.ChangePivotCache ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Rng)
    
    pt.RefreshTable
    
    Sheets("Trending").Select
    
    Application.Calculation = xlCalculationAutomatic
    
    'Updating StopLoss Dollars
    MyCol = Range("B6").End(xlToRight).Column
    
    Cells(6, MyCol).EntireColumn.Insert
    Cells(6, MyCol).Value = DateStamp
    Cells(7, MyCol + 1).Select
    Range(Selection, Selection.End(xlDown)).Copy
    Cells(7, MyCol).PasteSpecial (xlPasteValues)
    Application.CutCopyMode = False

    'Updating StopLoss Counts
    MyCol = Range("XFD6").End(xlToLeft).Column
    
    Cells(6, MyCol).EntireColumn.Insert
    Cells(6, MyCol).Value = DateStamp
    Cells(7, MyCol + 1).Select
    Range(Selection, Selection.End(xlDown)).Copy
    Cells(7, MyCol).PasteSpecial (xlPasteValues)
    Application.CutCopyMode = False
    


'11.10.25 - VL - Update
' Update to hide everything but the needed columns and the most recent 6 months for both the count and the amount.


'Find last used column in row 6
LastCol = Cells(6, Columns.Count).End(xlToLeft).Column

'Find positions of "Cur" for Reserve and Cases
CurColReserve = 0
CurColCases = 0

For i = 1 To LastCol
    If UCase(Cells(6, i).Value) = "CUR" Then
        If CurColReserve = 0 Then
            CurColReserve = i 'First Cur for Reserve
        Else
            CurColCases = i 'Second Cur for Cases
            Exit For
        End If
    End If
Next i

'Unhide all first
Cells.EntireColumn.Hidden = False

'Always show A:B
Columns("A:B").Hidden = False

'--- Handle Reserve (Dollars) ---
If CurColReserve > 0 Then
    'Hide everything between column B and (CurColReserve - 6)
    For i = 3 To CurColReserve - 7
        Columns(i).Hidden = True
    Next i
End If

'--- Handle Cases (Counts) ---
If CurColCases > 0 Then
    'Hide everything before (CurColCases - 6) except A:B
    For i = CurColReserve + 2 To CurColCases - 7
        Columns(i).Hidden = True
    Next i
    
    'Hide everything after (CurColCases + 2)
    'For i = CurColCases + 3 To LastCol
    '    Columns(i).Hidden = True
    'Next i
End If
'End 11.10.25 Update
    
    Range("A1").Select

    Newbook.SaveAs (StopPath)
    Newbook.Close

    Sheets("Home").Select
    
    MsgBox ("The OOM Process is Complete")

End Sub


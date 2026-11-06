Attribute VB_Name = "m_CreateReports"
Sub Create_And_Save_Reports()

    Dim Currentbook As Workbook: Dim Newbook As Workbook
    Dim StopPath As String, LesserPath As String, LargeBPath As String, FlaggedPath As String, ZeroPath As String
    Dim PriorStopPath As String, PriorLesserPath As String
    Dim PriorLargeBPath As String, PriorFlaggedPath As String
    Dim l As Long: Dim DataSetLength As Long: Dim lastRow As Long: Dim i As Long
    Dim pt As PivotTable
    Dim Rng As Range: Dim DateStamp As Range
    Dim MyCol As Long: Dim MyRow As Long
    Dim LastCol As Long: Dim CurColReserve As Long, CurColCases As Long
    Dim ws As Worksheet
 

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

   '11.6.2026 - VL - Update for Trending for LargeB and Flagged per request from JHodges
   '=========================================================
' Update Large Balance report and Trended sheet
'=========================================================
Application.StatusBar = "Updating Large Balance report..."

UpdateReportWithTrend _
    SourceWorkbook:=Currentbook, _
    SourceSheetName:="LARGE BALANCE", _
    PriorReportPath:=PriorLargeBPath, _
    OutputReportPath:=LargeBPath, _
    DetailSheetName:="LARGE BALANCE", _
    TrendSheetName:="Trended", _
    FacilityHeader:="ChildFacilityName", _
    AccountHeader:="ntNumber", _
    BalanceHeader:="AccountBalance", _
    reportMonth:=DateStamp.value

Currentbook.Activate


'=========================================================
' Update Flagged Accounts report and Trended sheet
'=========================================================
Application.StatusBar = "Updating Flagged Accounts report..."

UpdateReportWithTrend _
    SourceWorkbook:=Currentbook, _
    SourceSheetName:="FLAGGED", _
    PriorReportPath:=PriorFlaggedPath, _
    OutputReportPath:=FlaggedPath, _
    DetailSheetName:="FLAGGED", _
    TrendSheetName:="Trended", _
    FacilityHeader:="Facility Name", _
    AccountHeader:="account number", _
    BalanceHeader:="account balance", _
    reportMonth:=DateStamp.value

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
        If ws.Cells(i, "L").value = 421564 Or ws.Cells(i, "L").value = 136105 Then
            ws.Cells(i, "V").value = "IN"
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
    Cells(6, MyCol).value = DateStamp
    Cells(7, MyCol + 1).Select
    Range(Selection, Selection.End(xlDown)).Copy
    Cells(7, MyCol).PasteSpecial (xlPasteValues)
    Application.CutCopyMode = False
    
        'Updating Lesser Of Counts
    MyCol = Range("XFD6").End(xlToLeft).Column
    
    Cells(6, MyCol).EntireColumn.Insert
    Cells(6, MyCol).value = DateStamp
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
    If UCase(Cells(6, i).value) = "CUR" Then
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
    Cells(6, MyCol).value = DateStamp
    Cells(7, MyCol + 1).Select
    Range(Selection, Selection.End(xlDown)).Copy
    Cells(7, MyCol).PasteSpecial (xlPasteValues)
    Application.CutCopyMode = False

    'Updating StopLoss Counts
    MyCol = Range("XFD6").End(xlToLeft).Column
    
    Cells(6, MyCol).EntireColumn.Insert
    Cells(6, MyCol).value = DateStamp
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
    If UCase(Cells(6, i).value) = "CUR" Then
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
Private Sub UpdateReportWithTrend( _
    ByVal SourceWorkbook As Workbook, _
    ByVal SourceSheetName As String, _
    ByVal PriorReportPath As String, _
    ByVal OutputReportPath As String, _
    ByVal DetailSheetName As String, _
    ByVal TrendSheetName As String, _
    ByVal FacilityHeader As String, _
    ByVal AccountHeader As String, _
    ByVal BalanceHeader As String, _
    ByVal reportMonth As Variant)

    Dim reportWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim DetailSheet As Worksheet
    Dim TrendSheet As Worksheet

    Dim sourceLastRow As Long
    Dim sourceLastColumn As Long

    Dim reportFileFormat As XlFileFormat
    Dim savedErrorNumber As Long
    Dim savedErrorDescription As String

    On Error GoTo ErrorHandler

    If Len(Trim$(PriorReportPath)) = 0 Then

        Err.Raise _
            vbObjectError + 2000, _
            "UpdateReportWithTrend", _
            "The prior report path is blank for " & _
            SourceSheetName & "."

    End If

    If Len(Dir$(PriorReportPath)) = 0 Then

        Err.Raise _
            vbObjectError + 2001, _
            "UpdateReportWithTrend", _
            "The prior report workbook could not be found:" & _
            vbCrLf & PriorReportPath

    End If

    Set sourceSheet = _
        SourceWorkbook.Worksheets(SourceSheetName)

    Set reportWorkbook = Workbooks.Open( _
        Filename:=PriorReportPath, _
        UpdateLinks:=0, _
        ReadOnly:=False, _
        AddToMru:=False)

    Set DetailSheet = _
        reportWorkbook.Worksheets(DetailSheetName)

    Set TrendSheet = _
        reportWorkbook.Worksheets(TrendSheetName)

    '--------------------------------------------------------
    ' Replace the detail data
    '--------------------------------------------------------
    DetailSheet.Cells.Clear

    sourceLastRow = LastUsedRowForReport(sourceSheet)
    sourceLastColumn = LastUsedColumnForReport(sourceSheet)

    sourceSheet.Range( _
        sourceSheet.Cells(1, 1), _
        sourceSheet.Cells( _
            sourceLastRow, _
            sourceLastColumn)).Copy

    DetailSheet.Range("A1").PasteSpecial xlPasteAll

    Application.CutCopyMode = False

    'The source worksheets have an extra first row.
    'This duplicates the original report-creation behavior.
    DetailSheet.Rows(1).Delete

    '--------------------------------------------------------
    ' Update the Trended sheet
    '--------------------------------------------------------
    UpdateStaticTrendedSheet _
        DetailSheet:=DetailSheet, _
        TrendSheet:=TrendSheet, _
        FacilityHeader:=FacilityHeader, _
        AccountHeader:=AccountHeader, _
        BalanceHeader:=BalanceHeader, _
        reportMonth:=reportMonth

    '--------------------------------------------------------
    ' Save to the current-month output path
    '--------------------------------------------------------
    reportFileFormat = FileFormatFromPath(OutputReportPath)

    Application.DisplayAlerts = False

    reportWorkbook.SaveAs _
        Filename:=OutputReportPath, _
        FileFormat:=reportFileFormat

    reportWorkbook.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:

    savedErrorNumber = Err.Number
    savedErrorDescription = Err.Description

    On Error Resume Next

    Application.CutCopyMode = False
    Application.DisplayAlerts = False

    If Not reportWorkbook Is Nothing Then
        reportWorkbook.Close SaveChanges:=False
    End If

    Application.DisplayAlerts = True
    On Error GoTo 0

    Err.Raise _
        savedErrorNumber, _
        "UpdateReportWithTrend", _
        savedErrorDescription

End Sub


Private Sub UpdateStaticTrendedSheet( _
    ByVal DetailSheet As Worksheet, _
    ByVal TrendSheet As Worksheet, _
    ByVal FacilityHeader As String, _
    ByVal AccountHeader As String, _
    ByVal BalanceHeader As String, _
    ByVal reportMonth As Variant)

    Const TREND_HEADER_ROW As Long = 3
    Const FIRST_DATA_ROW As Long = 4
    Const FACILITY_COLUMN As Long = 1
    Const MONTHS_TO_SHOW As Long = 6

    Dim headerRow As Long
    Dim facilityColumn As Long
    Dim accountColumn As Long
    Dim balanceColumn As Long

    Dim detailLastRow As Long
    Dim trendLastRow As Long
    Dim totalRow As Long

    Dim firstCurColumn As Long
    Dim secondCurColumn As Long
    Dim newARColumn As Long
    Dim newCountColumn As Long

    Dim rowNumber As Long
    Dim facilityName As String
    Dim accountNumber As String
    Dim accountBalance As Double

    Dim arTotals As Object
    Dim countTotals As Object

    Dim monthHeading As String

    Set arTotals = CreateObject("Scripting.Dictionary")
    Set countTotals = CreateObject("Scripting.Dictionary")

    arTotals.CompareMode = vbTextCompare
    countTotals.CompareMode = vbTextCompare

    '--------------------------------------------------------
    ' Find the source columns
    '--------------------------------------------------------
    headerRow = FindReportHeaderRow( _
                    DetailSheet, _
                    AccountHeader)

    If headerRow = 0 Then

        Err.Raise _
            vbObjectError + 2100, _
            "UpdateStaticTrendedSheet", _
            "The account header '" & AccountHeader & _
            "' was not found on " & DetailSheet.Name & "."

    End If

    facilityColumn = FindReportHeaderColumn( _
                            DetailSheet, _
                            headerRow, _
                            FacilityHeader)

    accountColumn = FindReportHeaderColumn( _
                            DetailSheet, _
                            headerRow, _
                            AccountHeader)

    balanceColumn = FindReportHeaderColumn( _
                            DetailSheet, _
                            headerRow, _
                            BalanceHeader)

    If facilityColumn = 0 Or _
       accountColumn = 0 Or _
       balanceColumn = 0 Then

        Err.Raise _
            vbObjectError + 2101, _
            "UpdateStaticTrendedSheet", _
            "One or more required columns were not found on " & _
            DetailSheet.Name & "." & vbCrLf & vbCrLf & _
            "Facility: " & FacilityHeader & vbCrLf & _
            "Account: " & AccountHeader & vbCrLf & _
            "Balance: " & BalanceHeader

    End If

    '--------------------------------------------------------
    ' Calculate current detail totals by facility
    '--------------------------------------------------------
    detailLastRow = LastUsedRowForReport(DetailSheet)

    For rowNumber = headerRow + 1 To detailLastRow

        accountNumber = CleanReportAccountNumber( _
                            DetailSheet.Cells( _
                                rowNumber, _
                                accountColumn))

        If Len(accountNumber) > 0 Then

            facilityName = Trim$(CStr( _
                DetailSheet.Cells( _
                    rowNumber, _
                    facilityColumn).Value2))

            If Len(facilityName) = 0 Then
                facilityName = "(Blank Facility)"
            End If

            accountBalance = ReportNumericValue( _
                DetailSheet.Cells( _
                    rowNumber, _
                    balanceColumn).Value2)

            If arTotals.Exists(facilityName) Then

                arTotals(facilityName) = _
                    CDbl(arTotals(facilityName)) + _
                    accountBalance

            Else

                arTotals.Add _
                    facilityName, _
                    accountBalance

            End If

            If countTotals.Exists(facilityName) Then

                countTotals(facilityName) = _
                    CLng(countTotals(facilityName)) + 1

            Else

                countTotals.Add _
                    facilityName, _
                    1&

            End If

        End If

    Next rowNumber

    '--------------------------------------------------------
    ' Find both Cur columns
    '--------------------------------------------------------
    FindTwoCurColumns _
        ws:=TrendSheet, _
        headerRow:=TREND_HEADER_ROW, _
        firstCurColumn:=firstCurColumn, _
        secondCurColumn:=secondCurColumn

    If firstCurColumn = 0 Or secondCurColumn = 0 Then

        Err.Raise _
            vbObjectError + 2102, _
            "UpdateStaticTrendedSheet", _
            "Two Cur columns were not found in row " & _
            TREND_HEADER_ROW & " of the Trended sheet."

    End If

    monthHeading = GetTrendMonthHeading(reportMonth)

    '--------------------------------------------------------
    ' Insert current AR month before first Cur
    '--------------------------------------------------------
    newARColumn = firstCurColumn

    TrendSheet.Columns(newARColumn).Insert _
        Shift:=xlToRight, _
        CopyOrigin:=xlFormatFromLeftOrAbove

    TrendSheet.Cells( _
        TREND_HEADER_ROW, _
        newARColumn).value = monthHeading

    'The Count section moved one column to the right.
    secondCurColumn = secondCurColumn + 1

    '--------------------------------------------------------
    ' Insert current Count month before second Cur
    '--------------------------------------------------------
    newCountColumn = secondCurColumn

    TrendSheet.Columns(newCountColumn).Insert _
        Shift:=xlToRight, _
        CopyOrigin:=xlFormatFromLeftOrAbove

    TrendSheet.Cells( _
        TREND_HEADER_ROW, _
        newCountColumn).value = monthHeading

    '--------------------------------------------------------
    ' Find ending rows
    '--------------------------------------------------------
    trendLastRow = TrendSheet.Cells( _
                        TrendSheet.Rows.Count, _
                        FACILITY_COLUMN).End(xlUp).Row

    If LCase$(Trim$(CStr( _
        TrendSheet.Cells( _
            trendLastRow, _
            FACILITY_COLUMN).value))) = "totals" Or _
       LCase$(Trim$(CStr( _
        TrendSheet.Cells( _
            trendLastRow, _
            FACILITY_COLUMN).value))) = "total" Then

        totalRow = trendLastRow

    Else

        totalRow = trendLastRow + 1
        TrendSheet.Cells(totalRow, FACILITY_COLUMN).value = "Totals"

    End If

    '--------------------------------------------------------
    ' Write current values into both new month and Cur
    '--------------------------------------------------------
    For rowNumber = FIRST_DATA_ROW To totalRow - 1

        facilityName = Trim$(CStr( _
            TrendSheet.Cells( _
                rowNumber, _
                FACILITY_COLUMN).Value2))

        If arTotals.Exists(facilityName) Then

            TrendSheet.Cells( _
                rowNumber, _
                newARColumn).value = _
                CDbl(arTotals(facilityName))

            TrendSheet.Cells( _
                rowNumber, _
                newARColumn + 1).value = _
                CDbl(arTotals(facilityName))

        Else

            TrendSheet.Cells( _
                rowNumber, _
                newARColumn).value = 0

            TrendSheet.Cells( _
                rowNumber, _
                newARColumn + 1).value = 0

        End If

        If countTotals.Exists(facilityName) Then

            TrendSheet.Cells( _
                rowNumber, _
                newCountColumn).value = _
                CLng(countTotals(facilityName))

            TrendSheet.Cells( _
                rowNumber, _
                newCountColumn + 1).value = _
                CLng(countTotals(facilityName))

        Else

            TrendSheet.Cells( _
                rowNumber, _
                newCountColumn).value = 0

            TrendSheet.Cells( _
                rowNumber, _
                newCountColumn + 1).value = 0

        End If

    Next rowNumber

    '--------------------------------------------------------
    ' Total formulas
    '--------------------------------------------------------
    TrendSheet.Cells( _
        totalRow, _
        newARColumn).Formula = _
        "=SUM(" & _
        TrendSheet.Range( _
            TrendSheet.Cells(FIRST_DATA_ROW, newARColumn), _
            TrendSheet.Cells(totalRow - 1, newARColumn) _
        ).Address(False, False) & ")"

    TrendSheet.Cells( _
        totalRow, _
        newARColumn + 1).Formula = _
        "=SUM(" & _
        TrendSheet.Range( _
            TrendSheet.Cells(FIRST_DATA_ROW, newARColumn + 1), _
            TrendSheet.Cells(totalRow - 1, newARColumn + 1) _
        ).Address(False, False) & ")"

    TrendSheet.Cells( _
        totalRow, _
        newCountColumn).Formula = _
        "=SUM(" & _
        TrendSheet.Range( _
            TrendSheet.Cells(FIRST_DATA_ROW, newCountColumn), _
            TrendSheet.Cells(totalRow - 1, newCountColumn) _
        ).Address(False, False) & ")"

    TrendSheet.Cells( _
        totalRow, _
        newCountColumn + 1).Formula = _
        "=SUM(" & _
        TrendSheet.Range( _
            TrendSheet.Cells(FIRST_DATA_ROW, newCountColumn + 1), _
            TrendSheet.Cells(totalRow - 1, newCountColumn + 1) _
        ).Address(False, False) & ")"

    '--------------------------------------------------------
    ' Formats
    '--------------------------------------------------------
    TrendSheet.Range( _
        TrendSheet.Cells(FIRST_DATA_ROW, newARColumn), _
        TrendSheet.Cells(totalRow, newARColumn + 1) _
    ).NumberFormat = "$#,##0;[Red]-$#,##0"

    TrendSheet.Range( _
        TrendSheet.Cells(FIRST_DATA_ROW, newCountColumn), _
        TrendSheet.Cells(totalRow, newCountColumn + 1) _
    ).NumberFormat = "#,##0"

    '--------------------------------------------------------
    ' Show only six months plus Cur in each section
    '--------------------------------------------------------
    CondenseTrendToRecentMonths _
        ws:=TrendSheet, _
        headerRow:=TREND_HEADER_ROW, _
        monthsToShow:=MONTHS_TO_SHOW

    TrendSheet.Activate
    TrendSheet.Range("A1").Select

End Sub


Private Sub CondenseTrendToRecentMonths( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal monthsToShow As Long)

    Dim firstCurColumn As Long
    Dim secondCurColumn As Long

    Dim firstARMonthColumn As Long
    Dim firstCountMonthColumn As Long
    Dim spacerColumn As Long
    Dim currentColumn As Long
    Dim lastColumn As Long

    FindTwoCurColumns _
        ws:=ws, _
        headerRow:=headerRow, _
        firstCurColumn:=firstCurColumn, _
        secondCurColumn:=secondCurColumn

    If firstCurColumn = 0 Or secondCurColumn = 0 Then

        Err.Raise _
            vbObjectError + 2200, _
            "CondenseTrendToRecentMonths", _
            "Two Cur columns were not found on " & ws.Name & "."

    End If

    lastColumn = ws.Cells( _
                    headerRow, _
                    ws.Columns.Count).End(xlToLeft).Column

    ws.Cells.EntireColumn.Hidden = False

    'A is the facility column.
    ws.Columns(1).Hidden = False

    'Keep six AR months immediately before AR Cur.
    firstARMonthColumn = _
        Application.Max(2, firstCurColumn - monthsToShow)

    For currentColumn = 2 To firstARMonthColumn - 1
        ws.Columns(currentColumn).Hidden = True
    Next currentColumn

    'The blank separator is immediately after first Cur.
    spacerColumn = firstCurColumn + 1
    ws.Columns(spacerColumn).Hidden = False

    'Keep six Count months immediately before Count Cur.
    firstCountMonthColumn = _
        Application.Max( _
            firstCurColumn + 2, _
            secondCurColumn - monthsToShow)

    For currentColumn = firstCurColumn + 2 _
                           To firstCountMonthColumn - 1

        ws.Columns(currentColumn).Hidden = True

    Next currentColumn

    'Hide anything after the second Cur.
    For currentColumn = secondCurColumn + 1 To lastColumn
        ws.Columns(currentColumn).Hidden = True
    Next currentColumn

End Sub


Private Sub FindTwoCurColumns( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByRef firstCurColumn As Long, _
    ByRef secondCurColumn As Long)

    Dim lastColumn As Long
    Dim currentColumn As Long
    Dim headerValue As String

    firstCurColumn = 0
    secondCurColumn = 0

    lastColumn = ws.Cells( _
                    headerRow, _
                    ws.Columns.Count).End(xlToLeft).Column

    For currentColumn = 1 To lastColumn

        headerValue = UCase$(Trim$(CStr( _
            ws.Cells( _
                headerRow, _
                currentColumn).Value2)))

        If headerValue = "CUR" Then

            If firstCurColumn = 0 Then

                firstCurColumn = currentColumn

            Else

                secondCurColumn = currentColumn
                Exit For

            End If

        End If

    Next currentColumn

End Sub


Private Function GetTrendMonthHeading( _
    ByVal reportMonth As Variant) As String

    Dim reportDate As Date
    Dim textValue As String

    If IsDate(reportMonth) Then

        reportDate = CDate(reportMonth)

    Else

        textValue = Trim$(CStr(reportMonth))

        If IsDate("1 " & textValue) Then

            reportDate = CDate("1 " & textValue)

        Else

            Err.Raise _
                vbObjectError + 2300, _
                "GetTrendMonthHeading", _
                "The reporting month could not be interpreted: " & _
                textValue

        End If

    End If

    GetTrendMonthHeading = _
        Format$(reportDate, "yy-mmm")

End Function


Private Function FindReportHeaderRow( _
    ByVal ws As Worksheet, _
    ByVal requiredHeader As String) As Long

    Dim rowNumber As Long

    For rowNumber = 1 To 25

        If FindReportHeaderColumn( _
            ws, _
            rowNumber, _
            requiredHeader) > 0 Then

            FindReportHeaderRow = rowNumber
            Exit Function

        End If

    Next rowNumber

End Function


Private Function FindReportHeaderColumn( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal requiredHeader As String) As Long

    Dim lastColumn As Long
    Dim currentColumn As Long

    Dim normalizedRequiredHeader As String
    Dim normalizedCellHeader As String

    normalizedRequiredHeader = _
        NormalizeReportHeader(requiredHeader)

    lastColumn = ws.Cells( _
                    headerRow, _
                    ws.Columns.Count).End(xlToLeft).Column

    For currentColumn = 1 To lastColumn

        normalizedCellHeader = NormalizeReportHeader( _
            ws.Cells( _
                headerRow, _
                currentColumn).Value2)

        If normalizedCellHeader = _
           normalizedRequiredHeader Then

            FindReportHeaderColumn = currentColumn
            Exit Function

        End If

    Next currentColumn

End Function


Private Function NormalizeReportHeader( _
    ByVal headerValue As Variant) As String

    Dim textValue As String
    Dim characterNumber As Long
    Dim currentCharacter As String
    Dim normalizedValue As String

    If IsError(headerValue) Or IsEmpty(headerValue) Then
        Exit Function
    End If

    textValue = LCase$(Trim$(CStr(headerValue)))

    For characterNumber = 1 To Len(textValue)

        currentCharacter = _
            Mid$(textValue, characterNumber, 1)

        If currentCharacter Like "[a-z0-9]" Then

            normalizedValue = _
                normalizedValue & currentCharacter

        End If

    Next characterNumber

    NormalizeReportHeader = normalizedValue

End Function


Private Function CleanReportAccountNumber( _
    ByVal sourceCell As Range) As String

    Dim cellValue As Variant
    Dim displayedValue As String

    cellValue = sourceCell.Value2

    If IsError(cellValue) Or IsEmpty(cellValue) Then
        Exit Function
    End If

    displayedValue = Trim$(sourceCell.Text)

    If Len(displayedValue) > 0 And _
       InStr(1, displayedValue, "###", vbTextCompare) = 0 And _
       InStr(1, displayedValue, "E+", vbTextCompare) = 0 Then

        CleanReportAccountNumber = displayedValue

    ElseIf IsNumeric(cellValue) Then

        CleanReportAccountNumber = _
            Format$(CDbl(cellValue), "0")

    Else

        CleanReportAccountNumber = _
            Trim$(CStr(cellValue))

    End If

End Function


Private Function ReportNumericValue( _
    ByVal value As Variant) As Double

    Dim textValue As String

    If IsError(value) Or IsEmpty(value) Then
        Exit Function
    End If

    If IsNumeric(value) Then

        ReportNumericValue = CDbl(value)
        Exit Function

    End If

    textValue = Trim$(CStr(value))

    textValue = Replace(textValue, "$", "")
    textValue = Replace(textValue, ",", "")
    textValue = Replace(textValue, " ", "")

    If Left$(textValue, 1) = "(" And _
       Right$(textValue, 1) = ")" Then

        textValue = _
            "-" & Mid$( _
                textValue, _
                2, _
                Len(textValue) - 2)

    End If

    If IsNumeric(textValue) Then
        ReportNumericValue = CDbl(textValue)
    End If

End Function


Private Function LastUsedRowForReport( _
    ByVal ws As Worksheet) As Long

    Dim lastCell As Range

    On Error Resume Next

    Set lastCell = ws.Cells.Find( _
                        What:="*", _
                        After:=ws.Cells(1, 1), _
                        LookIn:=xlFormulas, _
                        LookAt:=xlPart, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False)

    On Error GoTo 0

    If lastCell Is Nothing Then
        LastUsedRowForReport = 1
    Else
        LastUsedRowForReport = lastCell.Row
    End If

End Function


Private Function LastUsedColumnForReport( _
    ByVal ws As Worksheet) As Long

    Dim lastCell As Range

    On Error Resume Next

    Set lastCell = ws.Cells.Find( _
                        What:="*", _
                        After:=ws.Cells(1, 1), _
                        LookIn:=xlFormulas, _
                        LookAt:=xlPart, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False)

    On Error GoTo 0

    If lastCell Is Nothing Then
        LastUsedColumnForReport = 1
    Else
        LastUsedColumnForReport = lastCell.Column
    End If

End Function


Private Function FileFormatFromPath( _
    ByVal filePath As String) As XlFileFormat

    Dim extensionPosition As Long
    Dim extensionName As String

    extensionPosition = InStrRev(filePath, ".")

    If extensionPosition = 0 Then

        FileFormatFromPath = xlOpenXMLWorkbook
        Exit Function

    End If

    extensionName = _
        LCase$(Mid$(filePath, extensionPosition + 1))

    Select Case extensionName

        Case "xlsx"
            FileFormatFromPath = xlOpenXMLWorkbook

        Case "xlsm"
            FileFormatFromPath = xlOpenXMLWorkbookMacroEnabled

        Case "xlsb"
            FileFormatFromPath = xlExcel12

        Case "xls"
            FileFormatFromPath = xlExcel8

        Case Else
            FileFormatFromPath = xlOpenXMLWorkbook

    End Select

End Function

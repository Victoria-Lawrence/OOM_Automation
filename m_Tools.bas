Attribute VB_Name = "m_Tools"
Option Explicit

Function MyFindLastRow(wkTarget As Worksheet) As Long

    If WorksheetFunction.CountA(wkTarget.Cells) <> 0 Then
        'MyFindLastRow = wkTarget.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row | Original
        MyFindLastRow = wkTarget.Cells.Find(What:="*", After:=Range("A1"), LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Else
        MsgBox ("Worksheet is empty: " & wkTarget.Name)
        MyFindLastRow = 0
    End If


End Function
Function FindHeader(TargetHeader As String, wkTarget As Worksheet, EndonError As Boolean) As Range
    Dim Update As Variant
    ' Only Search the first 5 rows for the header
    ' Example
    ' Set returnRange = FindHeader("Users", wkTarget)
    ' returnRange.Row = row # of TargetHeader
    ' returnRange.Column = Column #
    ' Checks (1,1) last
    Application.FindFormat.Clear
    
    Dim rngFound As Range, rngSearch As Range, lastColTarget As Long
    wkTarget.Select
    wkTarget.Cells(1, 1).Select
    
    lastColTarget = MyFindLastColumn(wkTarget)

    
    Set rngSearch = Range(Cells(1, 1), Cells(5, lastColTarget))
    
    If WorksheetFunction.CountA(wkTarget.Cells) <> 0 Then
        Set FindHeader = rngSearch.Cells.Find(What:=TargetHeader, After:=ActiveCell, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
    Else
        MsgBox ("Worksheet is empty")
        FindHeader = Nothing
    End If
    
    If FindHeader Is Nothing Then  ' Header not found
        'Using like search
        If EndonError = True Then
            MsgBox ("Cannot find header " & UCase(TargetHeader) & " on the Worksheet " & UCase(wkTarget.Name) & vbNewLine & vbNewLine & "Correct the HEADER on the worksheet and try again")
            wkTarget.Cells(2, 1).Select
            End
        End If
    End If

'    wkTarget.Cells(FindHeader.Row, FindHeader.Column).Select

End Function
Function MyFindLastColumn(wkTarget As Worksheet) As Long

    wkTarget.Select
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False
    If WorksheetFunction.CountA(wkTarget.Cells) <> 0 Then
        MyFindLastColumn = wkTarget.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    Else
        MsgBox (wkTarget.Name & " is empty")
        MyFindLastColumn = 0
    End If

End Function
Function FindString(TargetString As String, StartRow As Long, EndRow As Long, searchCol As Integer, sheetName As String) As Long
    
        Dim rngSearch As Range, rngFound As Range
        Dim wkSearched As Worksheet
        
        Set wkSearched = Sheets(sheetName)
        Application.FindFormat.Clear
        
        wkSearched.Activate
        Set rngSearch = Range(Cells(StartRow, searchCol), Cells(EndRow, searchCol))
        Set rngFound = rngSearch.Find(What:=TargetString, LookIn:=xlValues, LookAt:=xlPart)

        If (rngFound Is Nothing) Then
            FindString = 0
            Exit Function
        Else
            If (rngFound.Row = StartRow) Then
                FindString = 0
                Exit Function
            Else
                FindString = rngFound.Row
            End If
        End If

    End Function
Function FindMIDASHeader(TargetHeader As String, wkTarget As Worksheet, EndonError As Boolean) As Range
    Dim Update As Variant
    ' Only Search the first 5 rows for the header
    ' Example
    ' Set returnRange = FindHeader("Users", wkTarget)
    ' returnRange.Row = row # of TargetHeader
    ' returnRange.Column = Column #
    ' Checks (1,1) last
    Application.FindFormat.Clear
    
    Dim rngFound As Range, rngSearch As Range, lastColTarget As Long
    wkTarget.Select
    wkTarget.Cells(1, 1).Select
    
    lastColTarget = MyFindLastColumn(wkTarget)

    
    Set rngSearch = Range(Cells(23, 8), Cells(23, 9))
    
    If WorksheetFunction.CountA(wkTarget.Cells) <> 0 Then
        Set FindHeader = rngSearch.Cells.Find(What:=TargetHeader, After:=ActiveCell, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
    Else
        MsgBox ("Worksheet is empty")
        FindHeader = Nothing
    End If
    
    If FindHeader Is Nothing Then  ' Header not found
        'Using like search
        If EndonError = True Then
            MsgBox ("Cannot find header " & UCase(TargetHeader) & " on the Worksheet " & UCase(wkTarget.Name) & vbNewLine & vbNewLine & "Correct the HEADER on the worksheet and try again")
            wkTarget.Cells(2, 1).Select
            End
        End If
    End If

'    wkTarget.Cells(FindHeader.Row, FindHeader.Column).Select

End Function

Sub DeleteAllExcept(ParamArray SheetsToKeep() As Variant)
    Dim ws As Worksheet
    Dim keepSheet As Boolean
    Dim sheetName As Variant

    Application.DisplayAlerts = False ' Prevent confirmation prompts

    For Each ws In ThisWorkbook.Worksheets
        keepSheet = False
        For Each sheetName In SheetsToKeep
            If ws.Name = sheetName Then
                keepSheet = True
                Exit For
            End If
        Next sheetName

        If Not keepSheet Then
            ws.Delete
        End If
    Next ws

    Application.DisplayAlerts = True ' Re-enable alerts
End Sub


Function PullPath(wsName As String, searchColumn As String, searchText As String) As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Set the worksheet dynamically
    Set ws = ThisWorkbook.Sheets(wsName)
    
    ' Find last row in the specified column
    lastRow = ws.Cells(ws.Rows.Count, searchColumn).End(xlUp).Row
    
    ' Loop through the specified column
    For i = 1 To lastRow
        If ws.Cells(i, searchColumn).Value = searchText Then
            ' Return value from the column immediately to the right
            PullPath = ws.Cells(i, ws.Columns(searchColumn).Column + 1).Value
            Exit Function
        End If
    Next i
    
    ' If not found, return message
    PullPath = "Not Found"
End Function


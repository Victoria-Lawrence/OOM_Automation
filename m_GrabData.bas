Attribute VB_Name = "m_GrabData"
Sub Grab_Raw_OOM_File_UPDATED() ' This is updated to pick the correct file automatically.

'12.18.25 - VL: Checks to ensure that the OOM Detail SQL was saved down.
    Dim answer As VbMsgBoxResult
    
    answer = MsgBox("Have you saved down the OOM Detail Data in the OOM SQL Extract?", _
                    vbYesNo + vbQuestion, "Confirm Step")

    If answer = vbNo Then
        MsgBox "Please complete that step before running this macro.", vbExclamation, "Action Needed"
        Exit Sub   ' <-- stops the code
    End If
'12.18.25 End Update

 Dim Currentbook As Workbook
    Dim Newbook As Workbook
    Dim l As Long

'11.5.25 - VL - Created to ensure that if this is saved while in the wrong stage it starts off clean so that we don't have an issue of combined data from previous months
 Dim sheetsToDelete As Variant: Dim sheetsToClear As Variant
    'Delete the Following Sheets if they exst
    sheetsToDelete = Array("FLAGGED Raw", "LARGE BALANCE Raw", "LESSER OF Raw", "STOPLOSS Raw", "ZERO CONTRACTUALS Raw")
    DeleteSheetsIfExist sheetsToDelete
    'Clear info in these sheets if there are any
    sheetsToClear = Array("LARGE BALANCE", "FLAGGED", "LESSER OF", "STOPLOSS", "ZERO CONTRACTUALS") ' Replace with your sheet names
    ClearDataBelowFirstRow sheetsToClear
    sheetsToClear = Array("OOM Detail")
    ClearDataBelowFirstRowOOMDetail sheetsToClear
    
    Set Currentbook = ActiveWorkbook
    
    Dim wkDetail As Worksheet
    Set wkDetail = Currentbook.Sheets("OOM Detail")
        If wkDetail.AutoFilterMode Then
        wkDetail.ShowAllData
    End If
    wkDetail.Rows("2:" & wkDetail.Rows.Count).Delete
    
    Set Newbook = Workbooks.Open("\\isilon-data\data\DADS\RCA Project HB\OOM\OOM SQL Extract.xlsx")
    
    l = Range("A1000000").End(xlUp).Row
    
    Range("A2:BA" & l).Copy
    
    Currentbook.Activate
    wkDetail.Select
    
    Range("A2").PasteSpecial (xlPasteAll)
    '12.3.25 - VRL - Bring over the new OON Column
        'Set Newbook = Workbooks.Open("\\isilon-data\data\DADS\RCA Project HB\OOM\OOM SQL Extract.xlsx")
    Newbook.Activate
    l = Range("A1000000").End(xlUp).Row
        Range("BB2:BD" & l).Copy
    
    Currentbook.Activate
    wkDetail.Select
    Range("BC2:BE" & l).PasteSpecial (xlPasteAll)
    'End 12.3.25 Update
    

    Application.CutCopyMode = False
    
    Newbook.Close SaveChanges:=False
    
    Sheets("Home").Select
    
    MsgBox ("You may proceed to Step 2")

End Sub

Sub DeleteandClearSheets()
'11.5.25 - VL - Created to ensure that if this is saved while in the wrong stage it starts off clean so that we don't have an issue of combined data from previous months
 Dim sheetsToDelete As Variant: Dim sheetsToClear As Variant
    'Delete the Following Sheets if they exst
    sheetsToDelete = Array("FLAGGED Raw", "LARGE BALANCE Raw", "LESSER OF Raw", "STOPLOSS Raw")
    DeleteSheetsIfExist sheetsToDelete
    'Clear info in these sheets if there are any
    sheetsToClear = Array("LARGE BALANCE", "FLAGGED", "LESSER OF", "STOPLOSS", "OOM Detail") ' Replace with your sheet names
    ClearDataBelowFirstRow sheetsToClear

End Sub

Sub DeleteSheetsIfExist(sheetNames As Variant)
    Dim sheetName As Variant
    Dim ws As Worksheet
    Application.DisplayAlerts = False ' Prevent confirmation prompts

    For Each sheetName In sheetNames
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetName)
        If Not ws Is Nothing Then
            ws.Delete
        End If
        Set ws = Nothing
        On Error GoTo 0
    Next sheetName

    Application.DisplayAlerts = True ' Re-enable alerts
End Sub

Sub ClearDataBelowFirstRow(sheetNames As Variant)
    Dim sheetName As Variant
    Dim ws As Worksheet
    Application.ScreenUpdating = False

    For Each sheetName In sheetNames
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetName)
        If Not ws Is Nothing Then
            With ws
                .Range("3:" & .Rows.Count).ClearContents
            End With
        End If
        Set ws = Nothing
        On Error GoTo 0
    Next sheetName

    Application.ScreenUpdating = True
End Sub

Sub ClearDataBelowFirstRowOOMDetail(sheetNames As Variant)
    Dim sheetName As Variant
    Dim ws As Worksheet
    Application.ScreenUpdating = False

    For Each sheetName In sheetNames
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetName)
        If Not ws Is Nothing Then
            With ws
                .Range("2:" & .Rows.Count).ClearContents
            End With
        End If
        Set ws = Nothing
        On Error GoTo 0
    Next sheetName

    Application.ScreenUpdating = True
End Sub

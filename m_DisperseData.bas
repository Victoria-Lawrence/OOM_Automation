Attribute VB_Name = "m_DisperseData"
Sub Disperse_Data()
    
    Dim l As Long
    Dim MySheet As Worksheet, wkOOM As Worksheet
    Dim NumCols As Long: Dim CurrentCol As Long: Dim DataSetLength As Long: Dim MyCol As Long
    Dim MyCell As Range
    
    If Sheets("Home").Range("E1").Value = "" Or Sheets("Home").Range("E2").Value = "" Then
    
        MsgBox ("You haven't populated your inputs into cells E1 and E2. Please do so and click again")
    
    End If
    
    Set wkOOM = Sheets("OOM Detail")
    wkOOM.Select
    l = Range("A1000000").End(xlUp).Row
    Range("BB1").Value = "Total Reserves"
    Range("BB2:BB" & l).Formula = "=U2+X2+AA2+AD2"
    Range("BB2:BB" & l).Cells.Style = "Currency"
    
    'Looping through the 4 worksheets and moving data over from the OOM Detail worksheet based on its mapping
    For Each MySheet In Sheets(Array("LARGE BALANCE", "FLAGGED", "LESSER OF", "STOPLOSS", "ZERO CONTRACTUALS"))
        
        wkOOM.Select
        Range("A1:BE" & l).AutoFilter Field:=53, Criteria1:=MySheet.Name
        Range("A1:BE" & l).SpecialCells(xlCellTypeVisible).Copy
        Worksheets.Add
        ActiveSheet.Name = MySheet.Name & " Raw"
        Range("A1").PasteSpecial (xlPasteAll)
        Application.CutCopyMode = False
        
        DataSetLength = Range("A1000000").End(xlUp).Row
        
        MySheet.Select
        NumCols = Range(Range("A2"), Range("A2").End(xlToRight)).Columns.Count
        
        For CurrentCol = 1 To NumCols
            
            If Cells(1, CurrentCol) <> "[leave blank]" Then
                
                MyCol = Cells(1, CurrentCol).Value
                Range(Cells(3, CurrentCol), Cells(DataSetLength + 1, CurrentCol)).Value = Range(Sheets(MySheet.Name & " Raw").Cells(2, MyCol), Sheets(MySheet.Name & " Raw").Cells(DataSetLength, MyCol)).Value
                Range(Cells(3, CurrentCol), Cells(DataSetLength + 1, CurrentCol)).NumberFormat = Range(Sheets(MySheet.Name & " Raw").Cells(2, MyCol), Sheets(MySheet.Name & " Raw").Cells(DataSetLength, MyCol)).NumberFormat
                
            End If
            
        Next
        
        If MySheet.Name = "LARGE BALANCE" Then
            
            'Removing "Epic-" from AcctNumber
            For Each MyCell In Range("A3:A" & DataSetLength + 1)
                
                MyCell.Value = Right(MyCell.Value, Len(MyCell.Value) - 5)
                MyCell.NumberFormat = "0"
            Next MyCell
            
            'Flipping the sign for RCA Est Contractual
            For Each MyCell In Range("P3:P" & DataSetLength + 1)
                
                MyCell.Value = -MyCell.Value
                
            Next MyCell
            ActiveSheet.Columns("BC").Delete ' - Deletes the Lesser Of OON Column
        ElseIf MySheet.Name = "FLAGGED" Then
        
            'Removing "Epic-" from AcctNumber
            For Each MyCell In Range("B3:B" & DataSetLength + 1)
                
                MyCell.Value = Right(MyCell.Value, Len(MyCell.Value) - 5)
                 MyCell.NumberFormat = "0"
            Next MyCell
            
            
            '% Expression
            Range("R3:R" & DataSetLength + 1).Formula = "=Q3/N3"
            Range("R3:R" & DataSetLength + 1).Cells.Style = "Percent"
        
        ElseIf MySheet.Name = "LESSER OF" Then
        
            'Removing "Epic-" from AcctNumber
            For Each MyCell In Range("B3:B" & DataSetLength + 1)
                
                MyCell.Value = Right(MyCell.Value, Len(MyCell.Value) - 5)
                MyCell.NumberFormat = "0"
            Next MyCell
            
            ElseIf MySheet.Name = "ZERO CONTRACTUALS" Then
        
            'Removing "Epic-" from AcctNumber
            For Each MyCell In Range("B3:B" & DataSetLength + 1)
                
                MyCell.Value = Right(MyCell.Value, Len(MyCell.Value) - 5)
                MyCell.NumberFormat = "0"
            Next MyCell
        
        ElseIf MySheet.Name = "STOPLOSS" Then
        
            'Removing "Epic-" from AcctNumber
            For Each MyCell In Range("F3:F" & DataSetLength + 1)
                
                MyCell.Value = Right(MyCell.Value, Len(MyCell.Value) - 5)
                MyCell.NumberFormat = "0"
            Next MyCell
            
            'STOPLOSS CONTRACTUAL
            Range("AA3:AA" & DataSetLength + 1).Formula = "=-((1-Y3)*V3)"
        
            'RCA CONTRACTUAL
            Range("AB3:AB" & DataSetLength + 1).Formula = "=T3"
        
            'DIFF
            Range("AC3:AC" & DataSetLength + 1).Formula = "=IF(S3=0,AA3-AB3,AA3-AB3+S3)"
        
        End If
        
        Columns.AutoFit
        Range("A1").Select
        
    Next MySheet
    
    Sheets("Home").Select
    
    MsgBox ("You may proceed to the next step")
    
End Sub

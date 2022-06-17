Sub RollForward()

    'Subroutine rolls the workbook forward each month to include current month data
    'the password for this workbook is: !R!#xd894k
    
    Workbooks("2022 Family Budget").Activate
    Dim LastRow As Integer, Current As Integer, LastDate As Integer
    Worksheets("Assets").Visible = True
    Worksheets("Retirement Accounts").Visible = True
    Worksheets("Liabilities").Visible = True
    Worksheets("Income").Visible = True
    Worksheets("Budget").Unprotect "!R!#xd894k"
    Worksheets("Monthly Cash Flow").Unprotect "!R!#xd894k"
    Worksheets("Retirement Accounts").Unprotect "!R!#xd894k"
    Worksheets("Financial Statements").Unprotect "!R!#xd894k"
    Worksheets("Liabilities").Unprotect "!R!#xd894k"
    Worksheets("Liabilities").Activate
    Worksheets("Liabilities").Select
    If Date > Worksheets("Liabilities").Range("A6") Then
        Range("A5:F6").Copy
        Range("A5:F6").PasteSpecial (xlPasteValues)
        Range("A5:F5").Delete (xlShiftUp)
    End If
    If Date > Worksheets("Liabilities").Range("H6") Then
        Range("H5:L6").Copy
        Range("H5:L6").PasteSpecial (xlPasteValues)
        Range("H5:L5").Delete (xlShiftUp)
        Range("L2").FormulaR1C1 = "=R5C[0]"
    End If
    If Date > Worksheets("Liabilities").Range("N6") Then
        Range("N5:S6").Copy
        Range("N5:S6").PasteSpecial (xlPasteValues)
        Range("N5:S5").Delete (xlShiftUp)
        Range("S2").FormulaR1C1 = "=R5C[0]"
    End If
    If Date > Worksheets("Liabilities").Range("U6") Then
        Range("U5:Y6").Copy
        Range("U5:Y6").PasteSpecial (xlPasteValues)
        Range("U5:Y5").Delete (xlShiftUp)
        Range("Y2").FormulaR1C1 = "=R5C[0]"
    End If
    If Date > Worksheets("Liabilities").Range("AA6") Then
        Range("AA5:AF6").Copy
        Range("AA5:AF6").PasteSpecial (xlPasteValues)
        Range("AA5:AF5").Delete (xlShiftUp)
        Range("AE2").FormulaR1C1 = "=R5C[0]"
    End If
    If Date > Worksheets("Liabilities").Range("AG6") Then
        Range("AH5:AK6").Copy
        Range("AH5:AK6").PasteSpecial xlPasteValues
        Range("AH5:AK5").Delete (xlShiftUp)
    End If
    Range("A3").ClearContents
    Worksheets("Financial Statements").Range("C15").FormulaR1C1 = "=SUM(Liabilities!R2C12,Liabilities!R5C19,Liabilities!R5C25,Liabilities!R5C31,Liabilities!R5C37,Liabilities!R5C47)"
    Worksheets("Liabilities").Protect "!R!#xd894k"
    Range("A5").Select
    
'Roll forward my retirement amounts each month.
    Worksheets("Retirement Accounts").Unprotect
    If Date >= Worksheets("Retirement Accounts").Range("D3") Then
        Worksheets("Retirement Accounts").Activate
        Worksheets("Retirement Accounts").Select
        Range("D3:G3").Copy
        Range("D3:G3").PasteSpecial xlPasteValues
        Range("D2:G2").Delete (xlShiftUp)
        Range("N9").Select
    End If
    Worksheets("Retirement Accounts").Range("M8").FormulaR1C1 = "=R2C7"
    Worksheets("Budget").Protect "!R!#xd894k"
    Worksheets("Monthly Cash Flow").Protect "!R!#xd894k"
    Worksheets("Retirement Accounts").Protect "!R!#xd894k"
    Worksheets("Financial Statements").Protect "!R!#xd894k"
    Worksheets("Financial Statements").Select
    Worksheets("Lists").Visible = xlSheetVeryHidden
    Worksheets("Assets").Visible = xlSheetVeryHidden
    Worksheets("Retirement Accounts").Visible = xlSheetVeryHidden
    Worksheets("Liabilities").Visible = xlSheetVeryHidden
    Worksheets("Income").Visible = xlSheetVeryHidden
    Worksheets("Financial Statements").Activate
    Range("C5").Select

        
End Sub

Sub Assets()

    Workbooks("2022 Family Budget").Activate
    Dim LastRow As Integer, Update As Integer
    LastRow = Range("B1").CurrentRegion.Rows.Count
    Update = 3
    Do While Update <= LastRow
        If Range("I" & Update).Value <= 0 Then
            Range("I" & Update).Value = 0
        End If
        Update = Update + 1
    Loop
    
End Sub

Sub add_asset()

    Workbooks("2022 Family Budget").Activate
    If Range("M3") = "" Then
        MsgBox "Please Name our new asset"
        Exit Sub
    End If
    If Range("M4") = "" Then
        MsgBox "Please select Balance Sheet category"
        Exit Sub
    End If
    If Range("M5") = "" Then
        MsgBox "Please select Asset category"
        Exit Sub
    End If
    If Range("M6") = "" Then
        MsgBox "Please select Asset sub-category"
        Exit Sub
    End If
    If Range("M8") = "" Then
        MsgBox "Please enter Asset original cost"
        Exit Sub
    End If
    Worksheets("Assets").Unprotect
    Dim LastRow As Integer
    LastRow = Range("B1").CurrentRegion.Rows.Count
    LastRow = LastRow + 1
    Range("B" & LastRow).Value = Range("M3")
    Range("C" & LastRow).Value = Range("M4")
    Range("D" & LastRow).Value = Range("M5")
    Range("E" & LastRow).Value = Range("M6")
    If Range("M7").Value <> "" Then
        Range("F" & LastRow).Value = Range("M7")
    End If
    If Range("M7").Value = "" Then
        Range("F" & LastRow).Value = Date
    End If
    Range("G" & LastRow).Value = Range("M8")
    If Range("M9").Value <> "" Then
        Range("H" & LastRow).Value = Range("M9") * 12
        Range("I3").Copy
        Range("I" & LastRow).PasteSpecial xlPasteAll
    End If
    If Range("M9").Value = "" Then
        Range("I" & LastRow).FormulaR1C1 = "=RC[-2]"
    End If
    Worksheets("Assets").Protect
    Range("M3:M9").ClearContents
    Range("B" & LastRow).Select
    
End Sub

Sub Edit_WB()

'This only exists to edit the workbook
    Workbooks("2022 Family Budget").Activate
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Visible = xlSheetVisible
    Next
    Worksheets("Will meeting").Visible = False
    Worksheets("Lists").Visible = False
    Worksheets("Budget").Unprotect "!R!#xd894k"
    Worksheets("Monthly Cash Flow").Unprotect "!R!#xd894k"
    Worksheets("Retirement Accounts").Unprotect "!R!#xd894k"
    Worksheets("Financial Statements").Unprotect "!R!#xd894k"
    Worksheets("Liabilities").Unprotect "!R!#xd894k"

End Sub


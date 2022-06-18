Sub PP_Reconciliation_Report()
'
' MW AP PowerPlan Reconciliation Report Macro
' Used to format the MW AP PowerPlan Reconciliation Report to hide numerous columns that are not used
' Testing comment for update to github
for the recon.
'
    ActiveSheet.Name = "MW AP PowerPlan Recon Report"
    Columns("A:I").EntireColumn.AutoFit
    Columns("A:A").ColumnWidth = 9
    Range("A1:M6").ClearFormats
    On Error Resume Next
        ActiveSheet.ShowAllData
    On Error GoTo 0
    If Range("O8") = "" Then
        Range("J:J").Copy
        Range("P:P").Insert
        Columns("A:K").Group
        Columns("M").Hidden = True
    End If
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
'Create Recon Columns and JE Information
'
    Dim LastRow As Long
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "L").End(xlUp).Row
    If Range("O8") = "" Then
        Columns("P:P").Select
        Selection.Replace What:="SN#", Replacement:="SN", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Selection.Replace What:="SN ", Replacement:="SN", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Range("O8").Value = "Serial Number"
        Range("O9").FormulaR1C1 = "=IFERROR(TRIM(IFERROR(LEFT(RIGHT(RC[1],LEN(RC[1])-FIND(""SN"",RC[1])+1),FIND("" "",RIGHT(RC[1],LEN(RC[1])-FIND(""SN"",RC[1])+1))),RIGHT(RC[1],LEN(RC[1])-FIND(""SN"",RC[1])+1))),""AP did not provide SN"")"
        Range("O9").Select
        Selection.AutoFill Destination:=Range("$O$9:$O" & LastRow)
        Range("O:O").Copy
        Range("O:O").PasteSpecial xlPasteValues
        Dim SN_Val As Integer
        SN_Val = 9
        Do While SN_Val <= LastRow
            If Cells(SN_Val, 15) = "AP did not provide SN" Then
                Cells(SN_Val, 15).FormulaR1C1 = "=IFERROR(TRIM(IFERROR(LEFT(RIGHT(RC[1],LEN(RC[1])-FIND(""HL"",RC[1])+1),FIND("" "",RIGHT(RC[1],LEN(RC[1])-FIND(""HL"",RC[1])+1))),RIGHT(RC[1],LEN(RC[1])-FIND(""HL"",RC[1])+1))),""AP did not provide SN"")"
            End If
            SN_Val = SN_Val + 1
        Loop
        Range("O:O").Copy
        Range("O:O").PasteSpecial xlPasteValues
        Selection.Replace What:="SN", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Selection.Replace What:=",", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Columns("O:O").EntireColumn.AutoFit
    End If
'The accountant performing the reconciliation may have to individually research some items
'In order not to undo their work, ILR #'s in this field should not be re-researched by the program.
'Additionally, some ILRs not appearing in the PP Rec should be removed. These are noted.
'ILR's always start with: OE, RE, RC, EF, VH
'also shield "Remove from Import"
    Range("P8").Value = "Researched ILR"
    Dim True_ILR As Integer
    True_ILR = 9
    Do While True_ILR <= LastRow
        If Left(Cells(True_ILR, 16), 2) <> "OE" Then
            If Left(Cells(True_ILR, 16), 2) <> "RE" Then
                If Left(Cells(True_ILR, 16), 2) <> "RC" Then
                    If Left(Cells(True_ILR, 16), 2) <> "EF" Then
                        If Left(Cells(True_ILR, 16), 2) <> "VH" Then
                            Cells(True_ILR, 16).FormulaR1C1 = _
                                "=IFERROR(INDEX('ILR Report for AP'!C2:C2,MATCH(RC[-1],'ILR Report for AP'!C7:C7,0)),""SN not in AP Report"")"
                        End If
                    End If
                End If
            End If
        End If
        True_ILR = True_ILR + 1
    Loop
    Range("P:P").Copy
    Range("P:P").PasteSpecial xlPasteValues
'If the ILR could not be mined from the VLOOKUP, a second option to gather the ILR is attempted
    Dim ILR_fromAP As Integer
    ILR_fromAP = 9
    Do While ILR_fromAP <= LastRow
        If Cells(ILR_fromAP, 16) = "SN not in AP Report" Then
            If Cells(ILR_fromAP, 8) <> "" Then
            Cells(ILR_fromAP, 16).FormulaR1C1 = _
                "=C8"
            Cells(ILR_fromAP, 18).Value = "Grabbed ILR from this report"
            End If
        End If
        ILR_fromAP = ILR_fromAP + 1
    Loop
'Once we have gathered the ILR's through a variety of means:
'ILR Validation If, Then, and Else process to validate ILR's
'1st Do While and If series: If AP has not provided a SN
'If the SN was not in the PP Rec, it may be a new lease and needs to be added to the Rec
'If the SN resulted in an ILR that was retired, this is noted in column Q
'And finally - if the SN resulted in an ILR that is in-service AND in the PP Rec, then the ILR is valid
'Oh, and I also need to code this so that it iterates over the entire column of Researched ILR's
    Range("Q8").Value = "ILR Validation"
    Dim ILR_Val As Integer
    ILR_Val = 9
    Do While ILR_Val <= LastRow
        If Cells(ILR_Val, 16) = "SN not in AP Report" Then
            Cells(ILR_Val, 17).Value = "Research SN"
        Else: Cells(ILR_Val, 17).FormulaR1C1 = _
            "=IFERROR(VLOOKUP(RC[-1],'PP Payment Reconciliation'!C7:C7,1,0),""ILR not in PP Rec"")"
        End If
        ILR_Val = ILR_Val + 1
    Loop
    Range("Q:Q").Copy
    Range("Q:Q").PasteSpecial xlPasteValues
    Range("Q9").Select
    ILR_Val = 9
    Do While ILR_Val <= LastRow
        If Left(Cells(ILR_Val, 17), 2) = "EF" Then
            Cells(ILR_Val, 17).Value = "In Service"
        ElseIf Left(Cells(ILR_Val, 17), 2) = "OE" Then
            Cells(ILR_Val, 17).Value = "In Service"
        ElseIf Left(Cells(ILR_Val, 17), 2) = "RC" Then
            Cells(ILR_Val, 17).Value = "In Service"
        ElseIf Left(Cells(ILR_Val, 17), 2) = "RE" Then
            Cells(ILR_Val, 17).Value = "In Service"
        ElseIf Left(Cells(ILR_Val, 17), 2) = "VH" Then
            Cells(ILR_Val, 17).Value = "In Service"
        ElseIf Cells(ILR_Val, 17) = "ILR not in PP Rec" Then
            Cells(ILR_Val, 18).FormulaR1C1 = _
                "=VLOOKUP(RC[-2],'ILR Report for AP'!C2:C4,3,0)"
            If Cells(ILR_Val, 18).Value = "Retired" Then
                Cells(ILR_Val, 18).FormulaR1C1 = _
                    "=CONCATENATE(""Retired "",VLOOKUP(RC[-2],'ILR Report for AP'!C2:C10,9,0))"
                Cells(ILR_Val, 17).Value = "Remove from import"
            End If
        End If
        ILR_Val = ILR_Val + 1
    Loop
    Range("R8").Value = "Comments"
    Range("S8").Value = "PP Co."
    Range("S9").FormulaR1C1 = _
        "=0&LEFT(IFERROR(VLOOKUP(RC[-3],'ILR Report for AP'!C2:C8,7,0),""""),3)"
    Range("S9").Select
    Selection.AutoFill Destination:=Range("$S$9:$S" & LastRow)
    Range("T8").Value = "AP Co."
    Range("T9").FormulaR1C1 = _
        "=LEFT(C2,4)"
    Range("T9").Select
    Selection.AutoFill Destination:=Range("$T$9:$T" & LastRow)
    Dim Co_Val As Integer
    Co_Val = 9
    Do While Co_Val <= LastRow
        If Cells(Co_Val, 17) <> "Research SN" Then
            If Cells(Co_Val, 17) <> "ILR not in PP Rec" Then
                If Cells(Co_Val, 17) <> "Remove from import" Then
                    If Cells(Co_Val, 19) <> Cells(Co_Val, 20) Then
                        Cells(Co_Val, 18).Value = "PP and AP Companies do not match"
                    End If
                End If
            End If
        End If
        Co_Val = Co_Val + 1
    Loop
    Range("U8").Value = "JE Co."
    Dim JE_Co_Row As Integer
    JE_Co_Row = 9
    Do While JE_Co_Row <= LastRow
        If Cells(JE_Co_Row, 17) = "Remove from import" Then
            Cells(JE_Co_Row, 21).FormulaR1C1 = _
                "=RC[-2]"
        ElseIf Cells(JE_Co_Row, 18) = "PP and AP Companies do not match" Then
            Cells(JE_Co_Row, 21).FormulaR1C1 = _
                "=RC[-2]"
            Cells(JE_Co_Row, 22).Value = "'1590500"
            Cells(JE_Co_Row, 18).Value = "Reclass to PP Company"
        End If
        JE_Co_Row = JE_Co_Row + 1
    Loop
'If PP and AP matches, Grab PP Company for potential reclass of retired ILRs
    Range("V8").Value = "JE Account"
    Dim JE_Acct_Row As Integer
    JE_Acct_Row = 9
    Do While JE_Acct_Row <= LastRow
        If Cells(JE_Acct_Row, 17) = "Remove from import" Then
            If Left(Cells(JE_Acct_Row, 16), 2) = "OE" Then
                Cells(JE_Acct_Row, 22).Value = "'7500030"
            ElseIf Left(Cells(JE_Acct_Row, 16), 2) = "VH" Then
                Cells(JE_Acct_Row, 22).Value = "'7500015"
            ElseIf Left(Cells(JE_Acct_Row, 16), 2) = "RE" Then
                Cells(JE_Acct_Row, 22).Value = "'7500010"
            ElseIf Left(Cells(JE_Acct_Row, 16), 2) = "EF" Then
                Cells(JE_Acct_Row, 22).Value = "'7500035"
            End If
        ElseIf Cells(JE_Acct_Row, 17) = "ILR not in PP Rec" Then
            If Cells(JE_Acct_Row, 16) <> "SN not in AP Report" Then
                Cells(JE_Acct_Row, 18).FormulaR1C1 = _
                    "=CONCATENATE(""ILR "",VLOOKUP(RC[-2],'ILR Report for AP'!C2:C4,3,0))"
                Cells(JE_Acct_Row, 18).Copy
                Range("R" & JE_Acct_Row).PasteSpecial xlPasteValues
                If Cells(JE_Acct_Row, 18) = "ILR Retired" Then
                    If Left(Cells(JE_Acct_Row, 16), 2) = "OE" Then
                        Cells(JE_Acct_Row, 22).Value = "'7500030"
                    ElseIf Left(Cells(JE_Acct_Row, 16), 2) = "VH" Then
                        Cells(JE_Acct_Row, 22).Value = "'7500015"
                    ElseIf Left(Cells(JE_Acct_Row, 16), 2) = "RE" Then
                        Cells(JE_Acct_Row, 22).Value = "'7500010"
                    ElseIf Left(Cells(JE_Acct_Row, 16), 2) = "EF" Then
                        Cells(JE_Acct_Row, 22).Value = "'7500035"
                    End If
                End If
'Above strings of If/Then arguments should establish correct account for reclass based on ILR#
'Space holder for potentially hard coding in accounts for prepaid items
'                If Cells(JE_Acct_Row, 18) = "ILR In-Service" Then
'                    Cells(JE_Acct_Row,18)
            End If
        End If
        JE_Acct_Row = JE_Acct_Row + 1
    Loop
    Range("W8").Value = "JE Plant"
'This field will have to be researched manually
    Columns("O:W").EntireColumn.AutoFit
    If Range("Q7") = "" Then
        Range("A8:W" & LastRow).AutoFilter
    End If
    Range("P7").Value = "Updated"
    Range("Q7").Formula = "=NOW()"
    Range("Q7").Copy
    Range("Q7").PasteSpecial xlPasteValues
    Range("O8:W" & LastRow).ClearFormats
    Range("O8:W" & LastRow).Select
    With Selection
        .HorizontalAlignment = xlLeft
    End With
    With Selection.Font
        .Name = "Calibri"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
        With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    Range("O8:W8").Select
    With Selection
        .VerticalAlignment = xlCenter
    End With
    Range("O9:W" & LastRow).Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    Range("P7:Q7").Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("Q7").Select
    Selection.NumberFormat = "mmm dd h:mm AM/PM"
    Range("P:P").ColumnWidth = 15
    Range("R9").Select
End Sub
Sub DL_Master_Temp()
'This macro sets up the direct labor template so that the master sheet can be used to create the
'associated journal entries to allocate direct labor to projects and offset payroll contra accounts
'for the employees that submitted the hours. This should be executed AFTER all lines received from
'the field have been entered and properly entered.
'Step 1: Calculate DL costs based on rate and validate projects
'The costs are specific to each employee and this information needs to be preserved and hard coded
'so that it isn't lost in future calculations.
    Worksheets("Project Entry Input").Unprotect
    Worksheets("Input Raw").Activate
    Worksheets("Input Raw").Select
    Dim LastRow As Integer
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    If Range("M1").Value = "No" Then
        Dim DL_COST As Integer
        DL_COST = 6
        Do While DL_COST <= LastRow
            Range("K" & DL_COST).FormulaR1C1 = "=IFERROR(R[0]C[-1]*VLOOKUP(R[0]C[1],'Employee Listing'!C4:C9,5,0),0)"
            DL_COST = DL_COST + 1
        Loop
        Range("J4").FormulaR1C1 = "=SUBTOTAL(9,R6C:R" & LastRow & "C)"
        Range("K4").FormulaR1C1 = "=SUBTOTAL(9,R6C:R" & LastRow & "C)"
        Range("N6").FormulaR1C1 = "=IFERROR(R[0]C[-3]/R[0]C[-4],"""")"
        Range("A6:N6").Copy
        Range("A7:N" & LastRow).PasteSpecial xlPasteFormats
        Range("N6").Copy
        Range("N7:N" & LastRow).PasteSpecial xlPasteAll
    End If
    If Range("M1").Value = "Yes" Then
        Range("K6:K" & LastRow).Copy
        Range("K6:K" & LastRow).PasteSpecial xlPasteValues
    End If
'Step 2: Gather information for in service projects.
'A separate template is sent and hours validated for in-service assets (projects that have
'already been gen and interfaced). We still need the relevant information pertaining to
'these projects. The following performs necessary procedures to reference an updated
'WIP Status Details report in order to get the information.
'
'Important note: You don't have to do ANYTHING to the WIP report. Just copy it into the tab
'and paste it over the top of the previous report.
    If Range("M1").Value = "No" Then
        Dim WIP_Info As Integer
        'Company Name
        WIP_Info = 6
        Do While WIP_Info <= LastRow
            If Range("B" & WIP_Info) = "" Then
                Range("B" & WIP_Info).FormulaR1C1 = _
                    "=INDEX('WIP DETAILS - UPDATE'!C3:C3,MATCH(RC[2],'WIP DETAILS - UPDATE'!C12:C12,0))"
            End If
            WIP_Info = WIP_Info + 1
        Loop
        'Plant Code
        WIP_Info = 6
        Do While WIP_Info <= LastRow
            If Range("C" & WIP_Info) = "" Then
                Range("C" & WIP_Info).FormulaR1C1 = _
                    "=INDEX('WIP DETAILS - UPDATE'!C4:C4,MATCH(RC[1],'WIP DETAILS - UPDATE'!C12:C12,0))"
            End If
            WIP_Info = WIP_Info + 1
        Loop
        'Project Name
        WIP_Info = 6
        Do While WIP_Info <= LastRow
            If Range("E" & WIP_Info) = "" Then
                Range("E" & WIP_Info).FormulaR1C1 = _
                    "=VLOOKUP(RC[-1],'WIP DETAILS - UPDATE'!C12:C13,2,FALSE)"
            End If
            WIP_Info = WIP_Info + 1
        Loop
        'Project Manager
        WIP_Info = 6
        Do While WIP_Info <= LastRow
            If Range("F" & WIP_Info) = "" Then
                Range("F" & WIP_Info).FormulaR1C1 = _
                    "=VLOOKUP(RC[-2],'WIP DETAILS - UPDATE'!C12:C14,3,FALSE)"
            End If
            WIP_Info = WIP_Info + 1
        Loop
        'Task Description
        WIP_Info = 6
        Do While WIP_Info <= LastRow
            If Range("H" & WIP_Info) = "" Then
                Range("H" & WIP_Info).FormulaArray = _
                    "=INDEX('WIP DETAILS - UPDATE'!C21:C21,MATCH(1,('WIP DETAILS - UPDATE'!C12:C12=RC[-4])*('WIP DETAILS - UPDATE'!C20:C20=RC[-1]),0))"
            End If
            WIP_Info = WIP_Info + 1
        Loop
        Range("A6:F" & LastRow).Copy
        Range("A6:F" & LastRow).PasteSpecial xlPasteValues
        Range("H6:J" & LastRow).Copy
        Range("H6:J" & LastRow).PasteSpecial xlPasteValues
        Range("L4").Formula = "=CONCATENATE(""Updated: "",TEXT(NOW(),""M/DD/YY H:MM AM/PM""))"
        Range("L4").Copy
        Range("L4").PasteSpecial xlPasteValues
        Range("A4").ClearContents
        Range("A4").Select
    End If
    If Range("M1").Value = "Yes" Then
        Dim Val_PE As Integer
        Val_PE = 6
        Do While Val_PE <= LastRow
            On Error GoTo NO_TASK
            Range("M" & Val_PE).FormulaArray = _
                "=INDEX('WIP DETAILS - UPDATE'!C17,MATCH(1,('WIP DETAILS - UPDATE'!C12=RC[-9])*('WIP DETAILS - UPDATE'!C20=RC[-6]),0))"
            Val_PE = Val_PE + 1
        Loop
        Val_PE = 6
        Do While Val_PE <= LastRow
            On Error GoTo NO_TASK
            If Range("M" & Val_PE).Value <> "Approved" Then
                Range("M" & Val_PE).Value = "Check"
            ElseIf Range("M" & Val_PE).Value = "Approved" Then
                Range("M" & Val_PE).Value = "OK"
            End If
            Val_PE = Val_PE + 1
        Loop
    End If

'Step 3: After the input form has been completed, create the project entry to charge labor expense to
'projects. All expense tasks should be coded to 403E, all capital tasks should be coded to 403C.
    If Range("M1").Value = "Yes" Then
        Dim PE_Input As Integer, Input_Raw As Integer
        PE_Input = 13
        Input_Raw = 6
        Do While Input_Raw <= LastRow
            Worksheets("Project Entry Input").Range("E9").FormulaR1C1 = "=CONCATENATE(LEFT(R[0]C[6],1),RIGHT(R[0]C[6],LEN(R[0]C[6])-FIND("" "",R[0]C[6])),TEXT(EOMONTH(R[0]C[8],0)+2,""MMDDYY""),"" Direct Labor Allocation "",TEXT(R[0]C[8],""Mmm-YY""))"
            Worksheets("Project Entry Input").Range("C" & PE_Input).FormulaR1C1 = "='Input Raw'!R" & Input_Raw & "C4"
            Worksheets("Project Entry Input").Range("D" & PE_Input).FormulaR1C1 = "='Input Raw'!R" & Input_Raw & "C7"
            Worksheets("Project Entry Input").Range("E" & PE_Input).FormulaR1C1 = "=TEXT(INT((EOMONTH(R9C13,0)-1)/7)*7+1,""DD-MMM-YY"")"
            Worksheets("Project Entry Input").Range("F" & PE_Input).FormulaR1C1 = "=TEXT(INT((EOMONTH(R9C13,0)-1)/7)*7+1,""DD-MMM-YY"")"
            Worksheets("Project Entry Input").Range("G" & PE_Input).FormulaR1C1 = "=IF(LEFT(R[0]C[-3],2)<>""99"",""403C Direct Labor"",""403E Direct Labor Exp"")"
            Worksheets("Project Entry Input").Range("H" & PE_Input).FormulaR1C1 = "=""DL Allocation ""&TEXT(R9C13,""MMM-YY"")"
            Worksheets("Project Entry Input").Range("I" & PE_Input).FormulaR1C1 = "=ROUND('Input Raw'!R" & Input_Raw & "C11,2)"
            Worksheets("Project Entry Input").Range("J" & PE_Input).Value = "PJ"
            Worksheets("Project Entry Input").Range("K" & PE_Input).FormulaR1C1 = "='Input Raw'!R" & Input_Raw & "C3"
            Worksheets("Project Entry Input").Range("M" & PE_Input).FormulaR1C1 = "=CONCATENATE(TEXT(R9C13,""MMM-YY""),"" Direct Labor - "",'Input Raw'!R" & Input_Raw & "C12)"
            Worksheets("Project Entry Input").Range("O" & PE_Input).Value = "N"
            Input_Raw = Input_Raw + 1
            PE_Input = PE_Input + 1
        Loop
        Worksheets("Project Entry Input").Activate
        Range("B" & PE_Input & ":O999").Select
        Selection.EntireRow.Delete
        Dim LAST_PEROW As Integer
        LAST_PEROW = Worksheets("Project Entry Input").Cells(Rows.Count, 3).End(xlUp).Row
        Worksheets("Project Entry Input").Range("B13:O13").Copy
        Worksheets("Project Entry Input").Range("B13:O" & LAST_PEROW + 1).PasteSpecial xlPasteFormats
        Worksheets("Project Entry Input").Columns("K").AutoFit
        Worksheets("Project Entry Input").Columns("M").AutoFit
        Worksheets("Project Entry Input").Activate
        Range("C13").Select
    End If
'For the in service date I would love to program a check so that I don't end up with
'a bunch of unnecessary array formulas in the sheet since most projects in the template
'will not be in service and therefore column J will be blank...
'    'In Service Date
'    WIP_Info = 6
'    Do While WIP_Info <= LastRow
'        If Range("I" & WIP_Info) = "" Then
'            Range("I" & WIP_Info).FormulaArray = _
'                "=INDEX('WIP DETAILS - UPDATE'!C25:C25,MATCH(1,('WIP DETAILS - UPDATE'!C12:C12='Master Template'!RC[-5])*('WIP DETAILS - UPDATE'!C20:C20='Master Template'!RC[-2]),0))"
'        End If
'    WIP_Info = WIP_Info + 1
'    Loop
    Worksheets("Project Entry Input").Protect
    Exit Sub
NO_TASK:
    MsgBox "Missing Tasks must be provided"
    Exit Sub
End Sub
Sub create_GL()
'Step 4: This is the fun part. The project entry debits the project (CIP) and credits CIP Clearing. Another
'entry needs to be made to move the payroll expense accounts associated with each employee that submitted hours
'to offset the balance that was just created in CIP clearing. There is a hidden tab that performs these
'calculations.
    If Worksheets("Input Raw").Range("M1").Value = "Yes" Then
        Worksheets("GL Calculations").Visible = True
        Worksheets("GL Calculations").Unprotect
        Worksheets("GL Calculations").Activate
        Worksheets("GL Calculations").Select
        ThisWorkbook.RefreshAll
'***********************this pivot isn't updating to contain new information**************************
'        ActiveSheet.PivotTables("Contras").PivotFields("Task Number").CurrentPage = _
'            "(All)"
'        With ActiveSheet.PivotTables("Contras").PivotFields("Task Number")
'            .PivotItems("99").Visible = False
'            .PivotItems("(blank)").Visible = False
'        End With
        Dim LASTCONTRAS_ROW As Integer
        LASTCONTRAS_ROW = Cells(Rows.Count, 2).End(xlUp).Row
        LASTCONTRAS_ROW = LASTCONTRAS_ROW - 1
        Range("F4").FormulaR1C1 = "=0&LEFT(R[0]C[-3],3)&"".1699000.""&LEFT(R[0]C[-2],4)"
        Range("G4").FormulaR1C1 = "=VLOOKUP(R[0]C[-5],'Employee Listing'!C4:C5,2,FALSE)"
        Range("H4").FormulaR1C1 = "=IF(LEFT(R[0]C[-2],4)=LEFT(R[0]C[-1],4),"""",""Yes"")"
        Range("I4").FormulaR1C1 = "=IF(R[0]C[-1]=""Yes"",LEFT(R[0]C[-2],4)&""-""&LEFT(R[0]C[-3],4),"""")"
        Range("J4").FormulaR1C1 = "=IF(R[0]C[-2]=""Yes"",LEFT(R[0]C[-4],4)&""-""&LEFT(R[0]C[-3],4),"""")"
        Range("K4").FormulaR1C1 = "=R[0]C5"
        Range("F4:K4").Copy
        Range("F5:K" & LASTCONTRAS_ROW).PasteSpecial xlPasteAll
        Range("F" & LASTCONTRAS_ROW + 1 & ":K999").ClearContents
        Range("B3").Select
        ThisWorkbook.RefreshAll
'        With ActiveSheet.PivotTables("CIP").PivotFields("Project CIP Clearing")
'            .PivotItems("(blank)").Visible = False
'        End With
'        With ActiveSheet.PivotTables("Payroll").PivotFields("Payroll Contra Account")
'            .PivotItems("(blank)").Visible = False
'        End With
'        With ActiveSheet.PivotTables("InterCo_Rec").PivotFields("Interco Receivable")
'            .PivotItems("").Visible = False
'            .PivotItems("(blank)").Visible = False
'        End With
'        With ActiveSheet.PivotTables("InterCo_Pay").PivotFields("Interco Payable")
'            .PivotItems("").Visible = False
'            .PivotItems("(blank)").Visible = False
'        End With
'        ActiveSheet.PivotTables("Expense").PivotFields("Company Org Name").CurrentPage = _
'            "99"
'Step 4 (continued): Once the data has been created, now we need to put it all into the GL entry
'template to copy into a blank GL entry. Several variables are used and have to be reset to ensure
'that sections of the GL entry do not overwrite something.
'**********************************************************************************************
        Worksheets("GL Entry Input").Range("B14:M999").ClearContents
        Dim GL_ACCT_INPUT As Integer, GL_CALC_LAST As Integer, GL_ENTRY As Integer
        'Debit CIP Clearing to offset what the project entry did and to inherit expense from Payroll Contras
        GL_ACCT_INPUT = 4
        GL_CALC_LAST = Cells(Rows.Count, 13).End(xlUp).Row - 1
        GL_ENTRY = 14
        Do While GL_ACCT_INPUT <= GL_CALC_LAST
            Worksheets("GL Entry Input").Range("E9").FormulaR1C1 = "=CONCATENATE(LEFT(R[0]C[6],1),RIGHT(R[0]C[6],LEN(R[0]C[6])-FIND("" "",R[0]C[6])),TEXT(EOMONTH(R[0]C[8],0)+2,""MMDDYY""),"" Direct LaborAllocation GL "",TEXT(R[0]C[8],""Mmm-YY""))"
            Worksheets("GL Entry Input").Range("C" & GL_ENTRY).FormulaR1C1 = "=TEXT(LEFT('GL Calculations'!R" & GL_ACCT_INPUT & "C13,4),""0000"")"
            Worksheets("GL Entry Input").Range("D" & GL_ENTRY).Value = "=TEXT(""1699000"",""0000000"")"
            Worksheets("GL Entry Input").Range("E" & GL_ENTRY).FormulaR1C1 = "=TEXT(RIGHT('GL Calculations'!R" & GL_ACCT_INPUT & "C13,4),""0000"")"
            Worksheets("GL Entry Input").Range("F" & GL_ENTRY).Value = "=TEXT(""000"",""000"")"
            Worksheets("GL Entry Input").Range("G" & GL_ENTRY).Value = "=TEXT(""0000000"",""0000000"")"
            Worksheets("GL Entry Input").Range("H" & GL_ENTRY).Value = "=TEXT(""00000000"",""00000000"")"
            Worksheets("GL Entry Input").Range("I" & GL_ENTRY).FormulaR1C1 = "=ROUND('GL Calculations'!R" & GL_ACCT_INPUT & "C14,2)"
            If Worksheets("GL Entry Input").Range("I" & GL_ENTRY).Value < 0 Then
                Worksheets("GL Entry Input").Range("J" & GL_ENTRY).FormulaR1C1 = "=-ROUND('GL Calculations'!R" & GL_ACCT_INPUT & "C14,2)"
                Worksheets("GL Entry Input").Range("I" & GL_ENTRY).ClearContents
            End If
            Worksheets("GL Entry Input").Range("K" & GL_ENTRY).FormulaR1C1 = "=CONCATENATE(TEXT(R9C13,""Mmm-YY""),"" Direct Labor Allocation"")"
            GL_ACCT_INPUT = GL_ACCT_INPUT + 1
            GL_ENTRY = GL_ENTRY + 1
        Loop
        GL_ENTRY = GL_ENTRY + 1
        'Credit Payroll Contras to undo payroll expense so it can be moved to offset CIP Clearing
        GL_ACCT_INPUT = 4
        GL_CALC_LAST = Cells(Rows.Count, 15).End(xlUp).Row - 1
        Do While GL_ACCT_INPUT <= GL_CALC_LAST
            Worksheets("GL Entry Input").Range("C" & GL_ENTRY).FormulaR1C1 = "=TEXT(LEFT('GL Calculations'!R" & GL_ACCT_INPUT & "C15,4),""0000"")"
            Worksheets("GL Entry Input").Range("D" & GL_ENTRY).Value = "=TEXT(""7009000"",""0000000"")"
            Worksheets("GL Entry Input").Range("E" & GL_ENTRY).FormulaR1C1 = "=TEXT(LEFT(RIGHT('GL Calculations'!R" & GL_ACCT_INPUT & "C15,8),4),""0000"")"
            Worksheets("GL Entry Input").Range("F" & GL_ENTRY).Value = "=TEXT(RIGHT('GL Calculations'!R" & GL_ACCT_INPUT & "C15,3),""000"")"
            Worksheets("GL Entry Input").Range("G" & GL_ENTRY).Value = "=TEXT(""0000000"",""0000000"")"
            Worksheets("GL Entry Input").Range("H" & GL_ENTRY).Value = "=TEXT(""00000000"",""00000000"")"
            Worksheets("GL Entry Input").Range("J" & GL_ENTRY).FormulaR1C1 = "=ROUND('GL Calculations'!R" & GL_ACCT_INPUT & "C16,2)"
            If Worksheets("GL Entry Input").Range("J" & GL_ENTRY).Value < 0 Then
                Worksheets("GL Entry Input").Range("I" & GL_ENTRY).FormulaR1C1 = "=-ROUND('GL Calculations'!R" & GL_ACCT_INPUT & "C16,2)"
                Worksheets("GL Entry Input").Range("J" & GL_ENTRY).ClearContents
            End If
            Worksheets("GL Entry Input").Range("K" & GL_ENTRY).FormulaR1C1 = "=CONCATENATE(TEXT(R9C13,""Mmm-YY""),"" Direct Labor Allocation"")"
            GL_ACCT_INPUT = GL_ACCT_INPUT + 1
            GL_ENTRY = GL_ENTRY + 1
        Loop
        GL_ENTRY = GL_ENTRY + 1
        'If the project company and the payroll company is not the same, an intercompany transfer is necessary
        'to create a receivable for the payroll company (getting rid of expense) and a payable for the project
        'company (inheriting expense). This is also guided by if either of the companies are consolidated and/or
        'a joint venture entity. Debits are entered and then credits.
'        Function IsInArray(COMPANY As String, UNCON_JV As Variant) As Boolean
'            IsInArray = UBound(Filter(arr, stringToBeFound)) > -1
'        End Function
        Dim UNCON_JV() As Variant, CON_JVLIB() As Variant, WHOLOWN_CON() As Variant, COMPANY As String, INTER_CO As String, SEARCH_INTERCO As Variant
        UNCON_JV() = Array("0286", "0288", "0291", "0293", "0295", "0324", "0330", "0331", "0345")
        CON_JVLIB() = Array("0266", "0268", "0269", "0270", "0271", "0272", "0273", "0274", "0278", "0279", "0328", "0332")
        WHOLOWN_CON() = Array("0200", "0250", "0252", "0258", "0280", "0281", "0283", "0290", "0302","0303", "0304", "0305", "0307", "0309", "0310", "0311", "0312", "0316", "0319", "0320", "0325", "0334", "0346", "0355")
        GL_ACCT_INPUT = 4
        GL_CALC_LAST = Cells(Rows.Count, 17).End(xlUp).Row - 1
        Do While GL_ACCT_INPUT <= GL_CALC_LAST
            Worksheets("GL Entry Input").Range("C" & GL_ENTRY).FormulaR1C1 = "=TEXT(LEFT('GL Calculations'!R" & GL_ACCT_INPUT & "C17,4),""0000"")"
            Worksheets("GL Entry Input").Range("D" & GL_ENTRY).FormulaR1C1 = "=TEXT(RIGHT('GL Calculations'!R" & GL_ACCT_INPUT & "C17,4),""0000"")"
            INTER_CO = "1190"
            COMPANY = Worksheets("GL Entry Input").Range("C" & GL_ENTRY).Value
            SEARCH_INTERCO = Filter(UNCON_JV, COMPANY)
            If UBound(SEARCH_INTERCO) > -1 Then
                INTER_CO = "1940"
            ElseIf UBound(SEARCH_INTERCO) = -1 Then
                SEARCH_INTERCO = Filter(CON_JVLIB, COMPANY)
                If UBound(SEARCH_INTERCO) > -1 Then
                    COMPANY = Worksheets("GL Entry Input").Range("D" & GL_ENTRY).Value
                    SEARCH_INTERCO = Filter(UNCON_JV, COMPANY)
                    If UBound(SEARCH_INTERCO) > -1 Then
                        INTER_CO = "1940"
                    Else: INTER_CO = "1191"
                    End If
                Else
                    COMPANY = Worksheets("GL Entry Input").Range("D" & GL_ENTRY).Value
                    SEARCH_INTERCO = Filter(UNCON_JV, COMPANY)
                    If UBound(SEARCH_INTERCO) > -1 Then
                        INTER_CO = "1940"
                    Else
                        SEARCH_INTERCO = Filter(CON_JVLIB, COMPANY)
                        If UBound(SEARCH_INTERCO) > -1 Then
                            INTER_CO = "1191"
                        Else: INTER_CO = "1190"
                        End If
                    End If
                End If
            End If
            Worksheets("GL Entry Input").Range("D" & GL_ENTRY).FormulaR1C1 = "=TEXT(" & INTER_CO & ",""0000"")&TEXT(RIGHT('GL Calculations'!R" & GL_ACCT_INPUT & "C17,4),""000"")"
            Worksheets("GL Entry Input").Range("E" & GL_ENTRY).Value = "=TEXT(""1000"",""0000"")"
            Worksheets("GL Entry Input").Range("F" & GL_ENTRY).Value = "=TEXT(""000"",""000"")"
            Worksheets("GL Entry Input").Range("G" & GL_ENTRY).Value = "=TEXT(""0000000"",""0000000"")"
            Worksheets("GL Entry Input").Range("H" & GL_ENTRY).Value = "=TEXT(""00000000"",""00000000"")"
            Worksheets("GL Entry Input").Range("I" & GL_ENTRY).FormulaR1C1 = "=ROUND('GL Calculations'!R" & GL_ACCT_INPUT & "C18,2)"
            If Worksheets("GL Entry Input").Range("I" & GL_ENTRY).Value < 0 Then
                Worksheets("GL Entry Input").Range("J" & GL_ENTRY).FormulaR1C1 = "=-ROUND('GL Calculations'!R" & GL_ACCT_INPUT & "C18,2)"
                Worksheets("GL Entry Input").Range("I" & GL_ENTRY).ClearContents
            End If
            Worksheets("GL Entry Input").Range("K" & GL_ENTRY).FormulaR1C1 = "=CONCATENATE(TEXT(R9C13,""Mmm-YY""),"" Direct Labor Allocation"")"
            GL_ACCT_INPUT = GL_ACCT_INPUT + 1
            GL_ENTRY = GL_ENTRY + 1
        Loop
        GL_ENTRY = GL_ENTRY + 1
        GL_ACCT_INPUT = 4
        GL_CALC_LAST = Cells(Rows.Count, 19).End(xlUp).Row - 1
        Worksheets("GL Entry Input").Select
        Do While GL_ACCT_INPUT <= GL_CALC_LAST
            Worksheets("GL Entry Input").Range("C" & GL_ENTRY).FormulaR1C1 = "=TEXT(LEFT('GL Calculations'!R" & GL_ACCT_INPUT & "C19,4),""0000"")"
            Worksheets("GL Entry Input").Range("D" & GL_ENTRY).FormulaR1C1 = "=TEXT(RIGHT('GL Calculations'!R" & GL_ACCT_INPUT & "C19,4),""0000"")"
            INTER_CO = "1190"
            COMPANY = Worksheets("GL Entry Input").Range("C" & GL_ENTRY).Value
            SEARCH_INTERCO = Filter(UNCON_JV, COMPANY)
            If UBound(SEARCH_INTERCO) > -1 Then
                INTER_CO = "2940"
            ElseIf UBound(SEARCH_INTERCO) = -1 Then
                SEARCH_INTERCO = Filter(CON_JVLIB, COMPANY)
                If UBound(SEARCH_INTERCO) > -1 Then
                    COMPANY = Worksheets("GL Entry Input").Range("D" & GL_ENTRY).Value
                    SEARCH_INTERCO = Filter(UNCON_JV, COMPANY)
                    If UBound(SEARCH_INTERCO) > -1 Then
                        INTER_CO = "2940"
                    Else: INTER_CO = "1191"
                    End If
                Else
                    COMPANY = Worksheets("GL Entry Input").Range("D" & GL_ENTRY).Value
                    SEARCH_INTERCO = Filter(UNCON_JV, COMPANY)
                    If UBound(SEARCH_INTERCO) > -1 Then
                        INTER_CO = "2940"
                    Else
                        SEARCH_INTERCO = Filter(CON_JVLIB, COMPANY)
                        If UBound(SEARCH_INTERCO) > -1 Then
                            INTER_CO = "1191"
                        Else: INTER_CO = "1190"
                        End If
                    End If
                End If
            End If
            Worksheets("GL Entry Input").Range("D" & GL_ENTRY).FormulaR1C1 = "=TEXT(" & INTER_CO & ",""0000"")&TEXT(RIGHT('GL Calculations'!R" & GL_ACCT_INPUT & "C19,4),""000"")"
            Worksheets("GL Entry Input").Range("E" & GL_ENTRY).Value = "=TEXT(""1000"",""0000"")"
            Worksheets("GL Entry Input").Range("F" & GL_ENTRY).Value = "=TEXT(""000"",""000"")"
            Worksheets("GL Entry Input").Range("G" & GL_ENTRY).Value = "=TEXT(""0000000"",""0000000"")"
            Worksheets("GL Entry Input").Range("H" & GL_ENTRY).Value = "=TEXT(""00000000"",""00000000"")"
            Worksheets("GL Entry Input").Range("J" & GL_ENTRY).FormulaR1C1 = "=ROUND('GL Calculations'!R" & GL_ACCT_INPUT & "C20,2)"
            If Worksheets("GL Entry Input").Range("J" & GL_ENTRY).Value < 0 Then
                Worksheets("GL Entry Input").Range("I" & GL_ENTRY).FormulaR1C1 = "=-ROUND('GL Calculations'!R" & GL_ACCT_INPUT & "C20,2)"
                Worksheets("GL Entry Input").Range("J" & GL_ENTRY).ClearContents
            End If
            Worksheets("GL Entry Input").Range("K" & GL_ENTRY).FormulaR1C1 = "=CONCATENATE(TEXT(R9C13,""Mmm-YY""),"" Direct Labor Allocation"")"
            GL_ACCT_INPUT = GL_ACCT_INPUT + 1
            GL_ENTRY = GL_ENTRY + 1
        Loop
        GL_ENTRY = GL_ENTRY + 1
        Worksheets("GL Entry Input").Activate
        Range("B" & GL_ENTRY + 1 & ":M999").Select
        Selection.EntireRow.Delete
        Range("B14:M14").Copy
        Range("B15:M" & GL_ENTRY).PasteSpecial xlPasteFormats
        Range("H" & GL_ENTRY).Value = "TOTAL"
        Range("I" & GL_ENTRY).FormulaR1C1 = "=SUM(R13C[0]:R" & GL_ENTRY - 1 & "C[0])"
        Range("J" & GL_ENTRY).FormulaR1C1 = "=SUM(R13C[0]:R" & GL_ENTRY - 1 & "C[0])"
        Range("K" & GL_ENTRY).FormulaR1C1 = "=IF(R[0]C[-2]<>R[0]C[-1],""Invalid GL Entry"","""")"
        Range("H" & GL_ENTRY).Select
        With Selection
            .HorizontalAlignment = xlCenter
        End With
        Range("H" & GL_ENTRY & ":J" & GL_ENTRY).Select
        Selection.Font.Bold = True
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
        Worksheets("GL Entry Input").Select
        Range("K" & GL_ENTRY).Select
        Worksheets("GL Calculations").Protect
'        Worksheets("GL Calculations").Visible (xlSheetVeryHidden)
    End If

End Sub
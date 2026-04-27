' ============================================================
'  Equity Valuation Model — VBA Macros
'  File: ValuationMacros.bas
'  Description: Automation macros for the valuation workbook
' ============================================================

Option Explicit

' ── Constants ────────────────────────────────────────────────────────────────

Private Const ASSUMP_SHEET As String = "Assumptions"
Private Const DCF_SHEET    As String = "DCF"
Private Const SENS_SHEET   As String = "Sensitivity"
Private Const COMPS_SHEET  As String = "Comps"
Private Const INCOME_SHEET As String = "Income Projection"
Private Const REPORT_DIR   As String = "reports\"


' ============================================================
' SUB: RunValuation
' Triggers full model recalculation and highlights key outputs.
' ============================================================

Public Sub RunValuation()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(DCF_SHEET)

    ' Flash output cells to draw attention
    Dim outCells As Range
    Set outCells = ws.Range("B25:B27")
    
    Dim i As Integer
    For i = 1 To 3
        outCells.Interior.Color = RGB(240, 165, 0)
        Application.Wait Now + TimeSerial(0, 0, 0)
        DoEvents
        outCells.Interior.Color = RGB(245, 247, 250)
        DoEvents
    Next i

    Application.ScreenUpdating = True

    MsgBox "✅ Valuation Complete!" & vbNewLine & vbNewLine & _
           "Enterprise Value:  $" & Format(ws.Range("B23").Value, "#,##0.0") & "M" & vbNewLine & _
           "Equity Value:      $" & Format(ws.Range("B25").Value, "#,##0.0") & "M" & vbNewLine & _
           "Value per Share:   $" & Format(ws.Range("B27").Value, "#,##0.00"), _
           vbInformation, "Valuation Summary"

End Sub


' ============================================================
' SUB: GenerateReport
' Exports the workbook as a PDF to the reports folder.
' ============================================================

Public Sub GenerateReport()

    Dim reportPath As String
    Dim ticker As String
    ticker = ThisWorkbook.Sheets(ASSUMP_SHEET).Range("B6").Value

    ' Create reports directory if not exists
    Dim fullDir As String
    fullDir = ThisWorkbook.Path & "\" & REPORT_DIR
    If Dir(fullDir, vbDirectory) = "" Then
        MkDir fullDir
    End If

    reportPath = fullDir & ticker & "_ValuationReport_" & _
                 Format(Now, "YYYYMMDD") & ".pdf"

    ' Select sheets to include in report
    Dim sheetsToExport() As String
    ReDim sheetsToExport(4)
    sheetsToExport(0) = ASSUMP_SHEET
    sheetsToExport(1) = INCOME_SHEET
    sheetsToExport(2) = DCF_SHEET
    sheetsToExport(3) = SENS_SHEET
    sheetsToExport(4) = COMPS_SHEET

    ThisWorkbook.Sheets(sheetsToExport).Select

    ' Export PDF
    On Error GoTo ErrHandler
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=reportPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    Sheets(ASSUMP_SHEET).Select
    MsgBox "✅ Report exported:" & vbNewLine & reportPath, vbInformation, "Report Generated"
    Exit Sub

ErrHandler:
    MsgBox "❌ Error exporting PDF: " & Err.Description, vbCritical, "Export Error"

End Sub


' ============================================================
' SUB: ImportFinancials
' Pastes raw financial data from clipboard and normalizes format.
' ============================================================

Public Sub ImportFinancials()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(ASSUMP_SHEET)

    ' Prompt user for key inputs
    Dim revenue As Double
    Dim ebitda As Double
    Dim netDebt As Double
    Dim shares As Double

    revenue = CDbl(InputBox("Enter LTM Revenue ($M):", "Import Financials", "12600"))
    If revenue = 0 Then Exit Sub

    ebitda  = CDbl(InputBox("Enter LTM EBITDA ($M):", "Import Financials", "3400"))
    netDebt = CDbl(InputBox("Enter Net Debt ($M):", "Import Financials", "2100"))
    shares  = CDbl(InputBox("Enter Diluted Shares Outstanding (M):", "Import Financials", "500"))

    ' Write to assumption cells
    ws.Range("B23").Value = revenue
    ws.Range("B40").Value = netDebt + ws.Range("B41").Value  ' Gross debt
    ws.Range("B43").Value = shares

    ' Derive EBITDA margin
    If revenue > 0 Then
        ws.Range("B34").Value = ebitda / revenue
    End If

    Application.Calculate

    MsgBox "✅ Financial data imported successfully." & vbNewLine & _
           "Revenue: $" & Format(revenue, "#,##0") & "M" & vbNewLine & _
           "EBITDA Margin: " & Format(ebitda / revenue, "0.0%"), _
           vbInformation, "Import Complete"

End Sub


' ============================================================
' SUB: SensitivityAnalysis
' Reruns sensitivity table for current model inputs.
' ============================================================

Public Sub SensitivityAnalysis()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SENS_SHEET)

    ' Highlight base case (WACC = row 8, g = col 4, i.e. 2.5% growth)
    ws.Cells.Interior.ColorIndex = xlNone  ' Reset

    Dim wacc As Double
    wacc = ThisWorkbook.Sheets(DCF_SHEET).Range("B6").Value

    ' Re-apply header fills
    ws.Rows(4).Interior.Color = RGB(30, 111, 191)
    ws.Range("A5:A11").Interior.Color = RGB(30, 111, 191)

    ' Find and highlight base case cell
    Dim r As Integer, c As Integer
    For r = 5 To 11
        For c = 2 To 6
            Dim cellWACC As Double, cellG As Double
            cellWACC = ws.Cells(r, 1).Value
            cellG    = ws.Cells(4, c).Value
            If Abs(cellWACC - wacc) < 0.001 And Abs(cellG - 0.025) < 0.001 Then
                ws.Cells(r, c).Interior.Color = RGB(240, 165, 0)
                ws.Cells(r, c).Font.Bold = True
                ws.Cells(r, c).Font.Color = RGB(255, 255, 255)
            End If
        Next c
    Next r

    Application.ScreenUpdating = True
    MsgBox "✅ Sensitivity table refreshed. Base case highlighted in amber.", _
           vbInformation, "Sensitivity Analysis"

End Sub


' ============================================================
' SUB: ResetAssumptions
' Resets all blue input cells to default values.
' ============================================================

Public Sub ResetAssumptions()

    Dim answer As Integer
    answer = MsgBox("Reset all assumptions to default values?", _
                    vbYesNo + vbQuestion, "Reset Assumptions")
    If answer = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(ASSUMP_SHEET)

    ' WACC inputs
    ws.Range("B11").Value = 0.04   ' Risk-free rate
    ws.Range("B12").Value = 0.055  ' ERP
    ws.Range("B13").Value = 1.15   ' Beta
    ws.Range("B14").Value = 0.045  ' Cost of debt
    ws.Range("B15").Value = 0.21   ' Tax rate
    ws.Range("B16").Value = 0.25   ' Debt weight
    ws.Range("B18").Value = 0.01   ' Size premium
    ws.Range("B19").Value = 0.00   ' Specific risk

    ' DCF
    ws.Range("B24").Value = 0.025  ' Terminal growth
    ws.Range("B25").Value = 12.0   ' Exit multiple

    ' FCF drivers
    ws.Range("B29").Value = 0.12
    ws.Range("B30").Value = 0.10
    ws.Range("B31").Value = 0.09
    ws.Range("B32").Value = 0.08
    ws.Range("B33").Value = 0.07
    ws.Range("B34").Value = 0.27   ' EBITDA margin
    ws.Range("B35").Value = 0.05   ' D&A
    ws.Range("B36").Value = 0.06   ' CapEx
    ws.Range("B37").Value = 0.015  ' NWC

    Application.Calculate
    MsgBox "✅ Assumptions reset to defaults.", vbInformation, "Reset Complete"

End Sub


' ============================================================
' FUNCTION: FormatMillions
' Utility to format a value as $XM or $XB.
' ============================================================

Private Function FormatMillions(value As Double) As String
    If Abs(value) >= 1000 Then
        FormatMillions = "$" & Format(value / 1000, "0.0") & "B"
    Else
        FormatMillions = "$" & Format(value, "#,##0") & "M"
    End If
End Function

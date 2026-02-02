Attribute VB_Name = "RetailSupplier_DashboardAutomation"
' =============================================================================
' Retail Supplier Performance & Inventory Optimization - One-Click Refresh & Report
' P&G → Walmart, Sam's Club, Costco. Overstock = cash locked; Understock = lost sales.
' Run: RefreshAndReport() from a button or Alt+F8
' =============================================================================

Option Explicit

' Safe conversion: handles Empty, text, and cell errors (#N/A, #DIV/0!, etc.)
Private Function SafeNum(cellVal As Variant) As Double
    On Error Resume Next
    If IsEmpty(cellVal) Or IsNull(cellVal) Then SafeNum = 0: Exit Function
    If IsError(cellVal) Then SafeNum = 0: Exit Function
    SafeNum = Val(cellVal)
    On Error GoTo 0
End Function

Private Function SafeLong(cellVal As Variant) As Long
    On Error Resume Next
    If IsEmpty(cellVal) Or IsNull(cellVal) Then SafeLong = 0: Exit Function
    If IsError(cellVal) Then SafeLong = 0: Exit Function
    SafeLong = CLng(Val(cellVal))
    On Error GoTo 0
End Function

Public Sub RefreshAndReport()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 1. Refresh all Power Query connections
    RefreshAllQueries

    ' 2. Refresh PivotTables
    RefreshAllPivotTables

    ' 3. Recalculate workbook
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    ' 4. Highlight stockout risks (Ending Inventory <= Reorder Point)
    On Error Resume Next
    HighlightStockoutRisks
    On Error GoTo ErrHandler

    ' 5. Apply conditional formatting for KPI risks (margin, turnover, etc.)
    On Error Resume Next
    ApplyKPIConditionalFormatting
    On Error GoTo ErrHandler

    ' 6. Run business alerts
    RunBusinessAlerts

    Done:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
ErrHandler:
    MsgBox "Error in RefreshAndReport: " & Err.Description, vbCritical
    Resume Done
End Sub

Private Sub RefreshAllQueries()
    Dim q As WorkbookQuery
    On Error Resume Next
    For Each q In ThisWorkbook.Queries
        q.Refresh
    Next q
    On Error GoTo 0
End Sub

Private Sub RefreshAllPivotTables()
    Dim pt As PivotTable
    On Error Resume Next
    For Each pt In ThisWorkbook.PivotTables
        pt.RefreshTable
    Next pt
    On Error GoTo 0
End Sub

' Highlight stockout risks: Ending Inventory <= Reorder Point
Private Sub HighlightStockoutRisks()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim endCol As Long, reorderCol As Long, endInvCol As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Inventory")
    If ws Is Nothing Then Set ws = ThisWorkbook.Sheets("Merged")
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    endInvCol = ColByName(ws, "Ending Inventory")
    reorderCol = ColByName(ws, "Reorder Point")
    If endInvCol = 0 Then endInvCol = 6
    If reorderCol = 0 Then reorderCol = 7

    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column))
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlExpression, Formula1:="=AND($" & ColLetter(endInvCol) & "2<=$" & ColLetter(reorderCol) & "2,$" & ColLetter(endInvCol) & "2<>"""")"
        .FormatConditions(1).Interior.Color = RGB(255, 220, 200)
    End With
End Sub

Private Function ColLetter(colNum As Long) As String
    Dim n As Long, s As String
    n = colNum
    s = ""
    Do While n > 0
        s = Chr(64 + ((n - 1) Mod 26) + 1) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function

Private Sub ApplyKPIConditionalFormatting()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("KPIs")
    If ws Is Nothing Then Set ws = ThisWorkbook.Sheets("Executive Summary")
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0

    ' Only format Gross Margin % (column D) - not Revenue (column B)
    With ws.Range("D2:D10")
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0.15"
        .FormatConditions(1).Interior.Color = RGB(255, 200, 200)
    End With
End Sub

' =============================================================================
' BUSINESS ALERTS
' =============================================================================

Public Sub RunBusinessAlerts()
    Dim msg As String
    msg = ""

    msg = msg & CheckLowInventoryRiskTopSKUs()
    msg = msg & CheckMarginErosionDueToDiscounting()
    msg = msg & CheckStockoutRisk()

    If Len(Trim(msg)) > 0 Then
        MsgBox "Retail Supplier Performance – Alerts" & vbCrLf & vbCrLf & msg, vbExclamation, "Supplier Performance"
    Else
        MsgBox "No critical alerts. All checks passed.", vbInformation
    End If
End Sub

' Alert: "Low inventory risk for top SKUs"
Private Function CheckLowInventoryRiskTopSKUs() As String
    Dim wsInv As Worksheet
    Dim lastRow As Long, r As Long, cnt As Long
    Dim sku As String, endInv As Long, reorder As Long
    Dim out As String
    Dim skuCol As Long, endCol As Long, reorderCol As Long

    out = ""
    On Error Resume Next
    Set wsInv = ThisWorkbook.Sheets("Inventory")
    If wsInv Is Nothing Then Set wsInv = ThisWorkbook.Sheets("Merged")
    If wsInv Is Nothing Then CheckLowInventoryRiskTopSKUs = "": Exit Function
    On Error GoTo 0

    lastRow = wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    skuCol = ColByName(wsInv, "SKU")
    endCol = ColByName(wsInv, "Ending Inventory")
    reorderCol = ColByName(wsInv, "Reorder Point")
    If skuCol = 0 Then skuCol = 1
    If endCol = 0 Then endCol = 6
    If reorderCol = 0 Then reorderCol = 7

    cnt = 0
    Dim lowList As String
    lowList = ""
    For r = 2 To lastRow
        endInv = SafeLong(wsInv.Cells(r, endCol).Value)
        reorder = SafeLong(wsInv.Cells(r, reorderCol).Value)
        If endInv <= reorder And endInv >= 0 And reorder > 0 Then
            cnt = cnt + 1
            If cnt <= 5 Then
                sku = wsInv.Cells(r, skuCol).Value
                lowList = lowList & "  - " & sku & " (End: " & endInv & ", Reorder: " & reorder & ")" & vbCrLf
            End If
        End If
    Next r

    If cnt >= 5 Then
        out = "LOW INVENTORY RISK FOR TOP SKUs" & vbCrLf
        out = out & cnt & " SKU/Store combinations with Ending Inventory at or below Reorder Point." & vbCrLf
        out = out & "Examples:" & vbCrLf & lowList
        out = out & "Risk: Lost sales, retailer penalties." & vbCrLf & vbCrLf
    End If
    CheckLowInventoryRiskTopSKUs = out
End Function

' Alert: "Margin erosion due to discounting"
Private Function CheckMarginErosionDueToDiscounting() As String
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim marginPct As Double
    Dim threshold As Double, marginCol As Long
    Dim out As String

    threshold = 15
    out = ""
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Sales")
    If ws Is Nothing Then Set ws = ThisWorkbook.Sheets("Merged")
    If ws Is Nothing Then CheckMarginErosionDueToDiscounting = "": Exit Function
    On Error GoTo 0

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    marginCol = ColByName(ws, "Gross Margin %")
    If marginCol = 0 Then marginCol = 10

    Dim lowMarginCount As Long
    lowMarginCount = 0
    For r = 2 To lastRow
        marginPct = SafeNum(ws.Cells(r, marginCol).Value)
        If marginPct < threshold Then
            lowMarginCount = lowMarginCount + 1
        End If
    Next r

    If lowMarginCount >= 50 Then
        out = "MARGIN EROSION DUE TO DISCOUNTING" & vbCrLf
        out = out & lowMarginCount & " transactions with Gross Margin % below " & threshold & "% target." & vbCrLf
        out = out & "Review promo depth and costs; consider list vs promo price impact." & vbCrLf & vbCrLf
    End If
    CheckMarginErosionDueToDiscounting = out
End Function

' Alert: Stockout risk (Ending Inventory = 0 or very low)
Private Function CheckStockoutRisk() As String
    Dim wsInv As Worksheet
    Dim lastRow As Long, r As Long, zeroCount As Long
    Dim out As String

    zeroCount = 0
    out = ""
    On Error Resume Next
    Set wsInv = ThisWorkbook.Sheets("Inventory")
    If wsInv Is Nothing Then Set wsInv = ThisWorkbook.Sheets("Merged")
    If wsInv Is Nothing Then CheckStockoutRisk = "": Exit Function
    On Error GoTo 0

    lastRow = wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim endCol As Long
    endCol = ColByName(wsInv, "Ending Inventory")
    If endCol = 0 Then endCol = 6

    For r = 2 To lastRow
        If SafeLong(wsInv.Cells(r, endCol).Value) = 0 Then zeroCount = zeroCount + 1
    Next r

    If zeroCount >= 3 Then
        out = "STOCKOUT RISK" & vbCrLf
        out = out & zeroCount & " SKU/Store combinations with zero ending inventory." & vbCrLf
        out = out & "Expedite replenishment to avoid lost sales." & vbCrLf & vbCrLf
    End If
    CheckStockoutRisk = out
End Function

Private Function ColByName(ws As Worksheet, colName As String) As Long
    Dim c As Long, v As Variant
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        v = ws.Cells(1, c).Value
        If Not IsError(v) And Not IsNull(v) Then
            If LCase(Trim(CStr(v))) = LCase(colName) Then
                ColByName = c
                Exit Function
            End If
        End If
    Next c
    ColByName = 0
End Function

' =============================================================================
' EXPORT SUMMARY REPORT
' =============================================================================

Public Sub ExportSummaryReport()
    Dim summarySheet As String
    Dim pdfPath As String

    summarySheet = "Executive Summary"
    If SheetExists(summarySheet) = False Then
        summarySheet = "KPIs"
        If SheetExists(summarySheet) = False Then
            MsgBox "Create a sheet named 'Executive Summary' or 'KPIs' with your summary before exporting.", vbExclamation
            Exit Sub
        End If
    End If

    pdfPath = ThisWorkbook.Path & "\Retail_Supplier_Summary_" & Format(Now, "yyyymmdd_hhnn") & ".pdf"

    On Error Resume Next
    ThisWorkbook.Sheets(summarySheet).ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard
    If Err.Number <> 0 Then MsgBox "PDF export failed: " & Err.Description
    On Error GoTo 0

    MsgBox "Export complete." & vbCrLf & "PDF: " & pdfPath, vbInformation
End Sub

Private Function SheetExists(shtName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(shtName)
    SheetExists = (Not ws Is Nothing)
    On Error GoTo 0
End Function

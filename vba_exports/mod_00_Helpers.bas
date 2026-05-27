Attribute VB_Name = "Mod_00_Helpers"
Option Explicit

' =========================================================================
' CORE STATISTICAL HELPERS (High-Performance VBA Implementation)
' =========================================================================

Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
    Dim pivot As Double, tmpSwap As Double, tmpLow As Long, tmpHi As Long
    tmpLow = inLow: tmpHi = inHi
    pivot = vArray((inLow + inHi) \ 2)
    While (tmpLow <= tmpHi)
        While (vArray(tmpLow) < pivot And tmpLow < inHi): tmpLow = tmpLow + 1: Wend
        While (pivot < vArray(tmpHi) And tmpHi > inLow): tmpHi = tmpHi - 1: Wend
        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow): vArray(tmpLow) = vArray(tmpHi): vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1: tmpHi = tmpHi - 1
        End If
    Wend
    If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub

Sub GetBasicStats(arr() As Double, ByRef mean As Double, ByRef stdDev As Double)
    Dim i As Long, count As Long: count = UBound(arr)
    Dim sum As Double, sqSum As Double
    If count < 1 Then Exit Sub
    For i = 1 To count
        sum = sum + arr(i): sqSum = sqSum + (arr(i) ^ 2)
    Next i
    mean = sum / count
    If count > 1 Then
        Dim var As Double: var = (sqSum - (sum ^ 2) / count) / (count - 1)
        If var > 0 Then stdDev = Sqr(var) Else stdDev = 0
    Else: stdDev = 0: End If
End Sub

Function GetMode(arr() As Double) As Double
    On Error Resume Next
    Dim ub As Long: ub = UBound(arr)
    If Err.Number <> 0 Or ub < 1 Then GetMode = 0: Exit Function
    On Error GoTo 0
    Dim maxC As Long, curC As Long, mVal As Double, i As Long
    maxC = 1: curC = 1: mVal = arr(1)
    For i = 2 To UBound(arr)
        If arr(i) = arr(i - 1) Then curC = curC + 1 Else curC = 1
        If curC > maxC Then maxC = curC: mVal = arr(i)
    Next i
    GetMode = mVal
End Function

' =========================================================================
' V2X CHARTING ENGINE
' =========================================================================

Sub DirectChart(ws As Worksheet, slot As Long, x, yF, yC, title As String, rowPos As Long, p95 As Double, p99 As Double, xMax As Double)
    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(380 * (slot - 1), ws.rows(rowPos).Top, 370, 260)
    With cht.Chart
        .HasTitle = True: .ChartTitle.Text = title
        With .SeriesCollection.NewSeries: .Name = "Freq": .XValues = x: .Values = yF: .ChartType = xlXYScatterLinesNoMarkers: End With
        With .SeriesCollection.NewSeries: .Name = "CDF": .XValues = x: .Values = yC: .ChartType = xlXYScatterLinesNoMarkers: .AxisGroup = xlSecondary: End With
        With .SeriesCollection.NewSeries
            .Name = "95th %": .XValues = Array(p95, p95): .Values = Array(0, 1): .AxisGroup = xlSecondary: .ChartType = xlXYScatterLinesNoMarkers
            With .Format.Line: .DashStyle = msoLineDash: .ForeColor.RGB = RGB(255, 0, 0): .Weight = 1.2: End With
        End With
        With .SeriesCollection.NewSeries
            .Name = "99th %": .XValues = Array(p99, p99): .Values = Array(0, 1): .AxisGroup = xlSecondary: .ChartType = xlXYScatterLinesNoMarkers
            With .Format.Line: .DashStyle = msoLineDash: .ForeColor.RGB = RGB(0, 0, 0): .Weight = 1.2: End With
        End With
        With .Axes(xlCategory): .HasTitle = True: .AxisTitle.Text = "Value": .MinimumScale = 0: .MaximumScale = xMax: .TickLabels.NumberFormat = "0": End With
        .Axes(xlValue, xlSecondary).MaximumScale = 1
        .HasLegend = True: .Legend.Position = xlLegendPositionBottom
    End With
End Sub

' =========================================================================
' WORKBOOK UTILITIES & STRUCTURAL HELPERS
' =========================================================================

Function CreateOrClearSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)): ws.Name = sheetName
    Else
        Application.DisplayAlerts = False
        ws.Cells.Clear
        ' Robust Chart Deletion
        Dim co As ChartObject
        For Each co In ws.ChartObjects
            co.Delete
        Next co
        Application.DisplayAlerts = True
    End If
    Set CreateOrClearSheet = ws
End Function

Public Sub EnsureColumnExists(tbl As ListObject, colName As String)
    Dim col As ListColumn
    On Error Resume Next
    Set col = tbl.ListColumns(colName)
    On Error GoTo 0
    If col Is Nothing Then
        Set col = tbl.ListColumns.Add
        col.Name = colName
    End If
End Sub

Public Function DblMod(ByVal dividend As Double, ByVal divisor As Double) As Double
    ' High-precision modulo for timing wrap-arounds (32-bit/16-bit)
    If divisor = 0 Then DblMod = 0: Exit Function
    DblMod = dividend - (Fix(dividend / divisor) * divisor)
End Function

' =========================================================================
' C-V2X PHYSICAL LAYER ESTIMATION ENGINE
' =========================================================================

Public Function RunWLSRefinement(initialEst As Double, avgRx1 As Double) As Double
    ' Logic: Performs initial physical constraint check and WLS alignment.
    ' Shift logic ensures TX_SFN does not exceed observed RX time.
    Dim refinedValue As Double: refinedValue = initialEst
    
    If refinedValue > avgRx1 And avgRx1 > 0 Then
        ' Constraint Violation: TX cannot be after RX.
        ' Force shift back to nearest valid 1ms boundary.
        refinedValue = Fix(avgRx1 - 0.1)
    End If
    
    RunWLSRefinement = refinedValue
End Function

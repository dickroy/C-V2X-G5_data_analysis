Attribute VB_Name = "mod_16_LinRegTXQTvsTX_SFN_est"
Option Explicit

'==========================================================
' LinReg_TXQTIMEvsTX_SFN_est V1.0.0
' Fixes:
' - regression title moved to column K (11) to avoid overwriting preliminary title
' - table headers Parameter / Value centered
' - column 1 entries right-justified
' - previous regression output cleared before redraw
'==========================================================

Private Const LINREG_OUTPUT_SHEET As String = "TX_SFN est Log"
Private Const RESID_BIN_WIDTH_MS As Double = 0.01

Public data As Variant
Public filteredCount As Long
Public idxTXID As Long
Public idxTXQ As Long
Public idxSFNCol As Long

Public dictS2V As Object
Public dictVC As Object

Public Sub Run_LinReg_TXQTIMEvsTX_SFN_est(Optional ByRef rerunInitialEstimation As Boolean = False)
    Dim wsSrc As Worksheet
    Dim lo As ListObject
    Dim srcData As Variant
    Dim r As Long
    Dim nRows As Long

    Set wsSrc = ThisWorkbook.Worksheets("ExpResults")
    Set lo = wsSrc.ListObjects("ExpResultsTable")
    If lo.DataBodyRange Is Nothing Then Exit Sub

    idxTXID = lo.ListColumns("TX_ID").Index
    idxTXQ = lo.ListColumns("TXQTIME").Index
    idxSFNCol = lo.ListColumns("TX_SFN_est").Index

    srcData = lo.DataBodyRange.Value
    nRows = UBound(srcData, 1)

    ReDim data(1 To nRows, 1 To 3)
    For r = 1 To nRows
        data(r, 1) = srcData(r, idxTXID)
        data(r, 2) = srcData(r, idxTXQ)
        data(r, 3) = srcData(r, idxSFNCol)
    Next r

    filteredCount = nRows
    idxTXID = 1
    idxTXQ = 2
    idxSFNCol = 3

    LinReg_TXTIMEvsTX_SFN_est rerunInitialEstimation
End Sub

Public Sub LinReg_TXTIMEvsTX_SFN_est(Optional ByRef rerunInitialEstimation As Boolean = False)
    Dim wsOut As Worksheet
    Dim vendorDict As Object
    Dim vendorKeys As Variant
    Dim i As Long
    Dim vendorCount As Long
    Dim txKey As String

    If IsEmpty(data) Then Exit Sub
    If filteredCount <= 0 Then Exit Sub

    Set wsOut = ThisWorkbook.Worksheets(LINREG_OUTPUT_SHEET)
    ClearPreviousRegressionOutput wsOut

    Set vendorDict = CreateObject("Scripting.Dictionary")

    If dictS2V Is Nothing Then
        Set dictS2V = CreateObject("Scripting.Dictionary")
        BuildTxToVendorMap dictS2V
    ElseIf dictS2V.count = 0 Then
        BuildTxToVendorMap dictS2V
    End If

    For i = 1 To filteredCount
        txKey = UCase$(Trim$(CStr(data(i, idxTXID))))
        If Len(txKey) > 0 Then
            If dictS2V.Exists(txKey) Then
                vendorDict(Trim$(CStr(dictS2V(txKey)))) = True
            End If
        End If
    Next i

    If vendorDict.count = 0 Then Exit Sub

    vendorKeys = vendorDict.Keys
    SortVariantStringArray vendorKeys

    Application.ScreenUpdating = False
    For vendorCount = LBound(vendorKeys) To UBound(vendorKeys)
        ProcessVendorSection wsOut, CStr(vendorKeys(vendorCount)), vendorCount - LBound(vendorKeys)
    Next vendorCount
    Application.ScreenUpdating = True

    PromptAndPersistVendorTXParams vendorKeys, rerunInitialEstimation
End Sub

Private Sub ClearPreviousRegressionOutput(ByVal ws As Worksheet)
    Dim co As ChartObject
    Dim shp As Shape
    Dim i As Long

    For i = ws.ChartObjects.count To 1 Step -1
        Set co = ws.ChartObjects(i)
        If co.Left >= ws.Columns("M").Left Then co.Delete
    Next i

    For i = ws.Shapes.count To 1 Step -1
        Set shp = ws.Shapes(i)
        If shp.Left >= ws.Columns("M").Left Then shp.Delete
    Next i

    ws.Range("M:R").ClearContents
End Sub

Private Sub ProcessVendorSection( _
    ByVal ws As Worksheet, _
    ByVal vendorID As String, _
    ByVal vendorIndex As Long)

    Dim xVals() As Double, yVals() As Double, resid() As Double
    Dim n As Long, r As Long, i As Long
    Dim txID As String, rowVendor As String
    Dim x As Double, y As Double
    Dim slope As Double, intercept As Double
    Dim varSlope As Double, varIntercept As Double
    Dim rss As Double
    Dim muRes As Double, sigmaRes As Double
    Dim stdSlope As Double, stdIntercept As Double
    Dim titleRow As Long
    Dim chartTop As Double
    Dim histLeft As Double

    titleRow = 3 + (vendorIndex * 24)
    chartTop = ws.rows(titleRow + 2).Top
    histLeft = ws.Columns("M").Left

    ws.Cells(titleRow - 2, "M").Value = "Linear Regression of TXQTIME against Initial TX_SFN_est"
    ws.Cells(titleRow - 2, "M").Font.Bold = True
    ws.Cells(titleRow - 2, "M").Font.Size = 13

    ws.Cells(titleRow - 1, "M").Value = "Vendor " & vendorID & " LINREG: TXQTIME vs TX_SFN_est"
    ws.Cells(titleRow - 1, "M").Font.Bold = True

    ReDim xVals(1 To filteredCount)
    ReDim yVals(1 To filteredCount)
    ReDim resid(1 To filteredCount)

    n = 0
    For r = 1 To filteredCount
        txID = UCase$(Trim$(CStr(data(r, idxTXID))))
        If Len(txID) > 0 Then
            If dictS2V.Exists(txID) Then
                rowVendor = Trim$(CStr(dictS2V(txID)))
                If StrComp(rowVendor, vendorID, vbTextCompare) = 0 Then
                    If IsNumeric(data(r, idxSFNCol)) And IsNumeric(data(r, idxTXQ)) Then
                        x = CDbl(data(r, idxSFNCol))
                        y = CDbl(data(r, idxTXQ))
                        n = n + 1
                        xVals(n) = x
                        yVals(n) = y
                    End If
                End If
            End If
        End If
    Next r

    If n < 2 Then Exit Sub

    ReDim Preserve xVals(1 To n)
    ReDim Preserve yVals(1 To n)
    ReDim Preserve resid(1 To n)

    FitOLS xVals, yVals, n, slope, intercept, varSlope, varIntercept, rss

    For i = 1 To n
        resid(i) = yVals(i) - (intercept + slope * xVals(i))
    Next i

    muRes = MeanOfDoubles(resid, n)
    sigmaRes = SampleStdDevOfDoubles(resid, n, muRes)

    stdSlope = Sqr(varSlope)
    stdIntercept = Sqr(varIntercept)

    DrawResidualHistogram ws, histLeft, chartTop, 420#, 250#, resid, n, muRes, sigmaRes, vendorID
    WriteRegressionTableAtCell ws, "Q5", vendorID, n, slope, intercept, stdSlope, stdIntercept, rss
    WriteResidualStatsTableAtCell ws, "Q14", vendorID, n, muRes, sigmaRes
End Sub

Private Sub PromptAndPersistVendorTXParams(ByRef vendorKeys As Variant, ByRef rerunInitialEstimation As Boolean)
    Dim loTxTable As ListObject
    Dim vIdx As Long
    Dim vendorID As String
    Dim currentMean As Double
    Dim currentSigma As Double
    Dim rowWalk As Long
    Dim promptMsg As String
    Dim userResponse As String
    Dim splitVals() As String
    Dim newMean As Double
    Dim newSigma As Double

    On Error Resume Next
    Set loTxTable = ThisWorkbook.Sheets("Exp Config & Data Proc Params").ListObjects("VendorID2TXTproc")
    On Error GoTo 0
    If loTxTable Is Nothing Then Exit Sub

    For vIdx = LBound(vendorKeys) To UBound(vendorKeys)
        vendorID = Trim$(CStr(vendorKeys(vIdx)))
        If vendorID = "" Then GoTo NextVendorPrompt

        currentMean = 0#
        currentSigma = 0#
        For rowWalk = 1 To loTxTable.ListRows.Count
            If Trim$(CStr(loTxTable.DataBodyRange.Cells(rowWalk, 1).Value)) = vendorID Then
                If IsNumeric(loTxTable.DataBodyRange.Cells(rowWalk, 2).Value) Then
                    currentMean = CDbl(loTxTable.DataBodyRange.Cells(rowWalk, 2).Value)
                End If
                If IsNumeric(loTxTable.DataBodyRange.Cells(rowWalk, 3).Value) Then
                    currentSigma = CDbl(loTxTable.DataBodyRange.Cells(rowWalk, 3).Value)
                End If
                Exit For
            End If
        Next rowWalk

        promptMsg = "LinReg results are rendered for Vendor " & vendorID & "." & vbCrLf & vbCrLf & _
                    "Current TX parameters for Vendor " & vendorID & ":" & vbCrLf & _
                    "  Tproc Mean : " & currentMean & " ms" & vbCrLf & _
                    "  Tproc Sigma: " & currentSigma & " ms" & vbCrLf & vbCrLf & _
                    "Enter new Mean,Sigma to update (e.g., 4.2,0.85)." & vbCrLf & _
                    "Press Cancel or leave blank to keep current values."

        userResponse = InputBox(promptMsg, "TX Tproc Update - Vendor " & vendorID, currentMean & "," & currentSigma)

        If Trim$(userResponse) <> "" Then
            splitVals = Split(userResponse, ",")
            If UBound(splitVals) = 1 Then
                If IsNumeric(Trim$(splitVals(0))) And IsNumeric(Trim$(splitVals(1))) Then
                    newMean = CDbl(Trim$(splitVals(0)))
                    newSigma = CDbl(Trim$(splitVals(1)))
                    If newMean <> currentMean Or newSigma <> currentSigma Then
                        For rowWalk = 1 To loTxTable.ListRows.Count
                            If Trim$(CStr(loTxTable.DataBodyRange.Cells(rowWalk, 1).Value)) = vendorID Then
                                loTxTable.DataBodyRange.Cells(rowWalk, 2).Value = newMean
                                loTxTable.DataBodyRange.Cells(rowWalk, 3).Value = newSigma
                                Exit For
                            End If
                        Next rowWalk
                        rerunInitialEstimation = True
                    End If
                Else
                    MsgBox "Invalid input for Vendor " & vendorID & ". Expected format: Mean,Sigma (e.g., 4.2,0.85). Values were not updated.", _
                           vbExclamation, "TX Tproc Update"
                End If
            Else
                MsgBox "Invalid input for Vendor " & vendorID & ". Expected exactly two values separated by a comma. Values were not updated.", _
                       vbExclamation, "TX Tproc Update"
            End If
        End If

NextVendorPrompt:
    Next vIdx
End Sub

Private Sub FitOLS( _
    ByRef xVals() As Double, _
    ByRef yVals() As Double, _
    ByVal n As Long, _
    ByRef slope As Double, _
    ByRef intercept As Double, _
    ByRef varSlope As Double, _
    ByRef varIntercept As Double, _
    ByRef rss As Double)

    Dim i As Long
    Dim sumX As Double, sumY As Double, sumXX As Double, sumXY As Double
    Dim xBar As Double, yBar As Double
    Dim sxx As Double, sxy As Double
    Dim yHat As Double
    Dim s2 As Double

    For i = 1 To n
        sumX = sumX + xVals(i)
        sumY = sumY + yVals(i)
        sumXX = sumXX + xVals(i) * xVals(i)
        sumXY = sumXY + xVals(i) * yVals(i)
    Next i

    xBar = sumX / n
    yBar = sumY / n
    sxx = sumXX - n * xBar * xBar
    sxy = sumXY - n * xBar * yBar

    If Abs(sxx) < 1E-30 Then
        slope = 0#
        intercept = yBar
        rss = 0#
        varSlope = 0#
        varIntercept = 0#
        Exit Sub
    End If

    slope = sxy / sxx
    intercept = yBar - slope * xBar

    rss = 0#
    For i = 1 To n
        yHat = intercept + slope * xVals(i)
        rss = rss + (yVals(i) - yHat) ^ 2
    Next i

    If n > 2 Then
        s2 = rss / (n - 2)
        varSlope = s2 / sxx
        varIntercept = s2 * (1# / n + (xBar * xBar / sxx))
    Else
        varSlope = 0#
        varIntercept = 0#
    End If
End Sub

Private Sub DrawResidualHistogram( _
    ByVal ws As Worksheet, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthPos As Double, _
    ByVal heightPos As Double, _
    ByRef resid() As Double, _
    ByVal n As Long, _
    ByVal muRes As Double, _
    ByVal sigmaRes As Double, _
    ByVal vendorID As String)

    Dim i As Long
    Dim minR As Double, maxR As Double
    Dim binLo As Double, binHi As Double
    Dim nBins As Long
    Dim counts() As Long
    Dim xCenters() As Double
    Dim histFreq() As Double
    Dim gaussY() As Double
    Dim curBin As Long
    Dim co As ChartObject

    minR = resid(1)
    maxR = resid(1)
    For i = 2 To n
        If resid(i) < minR Then minR = resid(i)
        If resid(i) > maxR Then maxR = resid(i)
    Next i

    binLo = Int(minR / RESID_BIN_WIDTH_MS) * RESID_BIN_WIDTH_MS
    binHi = (Int(maxR / RESID_BIN_WIDTH_MS) + 1#) * RESID_BIN_WIDTH_MS
    If binHi <= binLo Then binHi = binLo + RESID_BIN_WIDTH_MS

    nBins = CLng((binHi - binLo) / RESID_BIN_WIDTH_MS)
    If nBins < 1 Then nBins = 1

    ReDim counts(1 To nBins)
    ReDim xCenters(1 To nBins)
    ReDim histFreq(1 To nBins)
    ReDim gaussY(1 To nBins)

    For i = 1 To n
        curBin = CLng(Fix((resid(i) - binLo) / RESID_BIN_WIDTH_MS)) + 1
        If curBin < 1 Then curBin = 1
        If curBin > nBins Then curBin = nBins
        counts(curBin) = counts(curBin) + 1
    Next i

    For i = 1 To nBins
        xCenters(i) = binLo + (i - 0.5) * RESID_BIN_WIDTH_MS
        histFreq(i) = counts(i)
        If sigmaRes > 0# Then
            gaussY(i) = n * RESID_BIN_WIDTH_MS * NormalPdf(xCenters(i), muRes, sigmaRes)
        Else
            gaussY(i) = 0#
        End If
    Next i

    Set co = ws.ChartObjects.Add(leftPos, topPos, widthPos, heightPos)
    With co.Chart
        .HasTitle = True
        .ChartTitle.Text = "Vendor " & vendorID & " residuals: TXQTIME - fitted"
        .HasLegend = True

        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Frequency"
        .SeriesCollection(1).XValues = xCenters
        .SeriesCollection(1).Values = histFreq
        .SeriesCollection(1).ChartType = xlColumnClustered

        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "Gaussian Fit"
        .SeriesCollection(2).XValues = xCenters
        .SeriesCollection(2).Values = gaussY
        .SeriesCollection(2).ChartType = xlLine
        .SeriesCollection(2).AxisGroup = xlSecondary
        .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(220, 0, 0)
        .SeriesCollection(2).Format.Line.Weight = 2#

        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Residual (ms)"
            .TickLabels.Orientation = 90
        End With

        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Count"
        End With

        On Error Resume Next
        With .Axes(xlValue, xlSecondary)
            .HasTitle = True
            .AxisTitle.Text = "Gaussian PDF scaled"
        End With
        On Error GoTo 0

        .Legend.Position = xlLegendPositionBottom
    End With
End Sub

Private Sub WriteRegressionTableAtCell( _
    ByVal ws As Worksheet, _
    ByVal anchorCell As String, _
    ByVal vendorID As String, _
    ByVal n As Long, _
    ByVal slope As Double, _
    ByVal intercept As Double, _
    ByVal stdSlope As Double, _
    ByVal stdIntercept As Double, _
    ByVal rss As Double)

    Dim c As Range
    Set c = ws.Range(anchorCell)

    c.Value = "Vendor " & vendorID & " LS Regression"
    c.Font.Bold = True

    c.Offset(1, 0).Value = "Parameter"
    c.Offset(1, 1).Value = "Value"
    c.Offset(1, 0).Resize(1, 2).Font.Bold = True
    c.Offset(1, 0).HorizontalAlignment = xlCenter
    c.Offset(1, 1).HorizontalAlignment = xlCenter

    c.Offset(2, 0).Value = "N"
    c.Offset(2, 1).Value = n
    c.Offset(3, 0).Value = "Slope"
    c.Offset(3, 1).Value = slope
    c.Offset(4, 0).Value = "Intercept"
    c.Offset(4, 1).Value = intercept
    c.Offset(5, 0).Value = "Std. Dev. (Slope)"
    c.Offset(5, 1).Value = stdSlope
    c.Offset(6, 0).Value = "Std. Dev. (Intercept)"
    c.Offset(6, 1).Value = stdIntercept
    c.Offset(7, 0).Value = "RSS"
    c.Offset(7, 1).Value = rss

    c.Offset(2, 0).HorizontalAlignment = xlRight
    c.Offset(3, 0).HorizontalAlignment = xlRight
    c.Offset(4, 0).HorizontalAlignment = xlRight
    c.Offset(5, 0).HorizontalAlignment = xlRight
    c.Offset(6, 0).HorizontalAlignment = xlRight
    c.Offset(7, 0).HorizontalAlignment = xlRight

    c.Offset(2, 1).NumberFormat = "0"
    c.Offset(3, 1).NumberFormat = "0.000000000"
    c.Offset(4, 1).NumberFormat = "0.000000000"
    c.Offset(5, 1).NumberFormat = "0.000000000"
    c.Offset(6, 1).NumberFormat = "0.000000000"
    c.Offset(7, 1).NumberFormat = "0.000000000"
End Sub

Private Sub WriteResidualStatsTableAtCell( _
    ByVal ws As Worksheet, _
    ByVal anchorCell As String, _
    ByVal vendorID As String, _
    ByVal n As Long, _
    ByVal muRes As Double, _
    ByVal sigmaRes As Double)

    Dim c As Range
    Set c = ws.Range(anchorCell)

    c.Value = "Vendor " & vendorID & " residual Gaussian fit"
    c.Font.Bold = True

    c.Offset(1, 0).Value = "Parameter"
    c.Offset(1, 1).Value = "Value"
    c.Offset(1, 0).Resize(1, 2).Font.Bold = True
    c.Offset(1, 0).HorizontalAlignment = xlCenter
    c.Offset(1, 1).HorizontalAlignment = xlCenter

    c.Offset(2, 0).Value = "N"
    c.Offset(2, 1).Value = n
    c.Offset(3, 0).Value = "mu"
    c.Offset(3, 1).Value = muRes
    c.Offset(4, 0).Value = "sigma"
    c.Offset(4, 1).Value = sigmaRes

    c.Offset(2, 0).HorizontalAlignment = xlRight
    c.Offset(3, 0).HorizontalAlignment = xlRight
    c.Offset(4, 0).HorizontalAlignment = xlRight

    c.Offset(2, 1).NumberFormat = "#,##0"
    c.Offset(3, 1).NumberFormat = "0.000000"
    c.Offset(4, 1).NumberFormat = "0.000000"
End Sub

Private Sub BuildTxToVendorMap(ByRef outDict As Object)
    Dim ws As Worksheet
    Dim rng As Range
    Dim arr As Variant
    Dim i As Long
    Dim txID As String, vendorID As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Exp Config & Data Proc Params")
    Set rng = ws.Range("StationID2VendorID")
    On Error GoTo 0

    If rng Is Nothing Then Exit Sub

    arr = rng.Value
    For i = 1 To UBound(arr, 1)
        txID = UCase$(Trim$(CStr(arr(i, 1))))
        vendorID = Trim$(CStr(arr(i, 2)))
        If Len(txID) > 0 Then outDict(txID) = vendorID
    Next i
End Sub

Private Function MeanOfDoubles(ByRef arr() As Double, ByVal n As Long) As Double
    Dim i As Long, s As Double
    For i = 1 To n
        s = s + arr(i)
    Next i
    If n > 0 Then MeanOfDoubles = s / n
End Function

Private Function SampleStdDevOfDoubles(ByRef arr() As Double, ByVal n As Long, ByVal meanVal As Double) As Double
    Dim i As Long, s As Double
    If n < 2 Then Exit Function
    For i = 1 To n
        s = s + (arr(i) - meanVal) ^ 2
    Next i
    SampleStdDevOfDoubles = Sqr(s / (n - 1))
End Function

Private Function NormalPdf(ByVal x As Double, ByVal mu As Double, ByVal sigma As Double) As Double
    If sigma <= 0# Then Exit Function
    NormalPdf = (1# / (sigma * Sqr(2# * 3.14159265358979))) * Exp(-0.5 * ((x - mu) / sigma) ^ 2)
End Function

Private Sub SortVariantStringArray(ByRef arr As Variant)
    Dim changed As Boolean
    Dim i As Long
    Dim tmp As Variant

    If IsEmpty(arr) Then Exit Sub
    If UBound(arr) <= LBound(arr) Then Exit Sub

    Do
        changed = False
        For i = LBound(arr) To UBound(arr) - 1
            If CStr(arr(i)) > CStr(arr(i + 1)) Then
                tmp = arr(i)
                arr(i) = arr(i + 1)
                arr(i + 1) = tmp
                changed = True
            End If
        Next i
    Loop While changed
End Sub


Attribute VB_Name = "mod_06_PktCollision"
' Iteration: 12.6 (BOOKMARKED - FULL)
' Changes: 1. Added TLS Fit Parameters summary table (Quadratic 'a', Linear Slope/Intercept).
'          2. Renamed "Weighted Coll Prob (%)" table header to "Coll Prob (%)".
'          3. Maintained all Vertical/Horizontal axis labels and integer time formatting.
'          4. FULL CODE OUTPUT - Includes all mandatory helper functions.

Option Explicit

' --- MANDATORY HELPER FUNCTIONS ---

Function GetLinearTLS(xArr As Variant, yArr As Variant, threshold As Double) As Variant
    Dim i As Long, n As Long, sumX As Double, sumY As Double, meanX As Double, meanY As Double, sxx As Double, syy As Double, sxy As Double
    Dim vX() As Double, vY() As Double
    For i = 1 To UBound(xArr, 1)
        If xArr(i, 1) >= threshold Then
            n = n + 1: ReDim Preserve vX(1 To n): ReDim Preserve vY(1 To n)
            vX(n) = xArr(i, 1): vY(n) = yArr(i, 1): sumX = sumX + vX(n): sumY = sumY + vY(n)
        End If
    Next i
    If n < 2 Then GetLinearTLS = Array(0, 0, 0): Exit Function
    meanX = sumX / n: meanY = sumY / n
    For i = 1 To n: sxx = sxx + (vX(i) - meanX) ^ 2: syy = syy + (vY(i) - meanY) ^ 2: sxy = sxy + (vX(i) - meanX) * (vY(i) - meanY): Next i
    Dim slope As Double
    If sxy <> 0 Then slope = (syy - sxx + Sqr((syy - sxx) ^ 2 + 4 * sxy ^ 2)) / (2 * sxy) Else slope = 0
    GetLinearTLS = Array(0, slope, meanY - slope * meanX)
End Function

Sub QuickSortArray(vArray As Variant, ArrMin As Long, ArrMax As Long, ColumnToSort As Integer)
    Dim i As Long, j As Long, pivot As Variant, temp As Variant, col As Integer
    i = ArrMin: j = ArrMax: pivot = vArray((ArrMin + ArrMax) \ 2, ColumnToSort)
    Do While i <= j
        Do While vArray(i, ColumnToSort) < pivot And i < ArrMax: i = i + 1: Loop
        Do While pivot < vArray(j, ColumnToSort) And j > ArrMin: j = j - 1: Loop
        If i <= j Then
            For col = LBound(vArray, 2) To UBound(vArray, 2)
                temp = vArray(i, col): vArray(i, col) = vArray(j, col): vArray(j, col) = temp
            Next col
            i = i + 1: j = j - 1
        End If
    Loop
    If ArrMin < j Then QuickSortArray vArray, ArrMin, j, ColumnToSort
    If i < ArrMax Then QuickSortArray vArray, i, ArrMax, ColumnToSort
End Sub


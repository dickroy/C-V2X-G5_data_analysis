Option Explicit

' Version: V1.0.3-DEBUG
' DEBUG build of TX_SFN conflict resolution.
' Rule: GROUPS SHALL NEVER BE SPLIT.
' Purpose: trace every pool/cset decision and row move to diagnose why valid groups were altered.
Private Const MODULE_VERSION_TXSFNCR As String = "V1.0.3-DEBUG"
Private Const DEBUG_TXSFNCR As Boolean = True
Private Const DEBUG_VERBOSE_TXSFNCR As Boolean = True
Private Const DEBUG_LOG_SHEET As String = "CR DEBUG"
Private Const DEBUG_WARN_ROW_THRESHOLD As Long = 100
Private Const NO_RX_TIME As Double = -1#
Private Const PROGRESS_STEP_ROWS As Long = 10000

#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter_TXSFNCR Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency_TXSFNCR Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef lpFrequency As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter_TXSFNCR Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency_TXSFNCR Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef lpFrequency As Currency) As Long
#End If

Private Type TCsetAnalysis
    MovesRequired As Long
    candidateCount As Long
    candidateRows() As Long
End Type

Private mData As Variant
Private mFilteredCount As Long
Private mIdxTXID As Long
Private mIdxTXQ As Long
Private mIdxSFNCol As Long
Private mIdxLEN As Long
Private mIdxTXperSFN As Long
Private mIdxGen As Long
Private mIdxAvg As Long
Private mIdxTotLat As Long
Private mIdxRxCnt As Long
Private mTxBitmap As String
Private mBitmapLen As Long
Private mRxStationIDs() As Long
Private mRxDataColIdx() As Long
Private mActiveRxCount As Long
Private mDictS2V As Object
Private mDictVC As Object
Private mDictA2P As Object
Private mDictP2R As Object
Private mDictP2Sigma As Object
Private mInitialSFN() As Long
Private mCurrentSFN() As Long
Private mRowTXID() As Long
Private mRowTXQTime() As Double
Private mRowNsch() As Long
Private mRowPduKey() As String
Private mRowMinRxTime() As Double
Private mRowOriginalIndex() As Long
Private mRowValidInput() As Boolean
Private mRowRXCount() As Long
Private mWritten() As Boolean
Private mPoolRows() As Long
Private mPoolCount As Long
Private mPoolMinSFN As Long
Private mPoolMaxSFN As Long
Private mPoolCenter As Double
Private mPoolCestStartRows() As Long
Private mPoolCestEndRows() As Long
Private mPoolCestSFN() As Long
Private mPoolCestCount As Long
Private mOutputData() As Variant
Private mOutputCount As Long
Private mOutputWritePos As Long
Private mScanPos As Long
Private mPoolCountResolved As Long
Private mMaxObservedPoolSize As Long
Private mRemainingViolations As Long
Private mUnresolvedAttemptCount As Long
Private mDiagCount As Long
Private mFindConflictSeconds As Double
Private mBuildPoolSeconds As Double
Private mResolvePoolSeconds As Double
Private mWritePoolSeconds As Double
Private mFinalizeSeconds As Double
Private mResolveEntirePoolSeconds As Double
Private mResolveOneCsetSeconds As Double
Private mTryPlaceMoveSeconds As Double
Private mLegalCheckSeconds As Double
Private mBucketExcludeSeconds As Double
Private mBucketAddSeconds As Double
Private mValidateBucketSeconds As Double
Private mResolveOneCsetExtractSeconds As Double
Private mResolveOneCsetPostSeconds As Double
Private mDebugWs As Worksheet
Private mDebugRow As Long
Private mOriginalConflictCount As Long

Public Sub Run_TX_SFNConflictResolution_DEBUG()
    TX_SFNConflictResolution_DEBUG
End Sub

Public Sub TX_SFNConflictResolution_DEBUG( _
    Optional ByRef data As Variant, Optional ByVal filteredCount As Long = 0, Optional ByVal idxSFNCol As Long = 0, Optional ByVal idxTXID As Long = 0, _
    Optional ByVal idxTXQ As Long = 0, Optional ByVal idxLEN As Long = 0, Optional ByVal idxTXperSFN As Long = 0, Optional ByVal idxRxCnt As Long = 0, _
    Optional ByVal idxAvg As Long = 0, Optional ByVal idxTotLat As Long = 0, Optional ByVal idxGen As Long = 0, Optional ByRef rxDataColIdx() As Long, _
    Optional ByRef rxStationIDs() As Long, Optional ByVal activeRxCount As Long = 0, Optional ByRef dictS2V As Object, Optional ByRef dictVC As Object, _
    Optional ByRef dictA2P As Object, Optional ByRef dictP2R As Object, Optional ByRef dictP2Sigma As Object, Optional ByVal txBitmap As String = "", _
    Optional ByVal bitmapLen As Long = 0, Optional ByRef elapsedSeconds As Double)

    Dim startTime As Double
    Dim conflictStart As Long
    Dim poolMoved As Boolean
    Dim continueAllowed As VbMsgBoxResult
    Dim t0 As Double
    Dim tPhase As Double
    Dim totalFindSeconds As Double
    Dim totalBuildSeconds As Double
    Dim totalResolveSeconds As Double
    Dim totalWriteSeconds As Double
    Dim totalFinalSeconds As Double
    Dim totalLoopCount As Long
    Dim oldStatusBar As Variant

    If filteredCount > DEBUG_WARN_ROW_THRESHOLD Then
        continueAllowed = MsgBox("WARNING: This routine creates copious outputs." & vbCrLf & vbCrLf & _
                                 "Are you sure you want to continue?" & vbCrLf & vbCrLf & _
                                 "Rows to be processed: " & filteredCount, _
                                 vbYesNo + vbExclamation, _
                                 "TX_SFN Conflict Resolution DEBUG")
        If continueAllowed = vbNo Then Exit Sub
    End If

    startTime = MicroTimer_TXSFNCR()
    oldStatusBar = Application.StatusBar
    Application.StatusBar = "TX_SFN conflict resolution running..."

    InitializeContext data, filteredCount, idxSFNCol, idxTXID, idxTXQ, idxLEN, idxTXperSFN, idxRxCnt, idxAvg, idxTotLat, idxGen, rxDataColIdx, rxStationIDs, activeRxCount, dictS2V, dictVC, dictA2P, dictP2R, dictP2Sigma, txBitmap, bitmapLen
    SetupDebugLog
    LogDebugHeader

    If Not ValidateInputMonotoneTXSFN() Then
        LogDebug "ABORT", "Input SFN sequence is not monotone.", "", "", "", ""
        GoTo CleanExit
    End If

    PrepareRowDerivedData
    InitializeOutputBuffer
    LogDebug "INIT", "rows=" & mFilteredCount, "bitmapLen=" & mBitmapLen, "activeRx=" & mActiveRxCount, "", ""

    Do
        If (mScanPos Mod PROGRESS_STEP_ROWS) = 0 Then UpdateProgressBar mScanPos
        t0 = MicroTimer_TXSFNCR()

        tPhase = MicroTimer_TXSFNCR()
        conflictStart = FindNextConflictStart()
        totalFindSeconds = totalFindSeconds + (MicroTimer_TXSFNCR() - tPhase)

        If conflictStart <= 0 Then Exit Do

        BuildPoolFromConflictStart conflictStart
        LogPoolState conflictStart

        tPhase = MicroTimer_TXSFNCR()
        poolMoved = ResolveEntirePool()
        totalResolveSeconds = totalResolveSeconds + (MicroTimer_TXSFNCR() - tPhase)
        LogDebug "POOL_RESOLVE_COMPLETE", "moved=" & CStr(poolMoved), "poolCount=" & mPoolCount, "resolvedCount=" & mPoolCountResolved, "unresolvedAttempts=" & mUnresolvedAttemptCount, ""

        tPhase = MicroTimer_TXSFNCR()
        WriteResolvedPoolToOutput
        totalWriteSeconds = totalWriteSeconds + (MicroTimer_TXSFNCR() - tPhase)

        If mPoolCount > 0 Then
            mScanPos = mPoolRows(mPoolCount) + 1
        End If

        totalLoopCount = totalLoopCount + 1
        mFindConflictSeconds = mFindConflictSeconds + (MicroTimer_TXSFNCR() - t0)
    Loop

    t0 = MicroTimer_TXSFNCR()
    WriteUnwrittenRowsToOutput
    RecomputeFinalTXperSFN
    ValidateResolvedRXTimingOnly
    FinalizeOutputVariant data
    totalFinalSeconds = MicroTimer_TXSFNCR() - t0

    mFinalizeSeconds = mFinalizeSeconds + totalFinalSeconds
    elapsedSeconds = MicroTimer_TXSFNCR() - startTime

    Debug.Print "SUMMARY", "loops", totalLoopCount, "find", Format$(totalFindSeconds, "0.000"), "build", Format$(totalBuildSeconds, "0.000"), "resolve", Format$(totalResolveSeconds, "0.000"), "write", Format$(totalWriteSeconds, "0.000"), "final", Format$(totalFinalSeconds, "0.000"), "total", Format$(elapsedSeconds, "0.000")
    WriteTimingResultsToConflictResolutionLog totalLoopCount, totalFindSeconds, totalBuildSeconds, totalResolveSeconds, totalWriteSeconds, totalFinalSeconds, elapsedSeconds

CleanExit:
    Application.StatusBar = oldStatusBar
End Sub

Private Sub SetupDebugLog()
    On Error Resume Next
    Set mDebugWs = ThisWorkbook.Worksheets(DEBUG_LOG_SHEET)
    On Error GoTo 0
    If mDebugWs Is Nothing Then
        Set mDebugWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        mDebugWs.Name = DEBUG_LOG_SHEET
    End If
    mDebugWs.Cells.Clear
    mDebugRow = 1
End Sub

Private Sub LogDebugHeader()
    LogDebug "Event", "A", "B", "C", "D", "E"
    LogDebug "RULE", "GROUPS SHALL NEVER BE SPLIT", "", "", "", ""
    LogDebug "VERSION", MODULE_VERSION_TXSFNCR, "", "", "", ""
End Sub

Private Sub LogDebug(ByVal eventName As String, ByVal a As String, ByVal b As String, ByVal c As String, ByVal d As String, ByVal e As String)
    If DEBUG_TXSFNCR Then Debug.Print eventName, a, b, c, d, e
    If mDebugWs Is Nothing Then Exit Sub
    mDebugWs.Cells(mDebugRow, 1).Value = eventName
    mDebugWs.Cells(mDebugRow, 2).Value = a
    mDebugWs.Cells(mDebugRow, 3).Value = b
    mDebugWs.Cells(mDebugRow, 4).Value = c
    mDebugWs.Cells(mDebugRow, 5).Value = d
    mDebugWs.Cells(mDebugRow, 6).Value = e
    mDebugRow = mDebugRow + 1
End Sub

Private Sub LogPoolState(ByVal conflictStart As Long)
    Dim i As Long
    Dim s As String
    s = "rows="
    For i = 1 To mPoolCount
        s = s & mPoolRows(i) & IIf(i < mPoolCount, ",", "")
    Next i
    LogDebug "POOL_BUILT", "seedConflictStart=" & conflictStart, "poolCount=" & mPoolCount, "minSFN=" & mPoolMinSFN, "maxSFN=" & mPoolMaxSFN, s
End Sub

Private Function MicroTimer_TXSFNCR() As Double
    Dim cyTicks As Currency
    Dim cyFreq As Currency
    If QueryPerformanceFrequency_TXSFNCR(cyFreq) <> 0 Then
        QueryPerformanceCounter_TXSFNCR cyTicks
        If cyFreq > 0 Then MicroTimer_TXSFNCR = cyTicks / cyFreq
    End If
End Function

Private Sub InitializeContext(ByRef data As Variant, ByVal filteredCount As Long, ByVal idxSFNCol As Long, ByVal idxTXID As Long, ByVal idxTXQ As Long, ByVal idxLEN As Long, ByVal idxTXperSFN As Long, ByVal idxRxCnt As Long, ByVal idxAvg As Long, ByVal idxTotLat As Long, ByVal idxGen As Long, ByRef rxDataColIdx() As Long, ByRef rxStationIDs() As Long, ByVal activeRxCount As Long, ByRef dictS2V As Object, ByRef dictVC As Object, ByRef dictA2P As Object, ByRef dictP2R As Object, ByRef dictP2Sigma As Object, ByVal txBitmap As String, ByVal bitmapLen As Long)
    mData = data
    mFilteredCount = filteredCount
    mIdxSFNCol = idxSFNCol
    mIdxTXID = idxTXID
    mIdxTXQ = idxTXQ
    mIdxLEN = idxLEN
    mIdxTXperSFN = idxTXperSFN
    mIdxRxCnt = idxRxCnt
    mIdxAvg = idxAvg
    mIdxTotLat = idxTotLat
    mIdxGen = idxGen
    mRxDataColIdx = rxDataColIdx
    mRxStationIDs = rxStationIDs
    mActiveRxCount = activeRxCount
    Set mDictS2V = dictS2V
    Set mDictVC = dictVC
    Set mDictA2P = dictA2P
    Set mDictP2R = dictP2R
    Set mDictP2Sigma = dictP2Sigma
    mTxBitmap = txBitmap
    mBitmapLen = bitmapLen
    mScanPos = 1
    mOutputWritePos = 1
    mPoolCount = 0
    mPoolCestCount = 0
    mOriginalConflictCount = 0
    If DEBUG_TXSFNCR Then Debug.Print "TX_SFNCR init: filteredCount=" & mFilteredCount & " bitmapLen=" & mBitmapLen
End Sub

Private Function ValidateInputMonotoneTXSFN() As Boolean
    Dim r As Long
    Dim prevVal As Long
    Dim curVal As Long

    ValidateInputMonotoneTXSFN = True
    If mFilteredCount <= 1 Then Exit Function

    prevVal = CLng(mData(1, mIdxSFNCol))
    For r = 2 To mFilteredCount
        curVal = CLng(mData(r, mIdxSFNCol))
        If curVal < prevVal Then
            ValidateInputMonotoneTXSFN = False
            Exit Function
        End If
        prevVal = curVal
    Next r
End Function

Private Sub PrepareRowDerivedData()
    Dim r As Long

    ReDim mInitialSFN(1 To mFilteredCount)
    ReDim mCurrentSFN(1 To mFilteredCount)
    ReDim mRowTXID(1 To mFilteredCount)
    ReDim mRowTXQTime(1 To mFilteredCount)
    ReDim mRowNsch(1 To mFilteredCount)
    ReDim mRowPduKey(1 To mFilteredCount)
    ReDim mRowMinRxTime(1 To mFilteredCount)
    ReDim mRowOriginalIndex(1 To mFilteredCount)
    ReDim mRowValidInput(1 To mFilteredCount)
    ReDim mRowRXCount(1 To mFilteredCount)
    ReDim mWritten(1 To mFilteredCount)

    For r = 1 To mFilteredCount
        mInitialSFN(r) = CLng(mData(r, mIdxSFNCol))
        mCurrentSFN(r) = mInitialSFN(r)
        mRowTXID(r) = CLng(mData(r, mIdxTXID))
        mRowTXQTime(r) = CDbl(mData(r, mIdxTXQ))
        mRowNsch(r) = CLng(mData(r, mIdxLEN))
        mRowPduKey(r) = CStr(mData(r, mIdxLEN))
        mRowMinRxTime(r) = GetRowMinRxTime(r)
        mRowOriginalIndex(r) = r
        mRowValidInput(r) = True
        If mIdxRxCnt > 0 Then
            mRowRXCount(r) = CLng(mData(r, mIdxRxCnt))
        Else
            mRowRXCount(r) = 0
        End If
    Next r
End Sub

Private Function FindNextConflictStart() As Long
    Dim r As Long
    For r = mScanPos To mFilteredCount - 1
        If mCurrentSFN(r) = mCurrentSFN(r + 1) Then
            FindNextConflictStart = r
            Exit Function
        End If
    Next r
End Function

Private Sub BuildPoolFromConflictStart(ByVal startRow As Long)
    Dim leftRow As Long
    Dim rightRow As Long
    Dim i As Long

    If startRow < 1 Or startRow >= mFilteredCount Then Exit Sub

    leftRow = startRow
    rightRow = startRow + 1

    Do While leftRow > 1 And mCurrentSFN(leftRow - 1) = mCurrentSFN(leftRow)
        leftRow = leftRow - 1
    Loop

    Do While rightRow < mFilteredCount And mCurrentSFN(rightRow + 1) = mCurrentSFN(rightRow)
        rightRow = rightRow + 1
    Loop

    mPoolCount = rightRow - leftRow + 1
    If mPoolCount <= 0 Then Exit Sub

    ReDim mPoolRows(1 To mPoolCount)
    For i = 1 To mPoolCount
        mPoolRows(i) = leftRow + i - 1
    Next i

    mPoolMinSFN = mCurrentSFN(leftRow)
    mPoolMaxSFN = mCurrentSFN(rightRow)
    mPoolCenter = (mPoolMinSFN + mPoolMaxSFN) / 2#

    If mPoolCount > mMaxObservedPoolSize Then
        mMaxObservedPoolSize = mPoolCount
    End If

    If DEBUG_TXSFNCR Then
        Debug.Print "POOL built", "startRow", startRow, "leftRow", leftRow, "rightRow", rightRow, "poolCount", mPoolCount, "minSFN", mPoolMinSFN, "maxSFN", mPoolMaxSFN
    End If

    BuildPoolCests
End Sub

Private Sub BuildPoolCests()
    Dim i As Long
    Dim r As Long
    Dim startIdx As Long
    Dim curSFN As Long

    If mPoolCount <= 0 Then Exit Sub

    ReDim mPoolCestStartRows(1 To mPoolCount)
    ReDim mPoolCestEndRows(1 To mPoolCount)
    ReDim mPoolCestSFN(1 To mPoolCount)

    mPoolCestCount = 0
    i = 1

    Do While i <= mPoolCount
        startIdx = i
        r = mPoolRows(i)
        curSFN = mCurrentSFN(r)

        Do While i < mPoolCount
            If mCurrentSFN(mPoolRows(i + 1)) <> curSFN Then Exit Do
            i = i + 1
        Loop

        mPoolCestCount = mPoolCestCount + 1
        mPoolCestStartRows(mPoolCestCount) = startIdx
        mPoolCestEndRows(mPoolCestCount) = i
        mPoolCestSFN(mPoolCestCount) = curSFN

        If DEBUG_TXSFNCR Then
            Debug.Print "CEST chunk", "idx", mPoolCestCount, "rows", startIdx & ".." & i, "sfn", curSFN, "count", (i - startIdx + 1)
        End If

        i = i + 1
    Loop
End Sub

Private Function ResolveEntirePool() As Boolean
    Dim cestIdx As Long
    Dim orderedCests() As Long
    Dim i As Long
    Dim t As Double

    t = MicroTimer_TXSFNCR()
    ResolveEntirePool = False

    If mPoolCestCount <= 0 Then
        mResolveEntirePoolSeconds = mResolveEntirePoolSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    orderedCests = GetOrderedCestIndexes()

    For i = LBound(orderedCests) To UBound(orderedCests)
        cestIdx = orderedCests(i)
        If cestIdx > 0 Then
            If ResolveOneCset(cestIdx) Then
                ResolveEntirePool = True
            End If
        End If
    Next i

    mResolveEntirePoolSeconds = mResolveEntirePoolSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function ResolveOneCset(ByVal cestIdx As Long) As Boolean
    Dim rows() As Long
    Dim rowCount As Long
    Dim changed As Boolean
    Dim movedRowIdx As Long
    Dim sourceSFN As Long
    Dim t As Double
    Dim tExtract As Double
    Dim tPost As Double

    t = MicroTimer_TXSFNCR()
    ResolveOneCset = False

    If cestIdx < 1 Or cestIdx > mPoolCestCount Then
        mResolveOneCsetSeconds = mResolveOneCsetSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    rowCount = mPoolCestEndRows(cestIdx) - mPoolCestStartRows(cestIdx) + 1
    If rowCount <= 1 Then
        mResolveOneCsetSeconds = mResolveOneCsetSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    tExtract = MicroTimer_TXSFNCR()
    rows = ExtractPoolRows(mPoolCestStartRows(cestIdx), mPoolCestEndRows(cestIdx))
    sourceSFN = mPoolCestSFN(cestIdx)
    mResolveOneCsetExtractSeconds = mResolveOneCsetExtractSeconds + (MicroTimer_TXSFNCR() - tExtract)

    If DEBUG_TXSFNCR Then
        Debug.Print "TryResolveCset", "cestIdx", cestIdx, "rowCount", rowCount, "sourceSFN", sourceSFN
    End If

    changed = TryPlaceOneMovedRow_NoSourceRetest(rows, rowCount, sourceSFN, movedRowIdx)

    If changed Then
        ResolveOneCset = True
        mPoolCountResolved = mPoolCountResolved + 1

        If DEBUG_TXSFNCR Then
            Debug.Print "CSET resolved", "cestIdx", cestIdx, "movedRowIdx", movedRowIdx, "newSFN", mCurrentSFN(movedRowIdx)
        End If

        tPost = MicroTimer_TXSFNCR()
        BuildPoolFromConflictStart movedRowIdx
        mResolveOneCsetPostSeconds = mResolveOneCsetPostSeconds + (MicroTimer_TXSFNCR() - tPost)
    Else
        mUnresolvedAttemptCount = mUnresolvedAttemptCount + 1
    End If

    mResolveOneCsetSeconds = mResolveOneCsetSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function GetOrderedCestIndexes() As Long()
    Dim idxs() As Long
    Dim used() As Boolean
    Dim i As Long
    Dim outPos As Long
    Dim bestIdx As Long
    Dim bestScore As Double
    Dim score As Double
    Dim haveBest As Boolean

    If mPoolCestCount <= 0 Then
        ReDim idxs(1 To 1)
        idxs(1) = 0
        GetOrderedCestIndexes = idxs
        Exit Function
    End If

    ReDim idxs(1 To mPoolCestCount)
    ReDim used(1 To mPoolCestCount)
    outPos = 1

    Do While outPos <= mPoolCestCount
        haveBest = False
        bestIdx = 0
        bestScore = 0#

        For i = 1 To mPoolCestCount
            If Not used(i) Then
                score = Abs(CDbl(mPoolCestSFN(i)) - mPoolCenter)
                If Not haveBest Then
                    haveBest = True
                    bestIdx = i
                    bestScore = score
                ElseIf score < bestScore Then
                    bestIdx = i
                    bestScore = score
                ElseIf score = bestScore Then
                    If mPoolCestSFN(i) < mPoolCestSFN(bestIdx) Then
                        bestIdx = i
                        bestScore = score
                    End If
                End If
            End If
        Next i

        idxs(outPos) = bestIdx
        used(bestIdx) = True
        outPos = outPos + 1
    Loop

    If DEBUG_TXSFNCR Then
        Dim s As String
        s = "CSET order: "
        For i = LBound(idxs) To UBound(idxs)
            s = s & idxs(i) & IIf(i < UBound(idxs), ",", "")
        Next i
        Debug.Print s & " | poolCenter=" & mPoolCenter
    End If

    GetOrderedCestIndexes = idxs
End Function

Private Function TryPlaceOneMovedRow_NoSourceRetest(ByRef candidateRows() As Long, ByVal candidateCount As Long, ByVal sourceSFN As Long, ByRef movedRowIdx As Long) As Boolean
    Dim i As Long
    Dim rowIdx As Long
    Dim testSFN As Long
    Dim delta As Long
    Dim originalSFN As Long
    Dim maxDec As Long
    Dim maxInc As Long
    Dim maxDelta As Long
    Dim t As Double

    t = MicroTimer_TXSFNCR()
    movedRowIdx = 0

    maxDec = GetNamedLong("maxTX_SFN_est_decrement")
    maxInc = GetNamedLong("maxTX_SFN_est_increment")
    maxDelta = IIf(maxDec > maxInc, maxDec, maxInc)

    If DEBUG_TXSFNCR Then
        Debug.Print "MOVE RANGE", "sourceSFN", sourceSFN, "maxDec", maxDec, "maxInc", maxInc, "maxDelta", maxDelta
    End If

    If maxDelta <= 0 Then
        mTryPlaceMoveSeconds = mTryPlaceMoveSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    For i = 1 To candidateCount
        rowIdx = candidateRows(i)
        originalSFN = mCurrentSFN(rowIdx)

        For delta = 1 To maxDelta
            If delta <= maxDec Then
                testSFN = sourceSFN - delta
                If DEBUG_TXSFNCR Then Debug.Print "TEST MOVE", "rowIdx", rowIdx, "sourceSFN", sourceSFN, "delta", delta, "testSFN", testSFN, "dir", "-"
                If IsOneMovedRowPlacementLegal_NoSourceRetest(rowIdx, sourceSFN, testSFN) Then
                    mCurrentSFN(rowIdx) = testSFN
                    If DoesMovedRowFormValidLocalGroup(rowIdx) Then
                        movedRowIdx = rowIdx
                        If DEBUG_TXSFNCR Then Debug.Print "ACCEPT move", "rowIdx", rowIdx, "sourceSFN", sourceSFN, "testSFN", testSFN, "minRx", mRowMinRxTime(rowIdx), "txID", mRowTXID(rowIdx)
                        TryPlaceOneMovedRow_NoSourceRetest = True
                        mTryPlaceMoveSeconds = mTryPlaceMoveSeconds + (MicroTimer_TXSFNCR() - t)
                        Exit Function
                    End If
                    mCurrentSFN(rowIdx) = originalSFN
                End If
            End If

            If delta <= maxInc Then
                testSFN = sourceSFN + delta
                If DEBUG_TXSFNCR Then Debug.Print "TEST MOVE", "rowIdx", rowIdx, "sourceSFN", sourceSFN, "delta", delta, "testSFN", testSFN, "dir", "+"
                If IsOneMovedRowPlacementLegal_NoSourceRetest(rowIdx, sourceSFN, testSFN) Then
                    mCurrentSFN(rowIdx) = testSFN
                    If DoesMovedRowFormValidLocalGroup(rowIdx) Then
                        movedRowIdx = rowIdx
                        If DEBUG_TXSFNCR Then Debug.Print "ACCEPT move", "rowIdx", rowIdx, "sourceSFN", sourceSFN, "testSFN", testSFN, "minRx", mRowMinRxTime(rowIdx), "txID", mRowTXID(rowIdx)
                        TryPlaceOneMovedRow_NoSourceRetest = True
                        mTryPlaceMoveSeconds = mTryPlaceMoveSeconds + (MicroTimer_TXSFNCR() - t)
                        Exit Function
                    End If
                    mCurrentSFN(rowIdx) = originalSFN
                End If
            End If
        Next delta
    Next i

    mTryPlaceMoveSeconds = mTryPlaceMoveSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function IsOneMovedRowPlacementLegal_NoSourceRetest(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean
    Dim t As Double
    t = MicroTimer_TXSFNCR()

    If rowIdx < 1 Or rowIdx > mFilteredCount Then
        If DEBUG_TXSFNCR Then Debug.Print "REJECT legal-check", "rowIdx", rowIdx, "reason", "out of range"
        mLegalCheckSeconds = mLegalCheckSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    If testSFN = sourceSFN Then
        If DEBUG_TXSFNCR Then Debug.Print "REJECT legal-check", "rowIdx", rowIdx, "reason", "testSFN equals sourceSFN", "sfn", sourceSFN
        mLegalCheckSeconds = mLegalCheckSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    If Not IsBitmapSFNAllowed(testSFN) Then
        If DEBUG_TXSFNCR Then Debug.Print "REJECT legal-check", "rowIdx", rowIdx, "reason", "bitmap disallows", "testSFN", testSFN
        mLegalCheckSeconds = mLegalCheckSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    If Not EvaluatePoolBucketExcludingRow(sourceSFN, rowIdx) Then
        If DEBUG_TXSFNCR Then Debug.Print "REJECT legal-check", "rowIdx", rowIdx, "reason", "excluding-source bucket invalid", "sourceSFN", sourceSFN
        mLegalCheckSeconds = mLegalCheckSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    If Not EvaluatePoolBucketWithAddedRow(testSFN, rowIdx) Then
        If DEBUG_TXSFNCR Then Debug.Print "REJECT legal-check", "rowIdx", rowIdx, "reason", "added-row bucket invalid", "testSFN", testSFN
        mLegalCheckSeconds = mLegalCheckSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    IsOneMovedRowPlacementLegal_NoSourceRetest = True
    mLegalCheckSeconds = mLegalCheckSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function TryForwardEscapeMove_NoSourceRetest(ByVal sourceSFN As Long, ByVal rowIdx As Long) As Boolean
    TryForwardEscapeMove_NoSourceRetest = False
End Function

Private Function ResolveTripleSplitDeterministic(ByRef rows() As Long, ByVal sourceSFN As Long) As Boolean
    Dim i As Long

    For i = LBound(rows) To UBound(rows)
        If IsOneMovedRowPlacementLegal_NoSourceRetest(rows(i), sourceSFN, sourceSFN - 1) Then
            mCurrentSFN(rows(i)) = sourceSFN - 1
            ResolveTripleSplitDeterministic = True
            Exit Function
        End If

        If IsOneMovedRowPlacementLegal_NoSourceRetest(rows(i), sourceSFN, sourceSFN + 1) Then
            mCurrentSFN(rows(i)) = sourceSFN + 1
            ResolveTripleSplitDeterministic = True
            Exit Function
        End If
    Next i
End Function

Private Function IsPoolForcedSplitFirstMoveLegal(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean
    IsPoolForcedSplitFirstMoveLegal = IsOneMovedRowPlacementLegal_NoSourceRetest(rowIdx, sourceSFN, testSFN)
End Function

Private Function IsPoolSingleRowPlacementLegal(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean
    IsPoolSingleRowPlacementLegal = IsOneMovedRowPlacementLegal_NoSourceRetest(rowIdx, sourceSFN, testSFN)
End Function

Private Function EvaluatePoolBucketExcludingRow(ByVal sfnVal As Long, ByVal excludeRowIdx As Long) As Boolean
    Dim rows() As Long
    Dim rowCount As Long
    Dim t As Double

    t = MicroTimer_TXSFNCR()
    rows = CollectRowsForSFN(sfnVal, excludeRowIdx, -1, rowCount)
    EvaluatePoolBucketExcludingRow = ValidateBucketRows(rows, rowCount)
    mBucketExcludeSeconds = mBucketExcludeSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function EvaluatePoolBucketWithAddedRow(ByVal sfnVal As Long, ByVal addedRowIdx As Long) As Boolean
    Dim rows() As Long
    Dim rowCount As Long
    Dim t As Double

    t = MicroTimer_TXSFNCR()
    rows = CollectRowsForSFN(sfnVal, -1, addedRowIdx, rowCount)
    EvaluatePoolBucketWithAddedRow = ValidateBucketRows(rows, rowCount)
    mBucketAddSeconds = mBucketAddSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function IsMoveWithinRowBounds(ByVal rowIdx As Long, ByVal testSFN As Long) As Boolean
    Dim prevSFN As Long
    Dim nextSFN As Long

    If rowIdx < 1 Or rowIdx > mFilteredCount Then Exit Function
    If testSFN < 0 Then Exit Function

    prevSFN = testSFN
    nextSFN = testSFN

    If rowIdx > 1 Then prevSFN = mCurrentSFN(rowIdx - 1)
    If rowIdx < mFilteredCount Then nextSFN = mCurrentSFN(rowIdx + 1)

    If rowIdx > 1 And testSFN < prevSFN Then Exit Function
    If rowIdx < mFilteredCount And testSFN > nextSFN Then Exit Function

    IsMoveWithinRowBounds = True
End Function

Private Function IsBitmapSFNAllowed(ByVal testSFN As Long) As Boolean
    Dim bitIdx As Long
    Dim bitChar As String

    If mBitmapLen <= 0 Or LenB(mTxBitmap) = 0 Then
        IsBitmapSFNAllowed = True
        Exit Function
    End If

    bitIdx = ((testSFN Mod mBitmapLen) + mBitmapLen) Mod mBitmapLen + 1

    If bitIdx > Len(mTxBitmap) Then
        IsBitmapSFNAllowed = True
        Exit Function
    End If

    bitChar = Mid$(mTxBitmap, bitIdx, 1)
    IsBitmapSFNAllowed = (bitChar <> "0")
End Function

Private Sub WriteResolvedPoolToOutput()
    Dim i As Long
    Dim rowIdx As Long

    For i = 1 To mPoolCount
        rowIdx = mPoolRows(i)
        If rowIdx >= 1 And rowIdx <= mFilteredCount Then
            CopyRowToOutput rowIdx
            mWritten(rowIdx) = True
        End If
    Next i
End Sub

Private Sub WriteUnwrittenRowsToOutput()
    Dim r As Long

    For r = 1 To mFilteredCount
        If Not mWritten(r) Then
            CopyRowToOutput r
            mWritten(r) = True
        End If
    Next r
End Sub

Private Sub RecomputeFinalTXperSFN()
    Dim r As Long
    Dim startR As Long
    Dim endR As Long
    Dim curSFN As Long
    Dim runLen As Long
    Dim i As Long

    If mFilteredCount <= 0 Then Exit Sub

    r = 1
    Do While r <= mFilteredCount
        curSFN = mCurrentSFN(r)
        startR = r

        Do While r <= mFilteredCount
            If mCurrentSFN(r) <> curSFN Then Exit Do
            r = r + 1
        Loop

        endR = r - 1
        runLen = endR - startR + 1

        If mIdxTXperSFN > 0 Then
            For i = startR To endR
                mOutputData(i, mIdxTXperSFN) = runLen
            Next i
        End If
    Loop
End Sub

Private Sub ValidateResolvedRXTimingOnly()
    Dim r As Long
    Dim rxMin As Double

    mRemainingViolations = 0
    For r = 1 To mFilteredCount
        rxMin = mRowMinRxTime(r)
        If rxMin <> NO_RX_TIME Then
            If CDbl(mCurrentSFN(r)) > rxMin Then
                mRemainingViolations = mRemainingViolations + 1
            End If
        End If
    Next r
End Sub

Private Sub FinalizeOutputVariant(ByRef data As Variant)
    data = mOutputData
End Sub

Private Sub CopyRowToOutput(ByVal rowIdx As Long)
    Dim c As Long
    Dim colCount As Long

    If rowIdx < 1 Or rowIdx > mFilteredCount Then Exit Sub

    colCount = UBound(mOutputData, 2)
    For c = LBound(mData, 2) To colCount
        mOutputData(rowIdx, c) = mData(rowIdx, c)
    Next c

    mOutputData(rowIdx, mIdxSFNCol) = mCurrentSFN(rowIdx)
End Sub

Private Function BuildSubsetExcludingOne(ByRef rowListIn() As Long, ByVal rowCountIn As Long, ByVal removeRowIdx As Long, ByRef rowListOut() As Long) As Long
    BuildSubsetExcludingOne = 0
End Function

Private Sub Sort3RowsByMinRxTime(ByRef rowA As Long, ByRef rowB As Long, ByRef rowC As Long)
End Sub

Private Sub QuickSortLongs(ByRef arr() As Long, ByVal first As Long, ByVal last As Long)
End Sub

Private Sub SortRowIndexByCurrentSFN(ByRef arr() As Long, ByVal first As Long, ByVal last As Long)
End Sub

Private Function CompareRowOrder(ByVal rowA As Long, ByVal rowB As Long) As Long
    CompareRowOrder = Sgn(mCurrentSFN(rowA) - mCurrentSFN(rowB))
End Function

Private Sub UpdateStatusBar()
    Application.StatusBar = "TX_SFN conflict resolution running..."
End Sub

Private Sub AddDiag(ByVal eventType As String, ByVal v1 As String, ByVal v2 As String, ByVal v3 As String, ByVal v4 As String, ByVal msg As String)
End Sub

Private Sub HistAddLong(ByRef dictObj As Object, ByVal keyVal As Long)
End Sub

Private Sub DumpHistogram(ByVal ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long, ByVal titleText As String, ByRef dictObj As Object)
End Sub

Private Function SafeDiv(ByVal numerator As Double, ByVal denominator As Double) As Double
    If denominator = 0# Then
        SafeDiv = 0#
    Else
        SafeDiv = numerator / denominator
    End If
End Function

Private Sub WriteDiagnosticLog_TXSFNCR(ByVal totalRows As Long, ByVal calcTime As Double)
End Sub

Private Function MicroTimer_TXSFNCR() As Double
    Dim cyTicks As Currency
    Dim cyFreq As Currency

    If QueryPerformanceFrequency_TXSFNCR(cyFreq) <> 0 Then
        QueryPerformanceCounter_TXSFNCR cyTicks
        If cyFreq > 0 Then MicroTimer_TXSFNCR = cyTicks / cyFreq
    End If
End Function

Private Sub InitializeOutputBuffer()
    Dim colCount As Long

    If mFilteredCount <= 0 Then Exit Sub

    colCount = UBound(mData, 2)
    ReDim mOutputData(1 To mFilteredCount, 1 To colCount)
    mOutputCount = 0
End Sub

Private Function ExtractPoolRows(ByVal startIdx As Long, ByVal endIdx As Long) As Long()
    Dim rows() As Long
    Dim i As Long
    Dim n As Long

    n = endIdx - startIdx + 1
    ReDim rows(1 To n)

    For i = 1 To n
        rows(i) = mPoolRows(startIdx + i - 1)
    Next i

    ExtractPoolRows = rows
End Function

Private Function CollectRowsForSFN(ByVal sfnVal As Long, ByVal excludeRowIdx As Long, ByVal includeRowIdx As Long, ByRef rowCount As Long) As Long()
    Dim rows() As Long
    Dim r As Long

    ReDim rows(1 To mFilteredCount)
    rowCount = 0

    For r = 1 To mFilteredCount
        If r <> excludeRowIdx Then
            If mCurrentSFN(r) = sfnVal Or r = includeRowIdx Then
                rowCount = rowCount + 1
                rows(rowCount) = r
            End If
        End If
    Next r

    If rowCount = 0 Then
        ReDim rows(1 To 1)
    ElseIf rowCount < mFilteredCount Then
        ReDim Preserve rows(1 To rowCount)
    End If

    CollectRowsForSFN = rows
End Function

Private Function ValidateBucketRows(ByRef rows() As Long, ByVal rowCount As Long) As Boolean
    Dim t As Double

    t = MicroTimer_TXSFNCR()

    If rowCount <= 1 Then
        ValidateBucketRows = True
        mValidateBucketSeconds = mValidateBucketSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    If Not IsGroupTXIDUnique(rows, rowCount) Then
        mValidateBucketSeconds = mValidateBucketSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    If Not IsGroupMergedStationsUnique(rows, rowCount) Then
        mValidateBucketSeconds = mValidateBucketSeconds + (MicroTimer_TXSFNCR() - t)
        Exit Function
    End If

    ValidateBucketRows = True
    mValidateBucketSeconds = mValidateBucketSeconds + (MicroTimer_TXSFNCR() - t)
End Function

Private Function IsGroupTXIDUnique(ByRef rows() As Long, ByVal rowCount As Long) As Boolean
    Dim dictTx As Object
    Dim i As Long
    Dim rowIdx As Long
    Dim txIDKey As String

    Set dictTx = CreateObject("Scripting.Dictionary")

    For i = 1 To rowCount
        rowIdx = rows(i)
        txIDKey = CStr(mRowTXID(rowIdx))
        If dictTx.Exists(txIDKey) Then Exit Function
        dictTx(txIDKey) = 1
    Next i

    IsGroupTXIDUnique = True
End Function

Private Function IsGroupMergedStationsUnique(ByRef rows() As Long, ByVal rowCount As Long) As Boolean
    Dim dictStations As Object
    Dim i As Long
    Dim rowIdx As Long

    Set dictStations = CreateObject("Scripting.Dictionary")

    For i = 1 To rowCount
        rowIdx = rows(i)
        If Not AddRowStationsToDict(rowIdx, dictStations) Then Exit Function
    Next i

    IsGroupMergedStationsUnique = True
End Function

Private Function AddRowStationsToDict(ByVal rowIdx As Long, ByRef dictStations As Object) As Boolean
    Dim txKey As String
    Dim st As Long
    Dim stationId As Long
    Dim stationKey As String
    Dim stationUB As Long

    txKey = CStr(mRowTXID(rowIdx))
    If dictStations.Exists(txKey) Then Exit Function
    dictStations(txKey) = 1

    stationUB = SafeLongArrayUBound(mRxStationIDs)
    For st = 1 To mActiveRxCount
        If HasRowRxForStation(rowIdx, st) Then
            If stationUB > 0 And st <= stationUB Then
                stationId = mRxStationIDs(st)
            Else
                stationId = st
            End If

            stationKey = CStr(stationId)
            If dictStations.Exists(stationKey) Then Exit Function
            dictStations(stationKey) = 1
        End If
    Next st

    AddRowStationsToDict = True
End Function

Private Function HasRowRxForStation(ByVal rowIdx As Long, ByVal stationOrdinal As Long) As Boolean
    Dim colIdx As Long
    Dim dataUB As Long

    If stationOrdinal < 1 Or stationOrdinal > mActiveRxCount Then Exit Function

    dataUB = SafeLongArrayUBound(mRxDataColIdx)
    If dataUB <= 0 Then Exit Function
    If stationOrdinal > dataUB Then Exit Function

    colIdx = mRxDataColIdx(stationOrdinal)
    If colIdx <= 0 Then Exit Function
    If IsEmpty(mData(rowIdx, colIdx)) Then Exit Function

    HasRowRxForStation = (LenB(Trim$(CStr(mData(rowIdx, colIdx)))) > 0)
End Function

Private Function DoesMovedRowFormValidLocalGroup(ByVal movedRowIdx As Long) As Boolean
    Dim movedSFN As Long

    movedSFN = mCurrentSFN(movedRowIdx)

    If Not EvaluatePoolBucketWithAddedRow(movedSFN, movedRowIdx) Then
        If DEBUG_TXSFNCR Then Debug.Print "REJECT local-group", "movedRowIdx", movedRowIdx, "movedSFN", movedSFN
        Exit Function
    End If

    DoesMovedRowFormValidLocalGroup = True
End Function

Private Function GetMaxMoveOffset(ByVal sourceSFN As Long) As Long
    Dim lowerRoom As Long
    Dim bitmapLimit As Long
    Dim maxOffset As Long

    lowerRoom = sourceSFN
    bitmapLimit = mBitmapLen
    If bitmapLimit <= 0 Then bitmapLimit = 64

    maxOffset = bitmapLimit
    If lowerRoom < maxOffset Then maxOffset = lowerRoom
    If maxOffset < 1 Then maxOffset = 1

    If DEBUG_TXSFNCR Then
        Debug.Print "GetMaxMoveOffset", "sourceSFN", sourceSFN, "lowerRoom", lowerRoom, "bitmapLimit", bitmapLimit, "maxOffset", maxOffset
    End If

    GetMaxMoveOffset = maxOffset
End Function

Private Function GetRowMinRxTime(ByVal rowIdx As Long) As Double
    Dim st As Long
    Dim colIdx As Long
    Dim rxVal As Double
    Dim dataUB As Long
    Dim hasValue As Boolean

    GetRowMinRxTime = NO_RX_TIME
    dataUB = SafeLongArrayUBound(mRxDataColIdx)
    If dataUB <= 0 Then Exit Function

    For st = 1 To mActiveRxCount
        If st > dataUB Then Exit For
        colIdx = mRxDataColIdx(st)
        If colIdx > 0 Then
            If Not IsEmpty(mData(rowIdx, colIdx)) Then
                If IsNumeric(mData(rowIdx, colIdx)) Then
                    rxVal = CDbl(mData(rowIdx, colIdx))
                    If Not hasValue Then
                        hasValue = True
                        GetRowMinRxTime = rxVal
                    ElseIf rxVal < GetRowMinRxTime Then
                        GetRowMinRxTime = rxVal
                    End If
                End If
            End If
        End If
    Next st
End Function

Private Function SafeLongArrayUBound(ByRef arr() As Long) As Long
    On Error Resume Next
    SafeLongArrayUBound = UBound(arr)
    If Err.Number <> 0 Then
        SafeLongArrayUBound = 0
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function IsAllOnesBitmap(ByVal txBitmap As String) As Boolean
    Dim i As Long

    If Len(txBitmap) = 0 Then Exit Function

    For i = 1 To Len(txBitmap)
        If Mid$(txBitmap, i, 1) <> "1" Then Exit Function
    Next i

    IsAllOnesBitmap = True
End Function

Private Function GetNamedLong(ByVal nameText As String) As Long
    On Error Resume Next
    GetNamedLong = CLng(Evaluate(ThisWorkbook.Names(nameText).RefersTo))
    On Error GoTo 0
End Function

Private Sub UpdateProgressBar(ByVal currentRow As Long)
    Application.StatusBar = "Conflict Resolution at row " & _
                            Format$(currentRow, "0") & _
                            " of " & Format$(mFilteredCount, "0")
End Sub

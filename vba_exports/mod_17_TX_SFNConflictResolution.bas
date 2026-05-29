Attribute VB_Name = "mod_17_TX_SFNConflictResolution"
Option Explicit

Private Const MODULE_VERSION_TXSFNCR_V5 As String = "V5.1.0"
Private Const LOG_FIRST_COLUMN_TXSFNCR_V5 As String = "U"
Private Const DEFAULT_SEARCH_RADIUS_TXSFNCR_V5 As Long = 4096

#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter_TXSFNCR_V5 Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency_TXSFNCR_V5 Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef lpFrequency As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter_TXSFNCR_V5 Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency_TXSFNCR_V5 Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef lpFrequency As Currency) As Long
#End If

Private Type TCsetAnalysisV5
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
Private mMissingPduSizes As Object

Private mInitialSFN() As Long
Private mCurrentSFN() As Long
Private mRowTXID() As Long
Private mRowTXQTime() As Double
Private mRowNsch() As Long
Private mRowPduKey() As String
Private mRowMinRxTime() As Double
Private mRowMinRxStation() As Long
Private mRowOriginalIndex() As Long
Private mRowValidInput() As Boolean

Private mWritten() As Boolean

Private mPoolRows() As Long
Private mPoolCount As Long
Private mPoolBucketRows As Object
Private mPoolBucketCounts As Object
Private mPoolMembership As Object

Private mOutputData() As Variant
Private mOutputCount As Long

Private mNschPerSubfr As Long
Private mMaxPasses As Long
Private mMaxSearchRadius As Long
Private mEscapeForwardLimit As Long
Private mDebugLogging As Boolean
Private mMaxTXSfnDecrement As Double

Private mPassCount As Long
Private mNudgeCount As Long
Private mPoolCountResolved As Long
Private mMaxObservedPoolSize As Long
Private mRemainingViolations As Long
Private mUnresolvedAttemptCount As Long
Private mDiagCount As Long

Private mDiagLog() As Variant
Private mNudgeLog() As Variant

' Outer phase timing
Private mFindConflictSeconds As Double
Private mBuildPoolSeconds As Double
Private mResolvePoolSeconds As Double
Private mWritePoolSeconds As Double
Private mFinalizeSeconds As Double

' Deeper timing instrumentation
Private mPoolFindBucketSeconds As Double
Private mEvaluateRowSetSeconds As Double
Private mAnalyzeSingleMoveSeconds As Double
Private mEvalBucketExcludeSeconds As Double
Private mEvalBucketAddSeconds As Double
Private mPoolRemoveBucketSeconds As Double
Private mBuildSubsetSeconds As Double

' Status bar timing instrumentation
Private mStatusBarSeconds As Double
Private mStatusBarCalls As Long
Private mStatusBarActualUpdates As Long
Private mLastStatusUpdateTime As Double
Private Const STATUS_UPDATE_INTERVAL_SECONDS_V5 As Double = 1#

' Call counters
Private mPoolFindBucketCalls As Long
Private mEvaluateRowSetCalls As Long
Private mAnalyzeSingleMoveCalls As Long
Private mEvalBucketExcludeCalls As Long
Private mEvalBucketAddCalls As Long
Private mPoolRemoveBucketCalls As Long
Private mBuildSubsetCalls As Long

' Scanner/group counters
Private mScannerGroupsFound As Long
Private mScannerGroupsValidated As Long

Private mOneMoveResolutionCount As Long
Private mTripleSplitResolutionCount As Long
Private mForwardEscapeCount As Long
Private mForwardEscapeDistanceTotal As Double
Private mForwardEscapeDistanceMax As Long

Private mPoolSizeHist As Object
Private mCsetOrderHist As Object
Private mPassesPerPoolHist As Object

Public Sub Run_TX_SFNConflictResolution()
    MsgBox "Run_TX_SFNConflictResolution is a wrapper. Call TX_SFNConflictResolution from PickExp.", _
           vbInformation, "TX_SFN Conflict Resolution " & MODULE_VERSION_TXSFNCR_V5
End Sub

Public Sub TX_SFNConflictResolution( _
    ByRef data As Variant, _
    ByVal filteredCount As Long, _
    ByVal idxSFNCol As Long, _
    ByVal idxTXID As Long, _
    ByVal idxTXQ As Long, _
    ByVal idxLEN As Long, _
    ByVal idxTXperSFN As Long, _
    ByVal idxRxCnt As Long, _
    ByVal idxAvg As Long, _
    ByVal idxTotLat As Long, _
    ByVal idxGen As Long, _
    ByRef rxDataColIdx() As Long, _
    ByRef rxStationIDs() As Long, _
    ByVal activeRxCount As Long, _
    ByRef dictS2V As Object, _
    ByRef dictVC As Object, _
    ByRef dictA2P As Object, _
    ByRef dictP2R As Object, _
    ByRef dictP2Sigma As Object, _
    ByVal txBitmap As String, _
    ByVal bitmapLen As Long, _
    ByRef elapsedSeconds As Double)

    Dim startTime As Double

    startTime = MicroTimer_TXSFNCR_V5()

    InitializeContext_V5 data, filteredCount, idxSFNCol, idxTXID, idxTXQ, idxLEN, _
                         idxTXperSFN, idxRxCnt, idxAvg, idxTotLat, idxGen, _
                         rxDataColIdx, rxStationIDs, activeRxCount, _
                         dictS2V, dictVC, dictA2P, dictP2R, dictP2Sigma, _
                         txBitmap, bitmapLen

    If mFilteredCount <= 0 Then
        elapsedSeconds = MicroTimer_TXSFNCR_V5() - startTime
        Exit Sub
    End If

    If Not ValidateInputMonotoneTXQTime_V5() Then
        elapsedSeconds = MicroTimer_TXSFNCR_V5() - startTime
        Exit Sub
    End If

    PrepareRowDerivedData_V5
    InitializeOutputBuffer_V5
    ResolveRowsGreedy_V5
    RecomputeFinalTXperSFN_V5
    FinalizeOutputVariant_V5 data

    elapsedSeconds = MicroTimer_TXSFNCR_V5() - startTime
    Application.StatusBar = False
    WriteDiagnosticLog_TXSFNCR_V5 mFilteredCount, elapsedSeconds
End Sub

Private Sub InitializeContext_V5( _
    ByRef data As Variant, _
    ByVal filteredCount As Long, _
    ByVal idxSFNCol As Long, _
    ByVal idxTXID As Long, _
    ByVal idxTXQ As Long, _
    ByVal idxLEN As Long, _
    ByVal idxTXperSFN As Long, _
    ByVal idxRxCnt As Long, _
    ByVal idxAvg As Long, _
    ByVal idxTotLat As Long, _
    ByVal idxGen As Long, _
    ByRef rxDataColIdx() As Long, _
    ByRef rxStationIDs() As Long, _
    ByVal activeRxCount As Long, _
    ByRef dictS2V As Object, _
    ByRef dictVC As Object, _
    ByRef dictA2P As Object, _
    ByRef dictP2R As Object, _
    ByRef dictP2Sigma As Object, _
    ByVal txBitmap As String, _
    ByVal bitmapLen As Long)

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

    mActiveRxCount = activeRxCount
    If mActiveRxCount > 0 Then
        ReDim mRxDataColIdx(1 To mActiveRxCount)
        ReDim mRxStationIDs(1 To mActiveRxCount)
        Dim i As Long
        For i = 1 To mActiveRxCount
            mRxDataColIdx(i) = rxDataColIdx(i)
            mRxStationIDs(i) = rxStationIDs(i)
        Next i
    End If

    Set mDictS2V = dictS2V
    Set mDictVC = dictVC
    Set mDictA2P = dictA2P
    Set mDictP2R = dictP2R
    Set mDictP2Sigma = dictP2Sigma

    Set mMissingPduSizes = CreateObject("Scripting.Dictionary")
    Set mPoolSizeHist = CreateObject("Scripting.Dictionary")
    Set mCsetOrderHist = CreateObject("Scripting.Dictionary")
    Set mPassesPerPoolHist = CreateObject("Scripting.Dictionary")

    mTxBitmap = Trim$(txBitmap)
    mBitmapLen = bitmapLen
    If mBitmapLen <= 0 Then mBitmapLen = Len(mTxBitmap)
    mNschPerSubfr = GetWorkbookNameLongSafe_TXSFNCR_V5("Nsch_per_subfr", 999999)
    mMaxTXSfnDecrement = GetWorkbookNameDoubleSafe_TXSFNCR_V5("maxTX_SFN_est_decrement", 0#)
    mMaxPasses = 1
    mMaxSearchRadius = DEFAULT_SEARCH_RADIUS_TXSFNCR_V5
    If mBitmapLen > 0 Then
        If mBitmapLen * 16 > mMaxSearchRadius Then mMaxSearchRadius = mBitmapLen * 16
    End If
    mEscapeForwardLimit = mMaxSearchRadius
    mLastStatusUpdateTime = 0#
End Sub

Private Function ValidateInputMonotoneTXQTime_V5() As Boolean
    ValidateInputMonotoneTXQTime_V5 = (Not IsEmpty(mData))
End Function

Private Sub PrepareRowDerivedData_V5()
    Dim r As Long
    Dim rawLenKey As String
    Dim mapVal As Variant
    Dim i As Long
    Dim rxTimeVal As Double
    Dim minRxTime As Double
    Dim minRxStation As Long
    Dim hasMinRx As Boolean

    ReDim mInitialSFN(1 To mFilteredCount)
    ReDim mCurrentSFN(1 To mFilteredCount)
    ReDim mRowTXID(1 To mFilteredCount)
    ReDim mRowTXQTime(1 To mFilteredCount)
    ReDim mRowNsch(1 To mFilteredCount)
    ReDim mRowPduKey(1 To mFilteredCount)
    ReDim mRowMinRxTime(1 To mFilteredCount)
    ReDim mRowMinRxStation(1 To mFilteredCount)
    ReDim mRowOriginalIndex(1 To mFilteredCount)
    ReDim mRowValidInput(1 To mFilteredCount)

    For r = 1 To mFilteredCount
        mRowOriginalIndex(r) = r

        If IsNumeric(mData(r, mIdxSFNCol)) Then
            mInitialSFN(r) = CLng(mData(r, mIdxSFNCol))
            mCurrentSFN(r) = mInitialSFN(r)
        End If

        If IsNumeric(mData(r, mIdxTXID)) Then
            mRowTXID(r) = CLng(mData(r, mIdxTXID))
        End If

        If IsNumeric(mData(r, mIdxTXQ)) Then
            mRowTXQTime(r) = CDbl(mData(r, mIdxTXQ))
        End If

        rawLenKey = Trim$(CStr(mData(r, mIdxLEN)))
        mRowPduKey(r) = rawLenKey
        If rawLenKey <> "" Then
            If Not mDictA2P Is Nothing Then
                If mDictA2P.Exists(rawLenKey) Then
                    mapVal = mDictA2P(rawLenKey)
                    If IsArray(mapVal) Then
                        If IsNumeric(mapVal(0)) Then mRowNsch(r) = CLng(mapVal(0))
                        mRowPduKey(r) = Trim$(CStr(mapVal(1)))
                    ElseIf IsNumeric(mapVal) Then
                        mRowNsch(r) = CLng(mapVal)
                    End If
                Else
                    mMissingPduSizes(rawLenKey) = True
                End If
            End If
        End If

        hasMinRx = False
        For i = 1 To mActiveRxCount
            If IsNumeric(mData(r, mRxDataColIdx(i))) Then
                rxTimeVal = CDbl(mData(r, mRxDataColIdx(i)))
                If rxTimeVal = 0 Then GoTo NextRxTime
                If Not hasMinRx Or rxTimeVal < minRxTime Then
                    hasMinRx = True
                    minRxTime = rxTimeVal
                    minRxStation = mRxStationIDs(i)
                End If
            End If
NextRxTime:
        Next i

        If hasMinRx Then
            mRowMinRxTime(r) = minRxTime
            mRowMinRxStation(r) = minRxStation
        End If

        mRowValidInput(r) = (mRowTXID(r) <> 0 And IsNumeric(mData(r, mIdxTXQ)) And mCurrentSFN(r) <> 0)
    Next r
End Sub

Private Sub ResolveRowsGreedy_V5()
    Dim rowOrder() As Long
    Dim rowCount As Long
    Dim r As Long
    Dim rowIdx As Long
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim preferredSFN As Long
    Dim resolvedSFN As Long
    Dim bucketMap As Object
    Dim t0 As Double

    rowCount = 0
    ReDim rowOrder(1 To mFilteredCount)

    For r = 1 To mFilteredCount
        If mRowValidInput(r) Then
            rowCount = rowCount + 1
            rowOrder(rowCount) = r
        End If
    Next r

    If rowCount = 0 Then Exit Sub
    ReDim Preserve rowOrder(1 To rowCount)
    SortRowIndexByCurrentSFN_V5 rowOrder, 1, rowCount

    Set bucketMap = CreateObject("Scripting.Dictionary")

    For r = 1 To rowCount
        rowIdx = rowOrder(r)
        preferredSFN = mInitialSFN(rowIdx)
        lowerBound = GetRowLowerBound_V5(rowIdx)
        upperBound = GetRowUpperBound_V5(rowIdx)

        t0 = MicroTimer_TXSFNCR_V5()
        resolvedSFN = FindNearestLegalSFN_V5(rowIdx, lowerBound, upperBound, preferredSFN, bucketMap)
        mFindConflictSeconds = mFindConflictSeconds + (MicroTimer_TXSFNCR_V5() - t0)

        If resolvedSFN = 0 Then
            resolvedSFN = ChooseFallbackSFN_V5(lowerBound, upperBound, preferredSFN)
            mUnresolvedAttemptCount = mUnresolvedAttemptCount + 1
        ElseIf resolvedSFN <> preferredSFN Then
            mNudgeCount = mNudgeCount + 1
        End If

        t0 = MicroTimer_TXSFNCR_V5()
        mCurrentSFN(rowIdx) = resolvedSFN
        RecordAssignment_V5 bucketMap, rowIdx, resolvedSFN
        mPoolCountResolved = r
        mBuildPoolSeconds = mBuildPoolSeconds + (MicroTimer_TXSFNCR_V5() - t0)

        UpdateStatusBar_V5
    Next r

    mPoolCountResolved = rowCount - mUnresolvedAttemptCount
End Sub

Private Function FindNearestLegalSFN_V5(ByVal rowIdx As Long, ByVal lowerBound As Long, ByVal upperBound As Long, ByVal preferredSFN As Long, ByVal bucketMap As Object) As Long
    Dim offset As Long
    Dim backwardSFN As Long
    Dim forwardSFN As Long
    Dim searchSpan As Long

    If upperBound < lowerBound Then Exit Function

    searchSpan = upperBound - lowerBound
    If searchSpan < 0 Then Exit Function

    For offset = 0 To searchSpan
        backwardSFN = preferredSFN - offset
        If backwardSFN >= lowerBound Then
            If IsCandidateLegal_V5(rowIdx, backwardSFN, bucketMap) Then
                FindNearestLegalSFN_V5 = backwardSFN
                Exit Function
            End If
        End If

        If offset > 0 Then
            forwardSFN = preferredSFN + offset
            If forwardSFN <= upperBound Then
                If IsCandidateLegal_V5(rowIdx, forwardSFN, bucketMap) Then
                    FindNearestLegalSFN_V5 = forwardSFN
                    Exit Function
                End If
            End If
        End If
    Next offset
End Function

Private Function ChooseFallbackSFN_V5(ByVal lowerBound As Long, ByVal upperBound As Long, ByVal preferredSFN As Long) As Long
    Dim candidateSFN As Long

    candidateSFN = preferredSFN
    If candidateSFN < lowerBound Then candidateSFN = lowerBound
    If candidateSFN > upperBound Then candidateSFN = upperBound
    If candidateSFN <= 0 Then candidateSFN = 1

    If IsBitmapSFNAllowed_V5(candidateSFN) Then
        ChooseFallbackSFN_V5 = candidateSFN
        Exit Function
    End If

    For candidateSFN = lowerBound To upperBound
        If IsBitmapSFNAllowed_V5(candidateSFN) Then
            ChooseFallbackSFN_V5 = candidateSFN
            Exit Function
        End If
    Next candidateSFN

    ChooseFallbackSFN_V5 = preferredSFN
    If ChooseFallbackSFN_V5 <= 0 Then ChooseFallbackSFN_V5 = 1
End Function

Private Function GetRowLowerBound_V5(ByVal rowIdx As Long) As Long
    Dim txProc As Double

    txProc = GetTxProcMean_V5(mRowTXID(rowIdx))
    GetRowLowerBound_V5 = CLng(Round(mRowTXQTime(rowIdx) + txProc, 0) - mMaxTXSfnDecrement)
    If GetRowLowerBound_V5 < 1 Then GetRowLowerBound_V5 = 1
End Function

Private Function GetRowUpperBound_V5(ByVal rowIdx As Long) As Long
    Dim rxProc As Double
    Dim rxSigma As Double
    Dim upperCandidate As Long

    If mRowMinRxStation(rowIdx) = 0 Then
        GetRowUpperBound_V5 = mInitialSFN(rowIdx) + mMaxSearchRadius
        If GetRowUpperBound_V5 < GetRowLowerBound_V5(rowIdx) Then
            GetRowUpperBound_V5 = GetRowLowerBound_V5(rowIdx)
        End If
        Exit Function
    End If

    rxProc = GetRxProcMean_V5(mRowMinRxStation(rowIdx), mRowPduKey(rowIdx))
    rxSigma = GetRxProcSigma_V5(mRowMinRxStation(rowIdx), mRowPduKey(rowIdx))
    upperCandidate = CLng(Fix(mRowMinRxTime(rowIdx) - rxProc + (3# * rxSigma)))
    If upperCandidate < GetRowLowerBound_V5(rowIdx) Then upperCandidate = GetRowLowerBound_V5(rowIdx)
    GetRowUpperBound_V5 = upperCandidate
End Function

Private Function GetTxProcMean_V5(ByVal txStationID As Long) As Double
    Dim vendorKey As String
    Dim txKey As String
    Dim txVals As Variant

    txKey = CStr(txStationID)
    vendorKey = ResolveVendorKey_V5(txKey)
    If vendorKey = "" Then Exit Function

    If mDictVC Is Nothing Then Exit Function
    If mDictVC.Exists(vendorKey) Then
        txVals = mDictVC(vendorKey)
        If IsArray(txVals) Then
            If IsNumeric(txVals(0)) Then GetTxProcMean_V5 = CDbl(txVals(0))
        End If
    End If
End Function

Private Function GetRxProcMean_V5(ByVal rxStationID As Long, ByVal pduKey As String) As Double
    Dim vendorKey As String
    Dim dictKey As String

    If mDictP2R Is Nothing Then Exit Function
    vendorKey = ResolveVendorKey_V5(CStr(rxStationID))
    If vendorKey = "" Then Exit Function

    dictKey = pduKey & "|" & vendorKey
    If mDictP2R.Exists(dictKey) Then
        If IsNumeric(mDictP2R(dictKey)) Then GetRxProcMean_V5 = CDbl(mDictP2R(dictKey))
    End If
End Function

Private Function GetRxProcSigma_V5(ByVal rxStationID As Long, ByVal pduKey As String) As Double
    Dim vendorKey As String
    Dim dictKey As String

    If mDictP2Sigma Is Nothing Then Exit Function
    vendorKey = ResolveVendorKey_V5(CStr(rxStationID))
    If vendorKey = "" Then Exit Function

    dictKey = pduKey & "|" & vendorKey
    If mDictP2Sigma.Exists(dictKey) Then
        If IsNumeric(mDictP2Sigma(dictKey)) Then GetRxProcSigma_V5 = CDbl(mDictP2Sigma(dictKey))
    End If
End Function

Private Function ResolveVendorKey_V5(ByVal stationKey As String) As String
    Dim lookupKey As String

    If mDictS2V Is Nothing Then Exit Function
    lookupKey = UCase$(Trim$(stationKey))
    If mDictS2V.Exists(lookupKey) Then
        ResolveVendorKey_V5 = Trim$(CStr(mDictS2V(lookupKey)))
    ElseIf mDictS2V.Exists(Trim$(stationKey)) Then
        ResolveVendorKey_V5 = Trim$(CStr(mDictS2V(Trim$(stationKey))))
    End If
End Function

Private Function IsCandidateLegal_V5(ByVal rowIdx As Long, ByVal testSFN As Long, ByVal bucketMap As Object) As Boolean
    Dim bucket As Object
    Dim txDict As Object
    Dim rxDict As Object
    Dim capDict As Object
    Dim i As Long
    Dim stationKey As String
    Dim currentCap As Long

    If testSFN <= 0 Then Exit Function
    If Not IsBitmapSFNAllowed_V5(testSFN) Then Exit Function

    If Not bucketMap.Exists(CStr(testSFN)) Then
        IsCandidateLegal_V5 = IsRowSelfConsistent_V5(rowIdx)
        Exit Function
    End If

    Set bucket = bucketMap(CStr(testSFN))
    Set txDict = bucket("tx")
    Set rxDict = bucket("rx")
    Set capDict = bucket("cap")

    If txDict.Exists(CStr(mRowTXID(rowIdx))) Then Exit Function
    If rxDict.Exists(CStr(mRowTXID(rowIdx))) Then Exit Function
    If Not IsRowSelfConsistent_V5(rowIdx) Then Exit Function

    For i = 1 To mActiveRxCount
        If IsRowReceivedByStation_V5(rowIdx, i) Then
            stationKey = CStr(mRxStationIDs(i))
            If txDict.Exists(stationKey) Then Exit Function

            currentCap = 0
            If capDict.Exists(stationKey) Then currentCap = CLng(capDict(stationKey))
            If currentCap + mRowNsch(rowIdx) > mNschPerSubfr Then Exit Function
        End If
    Next i

    IsCandidateLegal_V5 = True
End Function

Private Function IsRowSelfConsistent_V5(ByVal rowIdx As Long) As Boolean
    Dim i As Long

    If mRowNsch(rowIdx) > mNschPerSubfr Then Exit Function

    For i = 1 To mActiveRxCount
        If IsRowReceivedByStation_V5(rowIdx, i) Then
            If mRxStationIDs(i) = mRowTXID(rowIdx) Then Exit Function
        End If
    Next i

    IsRowSelfConsistent_V5 = True
End Function

Private Function IsRowReceivedByStation_V5(ByVal rowIdx As Long, ByVal rxIndex As Long) As Boolean
    If rxIndex < 1 Or rxIndex > mActiveRxCount Then Exit Function
    If IsNumeric(mData(rowIdx, mRxDataColIdx(rxIndex))) Then
        IsRowReceivedByStation_V5 = (CDbl(mData(rowIdx, mRxDataColIdx(rxIndex))) <> 0)
    End If
End Function

Private Sub RecordAssignment_V5(ByVal bucketMap As Object, ByVal rowIdx As Long, ByVal sfnVal As Long)
    Dim bucket As Object
    Dim txDict As Object
    Dim rxDict As Object
    Dim capDict As Object
    Dim i As Long
    Dim stationKey As String
    Dim updatedCap As Long

    If Not bucketMap.Exists(CStr(sfnVal)) Then
        Set bucket = CreateObject("Scripting.Dictionary")
        Set txDict = CreateObject("Scripting.Dictionary")
        Set rxDict = CreateObject("Scripting.Dictionary")
        Set capDict = CreateObject("Scripting.Dictionary")
        bucket("tx") = txDict
        bucket("rx") = rxDict
        bucket("cap") = capDict
        bucketMap(CStr(sfnVal)) = bucket
    Else
        Set bucket = bucketMap(CStr(sfnVal))
        Set txDict = bucket("tx")
        Set rxDict = bucket("rx")
        Set capDict = bucket("cap")
    End If

    txDict(CStr(mRowTXID(rowIdx))) = True

    For i = 1 To mActiveRxCount
        If IsRowReceivedByStation_V5(rowIdx, i) Then
            stationKey = CStr(mRxStationIDs(i))
            rxDict(stationKey) = True
            updatedCap = mRowNsch(rowIdx)
            If capDict.Exists(stationKey) Then updatedCap = updatedCap + CLng(capDict(stationKey))
            capDict(stationKey) = updatedCap
        End If
    Next i
End Sub

Private Sub InitializeOutputBuffer_V5()
    If IsEmpty(mData) Then Exit Sub
    mOutputData = mData
    mOutputCount = mFilteredCount
End Sub

Private Function FindNextConflictStart_V5() As Long
End Function

Private Sub BuildPoolFromConflictStart_V5(ByVal startRow As Long)
End Sub

Private Sub ResolveEntirePool_V5()
End Sub

Private Function TryResolveCset_V5(ByRef rows() As Long, ByVal rowCount As Long, ByVal sourceSFN As Long) As Boolean
End Function

Private Function AnalyzeCsetSingleMoves_V5(ByRef rows() As Long, ByVal rowCount As Long) As TCsetAnalysisV5
End Function

Private Function TryPlaceOneMovedRow_NoSourceRetest_V5(ByRef candidateRows() As Long, ByVal candidateCount As Long, ByVal sourceSFN As Long) As Boolean
End Function

Private Function IsOneMovedRowPlacementLegal_NoSourceRetest_V5(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean
End Function

Private Function TryForwardEscapeMove_NoSourceRetest_V5(ByVal sourceSFN As Long, ByVal rowIdx As Long) As Boolean
End Function

Private Function ResolveTripleSplitDeterministic_V5(ByRef rows() As Long, ByVal sourceSFN As Long) As Boolean
End Function

Private Function IsPoolForcedSplitFirstMoveLegal_V5(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean
End Function

Private Function IsPoolSingleRowPlacementLegal_V5(ByVal rowIdx As Long, ByVal sourceSFN As Long, ByVal testSFN As Long) As Boolean
End Function

Private Function EvaluatePoolBucketExcludingRow_V5(ByVal sfnVal As Long, ByVal excludeRowIdx As Long) As Boolean
End Function

Private Function EvaluatePoolBucketWithAddedRow_V5(ByVal sfnVal As Long, ByVal addedRowIdx As Long) As Boolean
End Function

Private Function IsMoveWithinRowBounds_V5(ByVal rowIdx As Long, ByVal testSFN As Long) As Boolean
    IsMoveWithinRowBounds_V5 = (testSFN >= GetRowLowerBound_V5(rowIdx) And testSFN <= GetRowUpperBound_V5(rowIdx))
End Function

Private Function IsBitmapSFNAllowed_V5(ByVal testSFN As Long) As Boolean
    Dim posIdx As Long

    If mBitmapLen <= 0 Or Len(mTxBitmap) = 0 Then
        IsBitmapSFNAllowed_V5 = True
        Exit Function
    End If

    posIdx = ((testSFN - 1) Mod mBitmapLen) + 1
    IsBitmapSFNAllowed_V5 = (Mid$(mTxBitmap, posIdx, 1) <> "0")
End Function

Private Sub WriteResolvedPoolToOutput_V5()
End Sub

Private Sub WriteUnwrittenRowsToOutput_V5()
End Sub

Private Sub RecomputeFinalTXperSFN_V5()
    Dim sfnCounts As Object
    Dim r As Long
    Dim sfnKey As String

    Set sfnCounts = CreateObject("Scripting.Dictionary")

    For r = 1 To mFilteredCount
        If mCurrentSFN(r) <> 0 Then
            sfnKey = CStr(mCurrentSFN(r))
            If sfnCounts.Exists(sfnKey) Then
                sfnCounts(sfnKey) = CLng(sfnCounts(sfnKey)) + 1
            Else
                sfnCounts(sfnKey) = 1
            End If
        End If
    Next r

    If mIdxTXperSFN <= 0 Then Exit Sub
    For r = 1 To mFilteredCount
        If mCurrentSFN(r) <> 0 Then
            mData(r, mIdxTXperSFN) = CLng(sfnCounts(CStr(mCurrentSFN(r))))
        End If
    Next r
End Sub

Private Sub FinalizeOutputVariant_V5(ByRef data As Variant)
    Dim r As Long

    For r = 1 To mFilteredCount
        If mCurrentSFN(r) <> 0 Then
            mData(r, mIdxSFNCol) = mCurrentSFN(r)
        End If
    Next r

    data = mData
End Sub

Private Function BuildSubsetExcludingOne_V5(ByRef rowListIn() As Long, ByVal rowCountIn As Long, ByVal removeRowIdx As Long, ByRef rowListOut() As Long) As Long
End Function

Private Sub Sort3RowsByMinRxTime_V5(ByRef rowA As Long, ByRef rowB As Long, ByRef rowC As Long)
End Sub

Private Sub QuickSortLongs_V5(ByRef arr() As Long, ByVal first As Long, ByVal last As Long)
    Dim low As Long
    Dim high As Long
    Dim pivot As Long
    Dim tempVal As Long

    low = first
    high = last
    pivot = arr((first + last) \ 2)

    Do While low <= high
        Do While arr(low) < pivot
            low = low + 1
        Loop
        Do While arr(high) > pivot
            high = high - 1
        Loop
        If low <= high Then
            tempVal = arr(low)
            arr(low) = arr(high)
            arr(high) = tempVal
            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then QuickSortLongs_V5 arr, first, high
    If low < last Then QuickSortLongs_V5 arr, low, last
End Sub

Private Sub SortRowIndexByCurrentSFN_V5(ByRef arr() As Long, ByVal first As Long, ByVal last As Long)
    Dim low As Long
    Dim high As Long
    Dim pivot As Long
    Dim tempVal As Long

    low = first
    high = last
    pivot = arr((first + last) \ 2)

    Do While low <= high
        Do While CompareRowOrder_V5(arr(low), pivot) < 0
            low = low + 1
        Loop
        Do While CompareRowOrder_V5(arr(high), pivot) > 0
            high = high - 1
        Loop
        If low <= high Then
            tempVal = arr(low)
            arr(low) = arr(high)
            arr(high) = tempVal
            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then SortRowIndexByCurrentSFN_V5 arr, first, high
    If low < last Then SortRowIndexByCurrentSFN_V5 arr, low, last
End Sub

Private Function CompareRowOrder_V5(ByVal rowA As Long, ByVal rowB As Long) As Long
    If mInitialSFN(rowA) < mInitialSFN(rowB) Then
        CompareRowOrder_V5 = -1
    ElseIf mInitialSFN(rowA) > mInitialSFN(rowB) Then
        CompareRowOrder_V5 = 1
    ElseIf mRowTXQTime(rowA) < mRowTXQTime(rowB) Then
        CompareRowOrder_V5 = -1
    ElseIf mRowTXQTime(rowA) > mRowTXQTime(rowB) Then
        CompareRowOrder_V5 = 1
    ElseIf mRowOriginalIndex(rowA) < mRowOriginalIndex(rowB) Then
        CompareRowOrder_V5 = -1
    ElseIf mRowOriginalIndex(rowA) > mRowOriginalIndex(rowB) Then
        CompareRowOrder_V5 = 1
    End If
End Function

Private Sub UpdateStatusBar_V5()
    Dim nowT As Double

    mStatusBarCalls = mStatusBarCalls + 1
    nowT = MicroTimer_TXSFNCR_V5()
    If (nowT - mLastStatusUpdateTime) < STATUS_UPDATE_INTERVAL_SECONDS_V5 Then Exit Sub

    Application.StatusBar = "TX_SFN conflict resolution: " & _
                            Format(mPoolCountResolved + mUnresolvedAttemptCount, "#,##0") & "/" & _
                            Format(mFilteredCount, "#,##0") & " rows processed"
    mLastStatusUpdateTime = nowT
    mStatusBarActualUpdates = mStatusBarActualUpdates + 1
End Sub

Private Sub AddDiag_V5(ByVal eventType As String, ByVal v1 As String, ByVal v2 As String, ByVal v3 As String, ByVal v4 As String, ByVal msg As String)
End Sub

Private Sub HistAddLong_V5(ByRef dictObj As Object, ByVal keyVal As Long)
    If dictObj Is Nothing Then Exit Sub
    If dictObj.Exists(CStr(keyVal)) Then
        dictObj(CStr(keyVal)) = CLng(dictObj(CStr(keyVal))) + 1
    Else
        dictObj(CStr(keyVal)) = 1
    End If
End Sub

Private Sub DumpHistogram_V5(ByVal ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long, ByVal titleText As String, ByRef dictObj As Object)
End Sub

Private Function SafeDiv_V5(ByVal numerator As Double, ByVal denominator As Double) As Double
    If denominator <> 0 Then SafeDiv_V5 = numerator / denominator
End Function

Private Sub WriteDiagnosticLog_TXSFNCR_V5(ByVal totalRows As Long, ByVal calcTime As Double)
    Dim ws As Worksheet
    Dim writeCol As Long
    Dim writeRow As Long
    Dim missingKey As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("TX_SFN est Log")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    writeCol = ws.Range(LOG_FIRST_COLUMN_TXSFNCR_V5 & "1").Column
    ws.Range(ws.Cells(2, writeCol), ws.Cells(200, writeCol + 4)).ClearContents

    writeRow = 2
    ws.Cells(writeRow, writeCol).Value = "TX_SFN CONFLICT RESOLUTION"
    ws.Cells(writeRow, writeCol).Font.Bold = True
    writeRow = writeRow + 1

    ws.Cells(writeRow, writeCol).Value = "Version"
    ws.Cells(writeRow, writeCol + 1).Value = MODULE_VERSION_TXSFNCR_V5
    writeRow = writeRow + 1

    ws.Cells(writeRow, writeCol).Value = "Rows processed"
    ws.Cells(writeRow, writeCol + 1).Value = totalRows
    writeRow = writeRow + 1

    ws.Cells(writeRow, writeCol).Value = "Rows moved"
    ws.Cells(writeRow, writeCol + 1).Value = mNudgeCount
    writeRow = writeRow + 1

    ws.Cells(writeRow, writeCol).Value = "Unresolved rows"
    ws.Cells(writeRow, writeCol + 1).Value = mUnresolvedAttemptCount
    writeRow = writeRow + 1

    ws.Cells(writeRow, writeCol).Value = "Runtime (s)"
    ws.Cells(writeRow, writeCol + 1).Value = Round(calcTime, 3)
    writeRow = writeRow + 1

    ws.Cells(writeRow, writeCol).Value = "Bitmap length"
    ws.Cells(writeRow, writeCol + 1).Value = mBitmapLen
    writeRow = writeRow + 1

    ws.Cells(writeRow, writeCol).Value = "Nsch/subframe"
    ws.Cells(writeRow, writeCol + 1).Value = mNschPerSubfr
    writeRow = writeRow + 2

    ws.Cells(writeRow, writeCol).Value = "Missing LEN mappings"
    ws.Cells(writeRow, writeCol).Font.Bold = True
    writeRow = writeRow + 1

    If mMissingPduSizes.Count = 0 Then
        ws.Cells(writeRow, writeCol).Value = "(none)"
    Else
        For Each missingKey In mMissingPduSizes.Keys
            ws.Cells(writeRow, writeCol).Value = CStr(missingKey)
            writeRow = writeRow + 1
        Next missingKey
    End If

    ws.Columns(writeCol).Resize(, 2).AutoFit
End Sub

Private Function GetWorkbookNameLongSafe_TXSFNCR_V5(ByVal nameText As String, ByVal defaultValue As Long) As Long
    Dim nm As Name
    Dim expr As String
    Dim v As Variant

    On Error GoTo UseDefault
    Set nm = ThisWorkbook.Names(nameText)
    expr = nm.RefersTo
    If Len(expr) > 0 Then
        If Left$(expr, 1) = "=" Then expr = Mid$(expr, 2)
    End If
    v = Application.Evaluate(expr)
    If IsError(v) Or Not IsNumeric(v) Then GoTo UseDefault
    GetWorkbookNameLongSafe_TXSFNCR_V5 = CLng(v)
    Exit Function

UseDefault:
    GetWorkbookNameLongSafe_TXSFNCR_V5 = defaultValue
End Function

Private Function GetWorkbookNameDoubleSafe_TXSFNCR_V5(ByVal nameText As String, ByVal defaultValue As Double) As Double
    Dim nm As Name
    Dim expr As String
    Dim v As Variant

    On Error GoTo UseDefault
    Set nm = ThisWorkbook.Names(nameText)
    expr = nm.RefersTo
    If Len(expr) > 0 Then
        If Left$(expr, 1) = "=" Then expr = Mid$(expr, 2)
    End If
    v = Application.Evaluate(expr)
    If IsError(v) Or Not IsNumeric(v) Then GoTo UseDefault
    GetWorkbookNameDoubleSafe_TXSFNCR_V5 = CDbl(v)
    Exit Function

UseDefault:
    GetWorkbookNameDoubleSafe_TXSFNCR_V5 = defaultValue
End Function

Private Function MicroTimer_TXSFNCR_V5() As Double
    Dim cyTicks As Currency
    Dim cyFreq As Currency

    If QueryPerformanceFrequency_TXSFNCR_V5(cyFreq) <> 0 Then
        QueryPerformanceCounter_TXSFNCR_V5 cyTicks
        If cyFreq > 0 Then
            MicroTimer_TXSFNCR_V5 = cyTicks / cyFreq
        End If
    End If
End Function

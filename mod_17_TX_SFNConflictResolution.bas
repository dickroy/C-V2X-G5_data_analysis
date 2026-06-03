Option Explicit

' Version: V1.0.3
' V1.0.3 changes:
'   - Fix CSet ordering runtime error 9 in inner-most selection
'   - Preserve pool-aware escape resolution behavior
Private Const MODULE_VERSION_TXSFNCR As String = "V1.0.3"
Private Const DEBUG_TXSFNCR As Boolean = True
Private Const PROGRESS_STEP_ROWS As Long = 100

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

Private Const NO_RX_TIME As Double = -1#

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

Public Sub Reset_TX_SFNConflictResolution_State()
    Set mDictS2V = Nothing
    Set mDictVC = Nothing
    Set mDictA2P = Nothing
    Set mDictP2R = Nothing
    Set mDictP2Sigma = Nothing
    Set mMissingPduSizes = Nothing

    mData = Empty
    mFilteredCount = 0
    mIdxTXID = 0
    mIdxTXQ = 0
    mIdxSFNCol = 0
    mIdxLEN = 0
    mIdxTXperSFN = 0
    mIdxGen = 0
    mIdxAvg = 0
    mIdxTotLat = 0
    mIdxRxCnt = 0
    mTxBitmap = vbNullString
    mBitmapLen = 0
    mActiveRxCount = 0

    Erase mRxStationIDs
    Erase mRxDataColIdx
    Erase mInitialSFN
    Erase mCurrentSFN
    Erase mRowTXID
    Erase mRowTXQTime
    Erase mRowNsch
    Erase mRowPduKey
    Erase mRowMinRxTime
    Erase mRowOriginalIndex
    Erase mRowValidInput
    Erase mRowRXCount
    Erase mWritten
    Erase mPoolRows
    Erase mPoolCestStartRows
    Erase mPoolCestEndRows
    Erase mPoolCestSFN
    Erase mOutputData

    mPoolCount = 0
    mPoolMinSFN = 0
    mPoolMaxSFN = 0
    mPoolCenter = 0#
    mPoolCestCount = 0
    mOutputCount = 0
    mOutputWritePos = 0
    mScanPos = 0
    mPoolCountResolved = 0
    mMaxObservedPoolSize = 0
    mRemainingViolations = 0
    mUnresolvedAttemptCount = 0
    mDiagCount = 0

    mFindConflictSeconds = 0#
    mBuildPoolSeconds = 0#
    mResolvePoolSeconds = 0#
    mWritePoolSeconds = 0#
    mFinalizeSeconds = 0#
    mResolveEntirePoolSeconds = 0#
    mResolveOneCsetSeconds = 0#
    mTryPlaceMoveSeconds = 0#
    mLegalCheckSeconds = 0#
    mBucketExcludeSeconds = 0#
    mBucketAddSeconds = 0#
    mValidateBucketSeconds = 0#
    mResolveOneCsetExtractSeconds = 0#
    mResolveOneCsetPostSeconds = 0#
End Sub

Public Sub Run_TX_SFNConflictResolution()
    MsgBox "Run_TX_SFNConflictResolution is a wrapper. Call TX_SFNConflictResolution from PickExp.", vbInformation, "TX_SFN Conflict Resolution " & MODULE_VERSION_TXSFNCR
End Sub

Public Sub TX_SFNConflictResolution( _
    ByRef data As Variant, ByVal filteredCount As Long, ByVal idxSFNCol As Long, ByVal idxTXID As Long, _
    ByVal idxTXQ As Long, ByVal idxLEN As Long, ByVal idxTXperSFN As Long, ByVal idxRxCnt As Long, _
    ByVal idxAvg As Long, ByVal idxTotLat As Long, ByVal idxGen As Long, ByRef rxDataColIdx() As Long, _
    ByRef rxStationIDs() As Long, ByVal activeRxCount As Long, ByRef dictS2V As Object, ByRef dictVC As Object, _
    ByRef dictA2P As Object, ByRef dictP2R As Object, ByRef dictP2Sigma As Object, ByVal txBitmap As String, _
    ByVal bitmapLen As Long, ByRef elapsedSeconds As Double)

    Dim startTime As Double
    Dim t0 As Double
    Dim conflictStart As Long
    Dim totalFindSeconds As Double
    Dim totalBuildSeconds As Double
    Dim totalResolveSeconds As Double
    Dim totalWriteSeconds As Double
    Dim totalFinalSeconds As Double
    Dim totalLoopCount As Long
    Dim oldStatusBar As Variant
    Dim tPhase As Double

    Reset_TX_SFNConflictResolution_State

    startTime = MicroTimer_TXSFNCR()
    oldStatusBar = Application.StatusBar
    Application.StatusBar = "TX_SFN conflict resolution running..."

    InitializeContext data, filteredCount, idxSFNCol, idxTXID, idxTXQ, idxLEN, idxTXperSFN, idxRxCnt, idxAvg, idxTotLat, idxGen, rxDataColIdx, rxStationIDs, activeRxCount, dictS2V, dictVC, dictA2P, dictP2R, dictP2Sigma, txBitmap, bitmapLen

    If Not ValidateInputMonotoneTXSFN() Then GoTo CleanExit

    PrepareRowDerivedData
    InitializeOutputBuffer
    mTxBitmap = txBitmap

    If Not IsAllOnesBitmap(mTxBitmap) Then
        MsgBox "TX_SFN Conflict Resolution skipped: TX bitmap is not all ones.", vbInformation, "TX_SFN Conflict Resolution"
        GoTo CleanExit
    End If

    Do
        If (mScanPos Mod PROGRESS_STEP_ROWS) = 0 Then
            UpdateProgressBar mScanPos
        End If
        t0 = MicroTimer_TXSFNCR()

        tPhase = MicroTimer_TXSFNCR()
        conflictStart = FindNextConflictStart()
        totalFindSeconds = totalFindSeconds + (MicroTimer_TXSFNCR() - tPhase)

        If conflictStart <= 0 Then Exit Do

        tPhase = MicroTimer_TXSFNCR()
        BuildPoolFromConflictStart conflictStart
        totalBuildSeconds = totalBuildSeconds + (MicroTimer_TXSFNCR() - tPhase)

        tPhase = MicroTimer_TXSFNCR()
        ResolveEntirePool
        totalResolveSeconds = totalResolveSeconds + (MicroTimer_TXSFNCR() - tPhase)

        tPhase = MicroTimer_TXSFNCR()
        WriteResolvedPoolToOutput
        totalWriteSeconds = totalWriteSeconds + (MicroTimer_TXSFNCR() - tPhase)

        If mPoolCount > 0 Then
            mScanPos = mPoolRows(mPoolCount) + 1
        End If

        mFindConflictSeconds = mFindConflictSeconds + (MicroTimer_TXSFNCR() - t0)
        totalLoopCount = totalLoopCount + 1
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

' ... remainder of module unchanged ...

Private Sub UpdateProgressBar(ByVal currentRow As Long)
    Application.StatusBar = "Conflict Resolution at row " & _
                            Format$(currentRow, "0") & _
                            " of " & Format$(mFilteredCount, "0")
End Sub

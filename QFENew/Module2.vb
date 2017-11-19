Module Module2
    Public qbuyonlastday() As Boolean
    Public Structure SeriesStatistics
        Public Main As SeriesStats
        Public last10 As SeriesStats, last20 As SeriesStats, last30 As SeriesStats
    End Structure
    Public Structure SeriesStats
        Public Days As Integer, startDay As Integer, endDay As Integer
        Public Trades As Integer, W As Integer, L As Integer, P As Single, N As Single
        Public wPcntg As Single, lPcntg As Single, totDH As Single, avgDH As Single, maxDH As Single
        Public Runs As Integer, expRuns As Single
        Public ROR1 As Single, ROR2 As Single, ROR3 As Single, ROR4 As Single, Kelly As Single
        Public winRatio As Single, stdDeviation As Single
        Public Peak As Single, Trough As Single, trdPeak As Integer, trdTrough As Integer
        Public sumSquaredDiff As Single, drawDown As IndicatorStats, drawUp As IndicatorStats
        Public Profits As IndicatorStats, Profits1 As IndicatorStats, WProfits As IndicatorStats, LProfits As IndicatorStats
        Public DH As IndicatorStats, DH1 As IndicatorStats, WDH As IndicatorStats, LDH As IndicatorStats
        Public Losers As IndicatorStats, Winners As IndicatorStats
        Public profits00 As Single, profits01 As Single, profits02 As Single, profits03 As Single, profits04 As Single, profits05 As Single
        Public profits06 As Single, profits07 As Single, profits08 As Single, profits09 As Single, profits10 As Single
        Public profits11 As Single, profits12 As Single, profits13 As Single, profits14 As Single, profits15 As Single
        Public zScore As IndicatorStats, zScoreMode As IndicatorStats, zScoreValue As IndicatorStats
        Public Correlation As IndicatorStats, COV As IndicatorStats
        Public Q As IndicatorStats, Q1 As IndicatorStats, absQuantum As Single, qThreshhold As Single
        Public bMA1 As IndicatorStats, bMA2 As IndicatorStats, bMA3 As IndicatorStats
        Public bTD0 As IndicatorStats, bTD1 As IndicatorStats, bTD2 As IndicatorStats, bTD3 As IndicatorStats, bTD4 As IndicatorStats, bTD5 As IndicatorStats
        Public bDOW As IndicatorStats, bMonth As IndicatorStats, bInDay As IndicatorStats, bTrigger As IndicatorStats
        Public bSignal As IndicatorStats, sSignal As IndicatorStats
        Public bEntryStats As IndicatorStats, sEntryStats As IndicatorStats
        Public sMaxDH As IndicatorStats
        Public sTD1 As IndicatorStats, sTP As IndicatorStats
    End Structure
    Public Structure IndicatorStats
        Public V As Single, min As Single, max As Single, avg As Single, tot As Single, Pcntg As Single
        Public Hits As Integer, Misses As Integer
        Public hitPcntg As Single, missPcntg As Single
        Public Peak As Single, Trough As Single, trdPeak As Integer, trdTrough As Integer
    End Structure
    Public Structure DayData
        Public DayNo As Integer
        Public dDate_ As String
    End Structure
    Public Structure Parameters
        Public bhit As Pivot, sHit As Pivot, bDOW As Pivot, sDOW As Pivot, bMonth As Pivot
        Public bZScoreMode As Pivot, bZScoreValue As Pivot
        Public sZScore As Pivot, BandH As Pivot
        Public Buy As Pivot, Sell As Pivot
        Public bInday As Pivot, TP As Pivot
        Public bTD0 As Pivot, bTD1 As Pivot, bTD2 As Pivot, bTD3 As Pivot, bTD4 As Pivot, bTD5 As Pivot
        Public bSignal As Pivot, sSignal As Pivot, bMode As Pivot
        Public bTrigger As Pivot, bEntry As Pivot, sTrigger As Pivot, sEntry As Pivot
        Public bMinDH As Pivot, sMinDH As Pivot, bMaxDH As Pivot, sMaxDH As Pivot
        Public bMA1 As Pivot, bMA2 As Pivot, bMA3 As Pivot
        Public sTD1 As Pivot, sTD2 As Pivot
        Public trigDay As Integer, trigDate As String
    End Structure
    Public Structure Results
        Public Symbol As String, sSymbol As String, lSymbol As String
        Public svSymbol As String, securityNumber As Integer, profitThreshhold As Single
        Public finalDay As Integer, finalDate_ As String
        Public timeDate As Date, Days As Integer, posOutliers As Integer, negOutliers As Integer
        Public bhQ As Single, bhShares As Single, bhGrossPr As Single, bhProfit As Single, bhProfitDay As Single, bhAveProfit As Single
        Public bhSPr As Single, bhEPr As Single, lSignalPrice As Single, totalDH As Single, totalProfit As Single, SP As Parameters, bsEntryExit As String
        Public SS As SeriesStatistics, SS1 As SeriesStatistics
        Public lastTrade As Trades, dist() As Single
        Public signal As String, TFSignal As String, tmpProfit As Single, tmpProfit1 As Single
        Public qBuyOnLastDay As Boolean, qBuyOnDay() As Boolean
        Public qBuyOnLastDay0_ As Byte, qPosLTr As Boolean
        Public tradesString As String
    End Structure
    Public Structure Counter
        Public currentSecurity As String, fileName As String
        Public Month As String, ZIterBase As Long
        Public ZIterNo As Long
        Public totalTrades As Long, totSecurities As Integer
        Public SecNo As Integer, FileString As String
        Public Onn As Long, Off As Long, qBuyandHold As Pivot
        Public signalstradesth As Long, signalsSaved As Long, signalsUnSaved As Long
        Public Incr_ As Long, Iteration As Long, totalIterations As Long
        Public SystemNumber As Integer, totTradeNo As Long
        Public iterationsWritten As Long
        Public xDayIdx As Integer, xDay As Integer, evdcnt As Integer
        Public this_Day As Integer, this_Date As String, this_Day1 As Integer, this_Date1 As String
        Public start_DOW As String, end_DOW As String
        Public first_Day As Integer, first_Date As String, first_Day1 As Integer, first_Date1 As String
        Public start_Day As Integer, start_Date As String
        Public last_Day As Integer, last_Date As String, last_day1 As Integer, last_date1 As String
        Public end_day As Integer, end_Date As String, finalDay As Integer, finalDate As String
        Public startSeconds As Long, endSeconds As Long, currentSeconds As Long, elapsedSeconds As Long, iterationsPerSecond As Single
        Public lastDayPrice As Single, lastDayTrigger As Single, Days As Integer
        Public threshHold As Single, THHits As Long, THMisses As Long
    End Structure
    Public Rmain1 As Results

End Module

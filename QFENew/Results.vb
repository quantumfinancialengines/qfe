Public Structure Prices
    Public open As Single
    Public high As Single
    Public low As Single
    Public close As Single
End Structure
Public Structure Pivot
    Public V As Single, Idx As Integer, min As Single, max As Single
    Public qP As Byte, qHit As Boolean, qHiti As Byte
    Public dbIdx As Integer, dbText As String, dbV As Integer
    Public Day As Integer, Date_ As String, Price As Single
    Public Text As String, Text1 As String
End Structure
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
    Public ZScoreM_ As Single, winRatio As Single, stdDeviation As Single
    Public Peak As Single, Trough As Single, trdPeak As Integer, trdTrough As Integer
    Public sumSquaredDiff As Single, drawDown As indicatorStats, drawUp As IndicatorStats
    Public Profits As IndicatorStats, Profits1 As IndicatorStats, WProfits As IndicatorStats, LProfits As IndicatorStats
    Public DH As IndicatorStats, DH1 As IndicatorStats, WDH As IndicatorStats, LDH As IndicatorStats
    Public Losers As indicatorStats, Winners As indicatorStats
    Public profits00 As Single, profits01 As Single, profits02 As Single, profits03 As Single, profits04 As Single, profits05 As Single
    Public profits06 As Single, profits07 As Single, profits08 As Single, profits09 As Single, profits10 As Single
    Public profits11 As Single, profits12 As Single, profits13 As Single, profits14 As Single, profits15 As Single
    Public zScrTrades As Integer, zScore As indicatorStats, zScoreStd As indicatorStats
    Public Correlation As indicatorStats, COV As IndicatorStats
    Public Q As IndicatorStats, Q1 As IndicatorStats, absQuantum As Single, qThreshhold As Single
    Public bExecute As indicatorStats, sExecute As indicatorStats
    Public bZScorestd As indicatorStats, bZscoreMode As indicatorStats
    Public bMA1 As indicatorStats, bMA2 As indicatorStats, bMA3 As indicatorStats
    Public bTD0 As indicatorStats, bTD1 As indicatorStats, bTD2 As indicatorStats, bTD3 As indicatorStats, bTD4 As indicatorStats, bTD5 As indicatorStats
    Public bDOW As indicatorStats, bInDay As indicatorStats, bTrigger As indicatorStats
    Public bSignal As indicatorStats, sSignal As indicatorStats
    Public bEntry As indicatorStats, sEntry As indicatorStats
    Public sMaxDH As indicatorStats
    Public sTD1 As indicatorStats
End Structure
Public Structure IndicatorStats
    Public V As Single, min As Single, max As Single, avg As Single, tot As Single, Pcntg As Single
    Public Hits As Integer, Misses As Integer, hitPcntg As Single, missPcntg As Single
    Public Peak As Single, Trough As Single, trdPeak As Integer, trdTrough As Integer
    Public max_DU As Single, max_DD As Single
End Structure
Public Structure DayData
    Public DayNo As Integer
    Public dDate_ As String
End Structure
Public Structure Sig
    Public qParam As Boolean, Idx As Integer, qHit As Boolean, qHiti As Byte, V As Single
    Public db As Integer, dbIdx As Integer, dbText As String
    Public Day As Integer
    Public Date_ As String, Signal As String, Text As String
    Public actualPrice As Single, executePrice As Single, signalPrice As Single, triggerPrice As Single
    Public actualDay As Integer, executeDay As Integer, signalDay As Integer, triggerDay As Integer
    Public actualDate As String, executeDate As String, signalDate As String, triggerDate As String
End Structure
Public Structure Parameters
    Public bhit As Pivot, sHit As Pivot, bDOW As Pivot, sDOW As Pivot
    Public bZScoreMode As Pivot, bZScoreStd As Pivot
    Public sZScore As Pivot, BandH As Pivot
    Public Buy As Pivot, Sell As Pivot
    Public bInday As Pivot
    Public bTD0 As Pivot, bTD1 As Pivot, bTD2 As Pivot, bTD3 As Pivot, bTD4 As Pivot, bTD5 As Pivot
    Public bSignal As Pivot, sSignal As Pivot, bMode As Pivot
    Public bExecute As Pivot, sExecute As Pivot, bTrigger As Pivot
    Public bEntry As Pivot, sEntry As Pivot
    Public bMinDH As Pivot, sMinDH As Pivot, bMaxDH As Pivot, sMaxDH As Pivot
    Public bMA1 As Pivot, bMA2 As Pivot, bMA3 As Pivot
    Public sTD1 As Pivot, sTD2 As Pivot
    Public trigDay As Integer, trigDate As String
End Structure
Public Structure Trades
    Public timeDate As Date, Price As Prices
    Public bPrice As Single, sPrice As Single, sPrice1 As Single, bAmt As Single, sAmt As Single, sAmt1 As Single, Shares As Single
    Public BuyOnOpen As Pivot, BuyOnLow As Pivot, BuyOnClose As Pivot, Bexemiss As Pivot, BuyOnTrigger As Pivot
    Public BuyOnNextOpen As Pivot, BuyOnNextClose As Pivot, qPastEndDay As Boolean
    Public SellOnOpen As Pivot, SellOnNextOpen As Pivot, SellOnClose As Pivot, SellOnExpired As Pivot
    Public pcntgProfit As Single
    Public TradeNo As Integer, oldTradeNo As Integer, totTradeNum As Long
    Public tradeDayNo As Integer, tradeDate As String, sliderDate As String, ldDayNo As Integer, ldDate As String
    Public bZscore As Sig, bZScoreStd As Sig, bDOW As Sig, sDOW As Sig
    Public bInDay As Sig, bEntry As Sig, sEntry As Sig
    Public MA1 As Sig, MA2 As Sig, MA3 As Sig, Expired As Sig
    Public bTD0 As Sig, bTD1 As Sig, bTD2 As Sig, bTD3 As Sig, bTD4 As Sig, bTD5 As Sig
    Public bActual As Sig, bTrigger As Sig, bExecute As Sig, sExecute As Sig, bSignal As Sig, sSignal As Sig
    Public Profit0 As Single, Profit1 As Single, Profit2 As Single, Profit3 As Single, Profit4 As Single
    Public totProfit As Single, avgProfit As Single, Quantum As Single, totQuantum As Single, absTotQuantum As Single
    Public lastProfit As Single
    Public DH As Single, dh1 As Single, avgDH As Single, totDH As Single, maxDH As Integer
    Public bDate As String, sDate As String, bDayNo As Integer, sDayNo As Integer
    Public Peak As Single, Trough As Single, Highest As Single, Lowest As Single
    Public DrawDn As Single, totDrawDn As Single, maxDrawDn As Single, maxDrawDnQ As Single
    Public Runs As Integer, ConsW As Integer, ConsL As Integer
    Public avgConsW As Single, avgConsL As Single, maxConsW As Integer, maxConsL As Integer
    Public Correl10 As Single, Correl20 As Single
    Public zScore00 As Single, zScore10 As Single, zScore20 As Single, zScore30 As Single
    Public tradeDay As DayData, ltradeDay As DayData
    Public Descr As String
End Structure
Public Structure TS
    Public tradeSeries() As Trades
End Structure
Public Structure TS1
    Public tradeSeries() As Trades
End Structure
Public Structure Results
    Public sSymbol As String, lSymbol As String, svSymbol As String, securityNumber As Integer, profitThreshhold As Single
    Public finalDay As Integer, finalDate_ As String
    Public timeDate As Date, Days As Integer, posOutliers As Integer, negOutliers As Integer
    Public bhQ As Single, bhShares As Single, bhGrossPr As Single, bhProfit As Single, bhProfitDay As Single, bhAveProfit As Single
    Public bhSPr As Single, bhEPr As Single, lSignalPrice As Single, totalDH As Single, totalProfit As Single, SP As Parameters, bsEntryExit As String
    Public SS As SeriesStatistics, SS1 As SeriesStatistics
    Public mTS() As TS, sTS() As TS, lastTrade As Trades, dist() As Single
    Public qBuyOnLastDay() As Boolean, qBuyOnLastDayI() As Integer
    Public signal As String, TFSignal As String, tmpProfit As Single, tmpProfit1 As Single
    Public lastBuyDate0 As String, lastBuyDate1 As String, lastBuyDate2 As String, lastBuyDate3 As String, lastBuyDate4 As String, lastBuyDate5 As String
    Public qBuyonlastDay0 As Boolean, qBuyonlastDay1 As Boolean, qBuyonlastDay2 As Boolean, qBuyonlastDay3 As Boolean, qBuyonlastDay4 As Byte, qBuyonlastDay5 As Boolean
    Public qBuyOnLastDay0_ As Byte, qPosLTr As Boolean
End Structure
Public Structure Counter
    Public totalTrades As Long, totSecurities As Integer, SecNo As Integer
    Public FileString As String, currentSecurity As String
    Public Onn As Long, Off As Long, qBuyandHold As Pivot
    Public signalstradesth As Long, signalsSaved As Long, signalsUnSaved As Long
    Public Incr_ As Long, Iteration As Long, totalIterations As Long, SystemNumber As Long, totTradeNo As Long
    Public iterationsWritten As Long
    Public xDayIdx As Integer, xDay As Integer, evdcnt As Integer
    Public this_Day As Integer, this_Date As String, this_Day1 As Integer, this_Date1 As String
    Public start_DOW As String, end_DOW As String
    Public first_Day As Integer, first_Date As String, first_Day1 As Integer, first_Date1 As String
    Public start_Day As Integer, start_Date As String
    Public last_Day As Integer, last_Date As String, last_day1 As Integer, last_date1 As String
    Public end_day As Integer, end_Date As String, finalDay As Integer, finalDate As String
    Public startSeconds As Long, endSeconds As Long, elapsedSeconds As Long, iterationsPerSecond As Single
    Public lastDayPrice As Single, lastDayTrigger As Single, Days As Integer
    Public threshHold As Single, THHits As Long, THMisses As Long
End Structure

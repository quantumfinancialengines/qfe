
Public Structure Trades
    Public timeDate As Date
    Public bPrice As Single, sPrice As Single, sPrice1 As Single
    Public bAmt As Single, sAmt As Single, sAmt1 As Single, Shares As Single
    Public TP As Single, qSellOnTP As Boolean
    Public BuyOnOpen As Pivot, BuyOnLow As Pivot, BuyOnClose As Pivot, Bexemiss As Pivot, BuyOnTrigger As Pivot
    Public BuyOnNextOpen As Pivot, BuyOnNextClose As Pivot, qPastEndDay As Boolean
    Public SellOnTP As Pivot
    Public SellOnOpen As Pivot, SellOnNextOpen As Pivot, SellOnClose As Pivot, SellOnExpired As Pivot
    Public pcntgProfit As Single
    Public TradeNo As Integer, oldTradeNo As Integer, totTradeNum As Long
    Public tradeDayNo As Integer, tradeDate As String, sliderDate As String, ldDayNo As Integer, ldDate As String
    Public bZScore As Sig, bZscoreMode As Sig, bZScoreValue As Sig
    Public bDOW As Sig, sDOW As Sig, bMonth As Sig
    Public bInDay As Sig, bEntryS As Sig, bOutDay As Sig
    Public MA1 As Sig, MA2 As Sig, MA3 As Sig, Expired As Sig
    Public bTD0 As Sig, bTD1 As Sig, bTD2 As Sig, bTD3 As Sig, bTD4 As Sig, bTD5 As Sig
    Public bActual As Sig, bEntry As Sig, sEntry As Sig, bSignal As Sig, sSignal As Sig, bTrigger As Sig
    Public sTP As Sig
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
    Public TPHits As Integer, TPMisses As Integer
    Public Descr As String
    Public bOpen As Single, bHigh As Single, bLow As Single, bClose As Single
    Public sOpen As Single, sHigh As Single, sLow As Single, sClose As Single
End Structure
Public Structure TS
        Public tradeSeries() As Trades
    End Structure
Public Structure TS1
    Public tradeSeries() As Trades
End Structure
Public Structure Pivot
    Public V As Single, Idx As Integer, min As Single, max As Single
    Public qP As Byte, qHit As Boolean, qHiti As Byte
    Public dbIdx As Integer, dbText As String, dbV As Integer
    Public Day As Integer, Date_ As String, Price As Single
    Public Text As String, Text1 As String
End Structure
Public Structure Sig
    Public qParam As Boolean, Idx As Integer, qHit As Boolean, qHiti As Byte, V As Single
    Public db As Integer, dbIdx As Integer, dbText As String, Day As Integer
    Public Date_ As String, Signal As String, Text As String
    Public actualPrice As Single, executePrice As Single, signalPrice As Single, triggerPrice As Single
    Public actualDay As Integer, executeDay As Integer, signalDay As Integer, triggerDay As Integer
    Public actualDate As String, executeDate As String, signalDate As String, triggerDate As String
End Structure


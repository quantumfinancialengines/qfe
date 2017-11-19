Imports System
Imports System.ComponentModel
Imports System.Threading
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Object

Public Class QFE_Form
    '   Public OleDBC As New OleDbCommand
    Inherits Form
    Public cnt As Integer
    '    Private components As System.ComponentModel.IContainer = Nothing
    Public OleDBC As New OleDbCommand, OleDBDR As OleDbDataReader
    Public strConnString0 As String, strConnString1 As String, strConnString2 As String
    Public h As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName)
    Public h1 As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName)
    '    Public newss01Row As QFEDb1DataSet.ss01Row
    Public foundFile As String
    Public op(2500) As Single, hi(2500) As Single, lo(2050) As Single, cl(2500) As Single
    Private Delegate Sub SetTextCallback(ByVal [text] As String)
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
        Public sumSquaredDiff As Single, drawDown As IndicatorStats, drawUp As IndicatorStats
        Public Profits As IndicatorStats, Profits1 As IndicatorStats, WProfits As IndicatorStats, LProfits As IndicatorStats
        Public DH As IndicatorStats, DH1 As IndicatorStats, WDH As IndicatorStats, LDH As IndicatorStats
        Public Losers As IndicatorStats, Winners As IndicatorStats
        Public profits00 As Single, profits01 As Single, profits02 As Single, profits03 As Single, profits04 As Single, profits05 As Single
        Public profits06 As Single, profits07 As Single, profits08 As Single, profits09 As Single, profits10 As Single
        Public profits11 As Single, profits12 As Single, profits13 As Single, profits14 As Single, profits15 As Single
        Public zScrTrades As Integer, zScore As IndicatorStats, zScoreStd As IndicatorStats
        Public Correlation As IndicatorStats, COV As IndicatorStats
        Public Q As IndicatorStats, Q1 As IndicatorStats, absQuantum As Single, qThreshhold As Single
        Public bExecute As IndicatorStats, sExecute As IndicatorStats
        Public bZScorestd As IndicatorStats, bZscoreMode As IndicatorStats
        Public bMA1 As IndicatorStats, bMA2 As IndicatorStats, bMA3 As IndicatorStats
        Public bTD0 As IndicatorStats, bTD1 As IndicatorStats, bTD2 As IndicatorStats, bTD3 As IndicatorStats, bTD4 As IndicatorStats, bTD5 As IndicatorStats
        Public bDOW As IndicatorStats, bInDay As IndicatorStats, bTrigger As IndicatorStats
        Public bSignal As IndicatorStats, sSignal As IndicatorStats
        Public bEntry As IndicatorStats, sEntry As IndicatorStats
        Public sMaxDH As IndicatorStats
        Public sTD1 As IndicatorStats
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
    Public Rmain As Results
    Public Counters As Counter
    Public Sub Put_Parameters01(ByRef R As Results)
        '       Dim OleDBC As New OleDbCommand
        '       Dim conn0 As New System.Data.OleDb.OleDbConnection
        With R
            .SS.Main.absQuantum = Math.Abs(.SS.Main.Q.V)
            .bsEntryExit =
                Format$(.SS.Main.Q.V, "0.000") & "bh" & Format$(.bhQ, "0.000") & "!" &
                Format$(.SS.Main.ROR1, "0.000") & "#" & Format$(.SS.Main.ROR2, "0.000") & "#" &
                Format$(.SS.Main.Correlation.V, "0.000") &
                "!" & .lSymbol & ":" &
                Strings.Left(.SP.bExecute.Text, 7) & ":" &
                Strings.Left(.SP.bDOW.Text1, 4) & "!!" &
                Strings.Left(.SP.bSignal.Text, 5) & "td1" &
                Strings.Left(.SP.bTD1.Text, 4) & "[" &
                Strings.Left(.SP.bTD1.dbText, 2) & "]:" & "td2" &
                Strings.Left(.SP.bTD2.Text, 4) & "[" &
                Strings.Left(.SP.bTD2.dbText, 2) & "]:" & "td3" &
                Strings.Left(.SP.bTD3.Text, 4) & "[" &
                Strings.Left(.SP.bTD3.dbText, 2) & "]:" & "td4" &
                Strings.Left(.SP.bTD4.Text, 4) & "[" &
                Strings.Left(.SP.bTD4.dbText, 2) & "]:" &
                Strings.Left(.SP.sTD1.Text, 4) &
                Strings.Left(.SP.sSignal.Text, 5) & "!!" &
                "%" & Format$(.SS.Main.wPcntg, "0.00") &
                "dtr" & Format$(.SS.Main.Trades, "0000") & "/" & Format$(.SS.Main.Days, "0000") &
                "w" & Format$(.SS.Main.W, "000") &
                "/l" & Format$(.SS.Main.L, "000") & "-" &
                "DH=" & Format$(.SS.Main.avgDH, "00.00") & "/" & .SP.sMaxDH.Text
            OleDBC.Connection = conn0
            OleDBC.CommandText = "Insert Into Parameters00 VALUES ('" & Counters.totalIterations &
                "','" & .sSymbol &
                "','" & .lSymbol &
                "','" & .bsEntryExit &
               "','" & Now() & "')"
            OleDBC.ExecuteNonQuery()
        End With
    End Sub
    Public Sub Put_Parameters11(ByRef R As Results)
        Dim OleDBC As New OleDbCommand
        Dim conn0 As New System.Data.OleDb.OleDbConnection
        With R
            OleDBC.Connection = conn0
            OleDBC.CommandText = "Insert Into Parameters01 VALUES ('" & Counters.totalIterations &
                "','" & .SP.bExecute.qP &
                "','" & .SP.bExecute.Text &
                "','" & .SS.Main.bExecute.Hits &
                "','" & .SP.bMA1.qP &
                "','" & .SP.bMA1.Text &
                "','" & .SS.Main.bMA1.Hits &
                "','" & .SP.bMA2.qP &
                "','" & .SP.bMA2.Text &
                "','" & .SS.Main.bMA2.Hits &
                "','" & .SP.bMA3.qP &
                "','" & .SP.bMA3.Text &
                "','" & .SS.Main.bMA3.Hits &
                "','" & .SP.bMinDH.V &
                "','" & .SP.bMinDH.Text &
                "','" & .SP.sMinDH.V &
                "','" & "00" &
                "','" & .SP.sMaxDH.V &
                "','" & .SP.sMaxDH.Text &
                "','" & .bsEntryExit &
                "','" & Now() & "')"
            OleDBC.ExecuteNonQuery()
        End With
    End Sub
    Public Sub Put_Parameters21(ByRef R As Results)
        Dim OleDBC As New OleDbCommand
        Dim conn0 As New System.Data.OleDb.OleDbConnection
        Static td1str As String, td2str As String, td3str As String, td4str As String, td5str As String, tdsstr As String
        Stop
        'Call countTD(R, Counters.end_day)
        With R
            td1str = Strings.Left(.SP.bTD1.Text, 4) &
                  Format$(.SP.bTD1.dbV, "00") & "]"
            td2str = Strings.Left(.SP.bTD2.Text, 4) &
                  Format$(.SP.bTD2.dbV, "00") & "]"
            '            td3str = Strings.Left(.SP.bTD3.Text, 4) & _
            '                 Format$(.SP.bTD3.dbV, "00") & "]"
            '          td4str = Strings.Left(.SP.bTD4.Text, 4) & _
            '               Format$(.SP.bTD4.dbV, "00") & "]"
            '        td5str = Strings.Left(.SP.bTD5.Text, 4) & _
            '             Format$(.SP.bTD5.dbV, "00") & "]"
            tdsstr = Strings.Left(.SP.sTD1.Text, 4) &
                  Format$(.SP.sTD1.dbV, "00") & "]"
            .SP.bTD1.qP = .SP.bTD1.Idx > 0
            OleDBC.Connection = conn0
            OleDBC.CommandText = "Insert Into Parameters02 VALUES ('" & Counters.totalIterations &
                "','" & .SS.Main.Days &
                "','" & .SS.Main.Trades &
                "','" & .SP.bTD1.qP &
                "','" & lastTrade.bTD1.qHiti &
                "','" & td1str &
                "','" & lastTrade.bTD1.Text & " " &
                "','" & .SS.Main.bTD1.Hits &
                "','" & .SS.Main.bTD1.Misses &
                "','" & .SP.bTD2.qP &
                "','" & lastTrade.bTD2.qHiti &
                "','" & td2str &
                "','" & lastTrade.bTD2.Text & " " &
                "','" & .SS.Main.bTD2.Hits &
                "','" & .SP.bTD3.qP &
                "','" & lastTrade.bTD3.qHiti &
                "','" & td3str &
                "','" & lastTrade.bTD3.Text & " " &
                "','" & .SS.Main.bTD3.Hits &
                "','" & .SP.bTD4.qP &
                "','" & lastTrade.bTD4.qHiti &
                "','" & td4str &
                "','" & lastTrade.bTD4.Text & " " &
                "','" & .SS.Main.bTD4.Hits &
                "','" & .SP.bTD5.qP &
                "','" & lastTrade.bTD5.qHiti &
                "','" & td5str &
                "','" & lastTrade.bTD5.Text & " " &
                "','" & .SS.Main.bTD5.Hits &
                "','" & .SP.sTD1.qP &
                "','" & .SP.sTD1.qHiti &
                "','" & tdsstr &
                "','" & .SP.sTD1.Text &
                "','" & .SS.Main.sTD1.Hits &
                "','" & Now() & "')"
            OleDBC.ExecuteNonQuery()
        End With
    End Sub
    Public Sub Put_Statistics011(ByRef R As Results, ByRef TS() As Trades)
        '        Dim OleDBC As New OleDbCommand
        '       Dim conn0 As New System.Data.OleDb.OleDbConnection
        With R
            OleDBC.Connection = conn0
            OleDBC.CommandText = "Insert Into Statistics01 VALUES ('" & Counters.totalIterations &
            "','" & .SS.Main.Runs &
            "','" & .SS.Main.expRuns &
            "','" & .SS.Main.zScore.V &
            "','" & .SS.Main.ZScoreM_ &
            "','" & R.lastTrade.TradeNo &
            "','" & R.lastTrade.Profit0 &
            "','" & R.lastTrade.bDayNo &
            "','" & R.lastTrade.bDate &
            "','" & R.lastTrade.sDate &
            "','" & R.lastTrade.bExecute.executePrice &
            "','" & R.lastTrade.sExecute.executePrice &
            "','" & Now() & "')"
            OleDBC.ExecuteNonQuery()
        End With
    End Sub
    Public Sub Put_Trades11(ByRef R As Results, ByRef TS() As Trades)
        Static totQ As Single, totPr As Single, totDH As Single, xx As Integer, xxx As Integer
        Static str1 As String, str2 As String
        '        Dim OleDBC As New OleDbCommand
        '        Dim conn0 As New System.Data.OleDb.OleDbConnection
        OleDBC.Connection = conn0
        totQ = 0.0
        totPr = 0.0
        totDH = 0.0
        str1 = R.lSymbol & ":" &
            Format$(Counters.SystemNumber, "00000") & ":" &
            Format$(Counters.totalIterations, "00000") & ":" &
            Format$(0, "0000")
        str2 = R.sSymbol & ":" & Format$(Counters.SystemNumber, "0000") & ":" & R.SP.bDOW.Text &
            R.SP.bTD1.Text & ":" & R.SP.bTD2.Text & ":" &
            R.SP.sMaxDH.Text & "=" & Format$(R.SS.Main.Q.avg, "0.00")
        R.qPosLTr = TS(R.SS.Main.Trades).Profit0 > 0.0
        If R.qPosLTr Then
            xxx = 1
        Else
            xxx = 0
        End If
        Counters.Incr_ = Counters.Incr_ + 1
        str1 = R.lSymbol & ":" &
            Format$(Counters.SystemNumber, "00000") & ":" &
            Format$(Counters.totalIterations, "00000") & ":" &
            Format$(0, "0000")
        With TS(1)
            .sEntry.Text = "sent"
            .bDate = dtStr2(Counters.start_Day)
            .sDate = dtStr2(Counters.start_Day)
            .timeDate = Strings.Mid(.bDate, 5, 2) & "-" & Strings.Mid(.bDate, 7, 2) & "-" & Strings.Left(.bDate, 4)
            OleDBC.CommandText = "Insert Into Trades VALUES ('" & Counters.Incr_ &
            "','" & Counters.totalIterations &
            "','" & str2 &
            "','" & str1 &
            "','" & R.sSymbol &
            "','" & 0 &
            "','" & 0 &
            "','" & R.SS.Main.Trades &
            "','" & .timeDate &
            "','" & .pcntgProfit &
            "','" & 0 &
            "','" & 0 &
             "','" & 0 &
           "','" & 0 &
            "','" & 0 &
            "','" & .bPrice &
            "','" & .sPrice &
            "','" & .bAmt &
            "','" & .sAmt &
            "','" & .Shares &
            "','" & .bDate &
            "','" & .sDate &
            "','" & .bDayNo &
            "','" & .sDayNo &
            "','" & .DH &
            "','" & .totDH &
            "','" & "dh" & Format$(.maxDH, "00") &
            "','" & R.SP.sMaxDH.Text &
            "','" & Strings.Left(.bEntry.Text, 4) &
            "','" & R.SP.bTD1.Text &
            "','" & R.SP.bTD2.Text & "-" &
            "','" & Strings.Left(R.SP.bTrigger.Text1, 13) & Format$(0, "000.000") &
            "','" & op(TS(1).bDayNo) &
            "','" & hi(TS(1).bDayNo) &
            "','" & lo(TS(1).bDayNo) &
            "','" & cl(TS(1).bDayNo) &
            "','" & op(TS(1).bDayNo) &
            "','" & hi(TS(1).bDayNo) &
            "','" & lo(TS(1).bDayNo) &
            "','" & cl(TS(1).bDayNo) &
            "','" & Strings.Left(.sEntry.Text, 4) &
            "','" & R.SP.bDOW.Text &
            "','" & 0 &
            "','" & R.SP.bZScoreMode.Text &
            "','" & TS(R.SS.Main.Trades).bZscore.qHiti &
            "','" & TS(R.SS.Main.Trades).Profit1 &
            "','" & TS(R.SS.Main.Trades).Profit2 &
            "','" & TS(R.SS.Main.Trades).Profit3 &
            "','" & Now() & "')"
            OleDBC.ExecuteNonQuery()
        End With
        For xx = 4 To R.SS.Main.Trades
            If TS(xx).bZscore.qHit Then
                TS(xx).bZscore.qHiti = 255
            Else
                TS(xx).bZscore.qHiti = 0
            End If
            With TS(xx)
                .TradeNo = xx
                str1 = R.lSymbol & ":" &
                    Format$(Counters.SystemNumber, "00000") & ":" &
                    Format$(Counters.totalIterations, "00000") & ":" &
                    Format$(.TradeNo, "0000")
                Counters.Incr_ = Counters.Incr_ + 1
                If .bDate <> "" Then
                    '              If .sDate = "" Then .sDate = dtStr2(Counters.end_day)
                    '              If .bEntry.Text = "" Then .bEntry.Text = "--"
                    '              If .sEntry.Text = "" Then .sEntry.Text = "--"
                    '                   If .TradeNo = 0 Then Stop
                    .timeDate = Strings.Mid(.bDate, 5, 2) & "-" & Strings.Mid(.bDate, 7, 2) & "-" & Strings.Left(.bDate, 4)
                    OleDBC.CommandText = "Insert Into Trades VALUES ('" & Counters.Incr_ &
                    "','" & Counters.totalIterations &
                    "','" & str2 &
                    "','" & str1 &
                    "','" & R.sSymbol &
                    "','" & .TradeNo &
                    "','" & .oldTradeNo &
                    "','" & R.SS.Main.Trades &
                    "','" & .timeDate &
                    "','" & .pcntgProfit &
                    "','" & .Quantum &
                    "','" & .totQuantum &
                    "','" & .Profit0 &
                    "','" & .totProfit &
                    "','" & .avgProfit &
                    "','" & .bPrice &
                    "','" & .sPrice &
                    "','" & .bAmt &
                    "','" & .sAmt &
                    "','" & .Shares &
                    "','" & .bDate &
                    "','" & .sDate & " " &
                    "','" & .bDayNo &
                    "','" & .sDayNo &
                    "','" & .DH &
                    "','" & .totDH &
                    "','" & "dh" & Format$(.maxDH, "00") &
                    "','" & R.SP.sMaxDH.Text &
                    "','" & Strings.Left(.bEntry.Text, 4) & ":" & Format$(.bPrice, "000.000") &
                    "','" & R.SP.bTD1.Text1 &
                    "','" & R.SP.bTD2.Text1 & "-" &
                    "','" & Strings.Left(R.SP.bTrigger.Text1, 13) & ":" & Format$(.bTrigger.triggerPrice, "000.000") &
                    "','" & op(TS(xx).bDayNo) &
                    "','" & hi(TS(xx).bDayNo) &
                    "','" & lo(TS(xx).bDayNo) &
                    "','" & cl(TS(xx).bDayNo) &
                    "','" & op(TS(xx).bDayNo - 1) &
                    "','" & hi(TS(xx).bDayNo - 1) &
                    "','" & lo(TS(xx).bDayNo - 1) &
                    "','" & cl(TS(xx).bDayNo - 1) &
                    "','" & Strings.Left(.sEntry.Text, 4) &
                    "','" & R.SP.bDOW.Text &
                    "','" & 0 &
                    "','" & R.SP.bZScoreMode.Text &
                    "','" & TS(xx).bZscore.qHiti &
                    "','" & TS(xx).Profit1 &
                    "','" & TS(xx).Profit2 &
                    "','" & TS(xx).Profit3 &
                    "','" & Now() & "')"
                    OleDBC.ExecuteNonQuery()
                End If
            End With
        Next xx
        Application.DoEvents()
    End Sub
    REM
    Public Sub Put_ldTrades1(ByRef R As Results, ByRef TS() As Trades)
        Static totQ As Single, totPr As Single, totDH As Single
        Static xx As Integer
        Static str1 As String
        Dim OleDBC As New OleDbCommand
        Dim conn0 As New System.Data.OleDb.OleDbConnection
        '        strConnString0 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Dad\Documents\dbactivity11.mdb"
        strConnString0 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=Z:\QFE_DB\QFE_StockData.mdb"
        conn0.ConnectionString = strConnString0
        conn0.Open()
        dbTextBox0.Text = strConnString0
        OleDBC.Connection = conn0
        totQ = 0.0
        totPr = 0.0
        totDH = 0.0
        If R.SS.Main.Trades > 0 Then
            For xx = R.SS.Main.Trades To R.SS.Main.Trades
                str1 = R.lSymbol & Format$(Counters.SystemNumber, "0000")
                With TS(xx)
                    '                totDH = totDH + .DH
                    '                totPr = totPr + .Profit
                    ''                totQ = (totPr / 10) / totDH
                    '                .totQuantum = totQ
                    '                .pcntgProfit = .Profit / 10.0
                    '               If .DH = 0 Then
                    '.Quantum = (.Profit / 10.0)
                    '               Else
                    '               .Quantum = (.Profit / 10.0) / .DH
                    '               End If
                    '                If Math.Abs(.Quantum) > 1.0 Then .Quantum = 0.0
                    '                If Math.Abs(.totQuantum) > 1.0 Then .totQuantum = 0.0
                    OleDBC.CommandText = "Insert Into ldTrades VALUES ('" & Counters.totalIterations &
                    "','" & .TradeNo &
                    "','" & .pcntgProfit &
                    "','" & .Quantum &
                    "','" & .totQuantum &
                    "','" & .Profit0 &
                    "','" & .totProfit &
                    "','" & .bPrice &
                    "','" & .sPrice &
                    "','" & .bAmt &
                    "','" & .sAmt &
                    "','" & .Shares &
                    "','" & .bDate &
                    "','" & .sDate &
                    "','" & .bDayNo &
                    "','" & .sDayNo &
                    "','" & Counters.end_day &
                    "','" & Counters.end_Date &
                    "','" & .DH &
                    "','" & .totDH &
                    "','" & .maxDH &
                    "','" & Now() & "')"
                    OleDBC.ExecuteNonQuery()
                End With
            Next xx
        End If
    End Sub
    Public Sub Put_Trades21(ByRef R As Results, ByRef TS() As Trades)
        Static xx As Integer, xxx As Long
        '       Dim OleDBC As New OleDbCommand
        '       Dim conn0 As New System.Data.OleDb.OleDbConnection
        '        OleDBC.Connection = conn0
        For xx = 1 To R.SS.Main.Trades
            With TS(xx)
                If .bZscore.qHit Then
                    .bZscore.qHiti = 255
                Else
                    .bZscore.qHiti = 0
                End If
                If .bExecute.qHit Then
                    .bExecute.qHiti = 255
                Else
                    .bExecute.qHiti = 0
                End If
                If .BuyOnOpen.qHit Then
                    .BuyOnOpen.qHiti = 255
                Else
                    .BuyOnOpen.qHiti = 0
                End If
                If .BuyOnClose.qHit Then
                    .BuyOnClose.qHiti = 255
                Else
                    .BuyOnClose.qHiti = 0
                End If
                If .BuyOnNextOpen.qHit Then
                    .BuyOnNextOpen.qHiti = 255
                Else
                    .BuyOnNextOpen.qHiti = 0
                End If
                If .SellOnOpen.qHit Then
                    .SellOnOpen.qHiti = 255
                Else
                    .SellOnOpen.qHiti = 0
                End If
                If .SellOnClose.qHit Then
                    .SellOnClose.qHiti = 255
                Else
                    .SellOnClose.qHiti = 0
                End If
                If xx = -1 Then
                    OleDBC.CommandText = "Insert Into Trades02 VALUES ('" & Counters.totalIterations &
                    "','" & R.lSymbol &
                    "','" & .TradeNo &
                    "','" & op(.tradeDayNo) &
                    "','" & hi(.tradeDayNo) &
                    "','" & lo(.tradeDayNo) &
                    "','" & cl(.tradeDayNo) &
                    "','" & op(.tradeDayNo - 1) &
                    "','" & hi(.tradeDayNo - 1) &
                    "','" & lo(.tradeDayNo - 1) &
                    "','" & cl(.tradeDayNo - 1) &
                    "','" & op(.tradeDayNo - 2) &
                    "','" & hi(.tradeDayNo - 2) &
                    "','" & lo(.tradeDayNo - 2) &
                    "','" & cl(.tradeDayNo - 2) &
                    "','" & .bDate &
                    "','" & .bExecute.Text &
                    "','" & .sExecute.Text &
                    "','" & .BuyOnOpen.qHiti &
                    "','" & .BuyOnClose.qHiti &
                    "','" & .BuyOnNextOpen.qHiti &
                    "','" & .BuyOnNextOpen.qHiti &
                    "','" & .BuyOnNextClose.qHiti &
                    "','" & .BuyOnNextClose.qHiti &
                    "','" & .SellOnOpen.qHiti &
                    "','" & .SellOnClose.qHiti &
                    "','" & .SellOnNextOpen.qHiti &
                    "','" & .SellOnExpired.qHiti &
                    "','" & .bExecute.executePrice &
                    "','" & .bSignal.signalPrice &
                    "','" & .bSignal.actualPrice &
                    "','" & .bSignal.executePrice &
                    "','" & .bSignal.triggerPrice &
                    "','" & .bSignal.signalPrice &
                    "','" & .sPrice &
                    "','" & 0.0 &
                    "','" & 0.0 &
                    "','" & .Peak &
                    "','" & .Trough &
                    "','" & .Highest &
                    "','" & .Lowest &
                    "','" & .DrawDn &
                    "','" & .maxDrawDn &
                    "','" & .maxDrawDnQ &
                    "','" & .Runs &
                    "','" & .ConsW &
                    "','" & .ConsL &
                    "','" & .avgConsW &
                    "','" & .avgConsL &
                    "','" & .maxConsW &
                    "','" & .maxConsL &
                    "','" & Now() & "')"
                    OleDBC.ExecuteNonQuery()
                    .totTradeNum = .totTradeNum + 1
                    .TradeNo = .TradeNo + 1
                End If
                xxx = xxx + 1
                OleDBC.CommandText = "Insert Into Trades02 VALUES ('" & xxx & "','" & Counters.totalIterations &
                "','" & R.lSymbol &
                "','" & .TradeNo &
                "','" & .tradeDayNo &
                "','" & .bDOW.Text &
                "','" & .bExecute.Text &
                "','" & .bExecute.Date_ &
                "','" & .bExecute.Date_ &
                "','" & op(.bDayNo) &
                "','" & hi(.bDayNo) &
                "','" & lo(.bDayNo) &
                "','" & cl(.bDayNo) &
                "','" & op(.bDayNo - 1) &
                "','" & hi(.bDayNo - 1) &
                "','" & lo(.bDayNo - 1) &
                "','" & cl(.bDayNo - 1) &
                "','" & op(.bDayNo - 2) &
                "','" & hi(.bDayNo - 2) &
                "','" & lo(.bDayNo - 2) &
                "','" & cl(.bDayNo - 2) &
                "','" & .bSignal.Text &
                "','" & .sSignal.Text &
                "','" & .bZscore.qHiti &
                "','" & .bExecute.qHiti &
                "','" & .BuyOnOpen.qHiti &
                "','" & .BuyOnClose.qHiti &
                "','" & .BuyOnNextOpen.qHiti &
                "','" & .BuyOnNextClose.qHiti &
                "','" & .BuyOnNextClose.qHiti &
                "','" & .BuyOnNextClose.qHiti &
                "','" & .SellOnOpen.qHiti &
                "','" & .SellOnClose.qHiti &
                "','" & .SellOnNextOpen.qHiti &
                "','" & .SellOnExpired.qHiti &
                "','" & .bSignal.signalPrice &
                "','" & .bTrigger.triggerPrice &
                "','" & .bActual.actualPrice &
                "','" & .bExecute.executePrice &
                "','" & .bTrigger.triggerPrice &
                "','" & .bActual.actualPrice &
                "','" & .sPrice &
                "','" & .Profit0 &
                "','" & .totProfit &
                "','" & .Peak &
                "','" & .Trough &
                "','" & .Highest &
                "','" & .Lowest &
                "','" & .DrawDn &
                "','" & .maxDrawDn &
                "','" & .maxDrawDnQ &
                "','" & .Runs &
                "','" & .ConsW &
                "','" & .ConsL &
                "','" & .avgConsW &
                "','" & .avgConsL &
                "','" & .maxConsW &
                "','" & .maxConsL &
                "','" & Now() & "')"
                OleDBC.ExecuteNonQuery()
            End With
        Next xx
        '       Application.DoEvents()
    End Sub
    Public Sub Put_Trades31(ByRef R As Results, ByRef TS() As Trades)
        Static xx As Integer
        Dim OleDBC As New OleDbCommand
        Dim conn0 As New System.Data.OleDb.OleDbConnection
        OleDBC.Connection = conn0
        For xx = 1 To R.SS.Main.Trades
            With TS(xx)
                If .bDOW.qHit Then
                    .bDOW.qHiti = 255
                Else
                    .bDOW.qHiti = 0
                End If
                If .bTD1.qHit Then
                    .bTD1.qHiti = 255
                Else
                    .bTD1.qHiti = 0
                End If
                If .bTD2.qHit Then
                    .bTD2.qHiti = 255
                Else
                    .bTD2.qHiti = 0
                End If
                If .bTD3.qHit Then
                    .bTD3.qHiti = 255
                Else
                    .bTD3.qHiti = 0
                End If
                If .bTD4.qHit Then
                    .bTD4.qHiti = 255
                Else
                    .bTD4.qHiti = 0
                End If
                If .bSignal.qHit Then
                    .bSignal.qHiti = 255
                Else
                    .bSignal.qHiti = 0
                End If
                If .bExecute.qHit Then
                    .bExecute.qHiti = 255
                Else
                    .bExecute.qHiti = 0
                End If
                If .MA1.qHit Then
                    .MA1.qHiti = 255
                Else
                    .MA1.qHiti = 0
                End If
                If .MA2.qHit Then
                    .MA2.qHiti = 255
                Else
                    .MA2.qHiti = 0
                End If
                If .MA3.qHit Then
                    .MA3.qHiti = 255
                Else
                    .MA3.qHiti = 0
                End If
                If .sDOW.qHit Then
                    .sDOW.qHiti = 255
                Else
                    .sDOW.qHiti = 0
                End If
                If .sSignal.qHit Then
                    .sSignal.qHiti = 255
                Else
                    .sSignal.qHiti = 0
                End If
                If .sExecute.qHit Then
                    .sExecute.qHiti = 255
                Else
                    .sExecute.qHiti = 0
                End If
                OleDBC.CommandText = "Insert Into Trades03 VALUES ('" & Counters.totalIterations &
                "','" & .TradeNo &
                "','" & .bExecute.qHiti &
                "','" & .bDOW.qHiti &
                "','" & .bTD1.qHiti &
                "','" & .bTD2.qHiti &
                "','" & .bSignal.qHiti &
                "','" & .bExecute.qHiti &
                "','" & .MA1.qHiti &
                "','" & .MA2.qHiti &
                "','" & .MA3.qHiti &
                "','" & .sExecute.qHiti &
                "','" & .bTD1.qHiti &
                "','" & .bTD2.qHiti &
                "','" & .sSignal.qHiti &
                "','" & .sExecute.qHiti &
                "','" & Now() & "')"
                OleDBC.ExecuteNonQuery()
            End With
        Next xx
    End Sub
    Public Sub ConnectDatabase1()
        '        Dim OleDBC As New OleDbCommand
        '        Dim conn0 As New System.Data.OleDb.OleDbConnection
        '        strConnString0 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Dad\Documents\dbactivity11.mdb"
        '        strConnString0 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=Z:\QFE_DB\QFE_StockData.mdb"
        '        strConnString0 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=\\WDMYCLOUD\Public\QFE_DB\QFE_StockData.mdb"
        '        strConnString2 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Dad\Desktop\Documents\QFE_Signals.mdb"
        '     strConnString2 = _
        '         "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=c:\users\dad\onedrive\Qfe\QFE_Signals.mdb"
        '        ::{018D5C66-4533-4307-9B53-224DE2ED1FE6} (file://FS1/Users/Dad/OneDrive)
        'strConnString2 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Dad\OneDrive\Qfe\QFE_Signals.mdb"
        ' ::{018D5C66-4533-4307-9B53-224DE2ED1FE6} (file://FS1/Users/Dad/OneDrive)
        '        strConnString2 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=S:\QFE_Signals1.mdb"
        strConnString2 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Qfe\QFE_Signals.mdb"
        conn2.ConnectionString = strConnString2
        conn2.Open()
        dbTextBox2.Text = strConnString2
        strConnString0 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Dad\Desktop\QFE_StockData1.mdb"
        conn0.ConnectionString = strConnString0
        MsgBox("Opening file", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, strConnString0)
        '        conn0.Open()
        dbTextBox0.Text = strConnString0
        strConnString1 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Qfe\TradesSmall.mdb"
        '        strConnString1 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Dad\Desktop\TradesSmall1.mdb"
        ' strConnString1 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Dad\Desktop\Documents\Database211.mdb"
        ' strConnString1 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=\\WDMYCLOUD\Public\QFE_DB\Database21.mdb"
        conn1.ConnectionString = strConnString1
        conn1.Open()
        dbTextBox1.Text = strConnString1
        '        strConnString2 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=\\WDMYCLOUD\Public\QFE_Signals.mdb"
    End Sub
    Public Function SetLastDayTrade1(ByRef R As Results, ByRef TS() As Trades, ByRef trd As Trades, ByRef thisDay As Integer) As Boolean
        With trd
            .bDOW.qHit = QExeBDOW(thisDay)
            R.SP.bDOW.Text1 = R.SP.bDOW.Text & ":" & Format$(op(thisDay), "000.000") &
                               "!" & Format$(cl(thisDay), "000.000")
            .bDOW.Text = R.SP.bDOW.Text1
            If .bDOW.qHit Then
                .bDOW.qHiti = 255
            Else
                .bDOW.qHiti = 0
            End If
            .bTD1.Idx = R.SP.bTD1.Idx
            .bTD1.dbIdx = R.SP.bTD1.dbIdx
            .bTD1.db = R.SP.bTD1.dbText
            R.SP.bTD1.Text1 = " "
            If .bTD1.Idx > 0 Then
                '                .bTD1.qHit = QTDBuySignal(thisDay, .bTD1)
                R.SP.bTD1.Text = Strings.Left(R.SP.bTD1.Date_, 12) & ":$" &
                 Format$(.bTD1.actualPrice, "000.000") &
                  "on" & .bTD1.Date_ & "&$" & Format$(.bTD1.triggerPrice, "000.000") & "on" & .bTD1.Date_
                .bTD1.Signal = R.SP.bTD1.Text
            Else
                .bTD1.qHit = True
                R.SP.bTD1.Text = Strings.Left(R.SP.bTD1.Date_, 12) & ":$" &
                 Format$(.bTD1.actualPrice, "000.000") &
                  "on" & .bTD1.Date_ & "&$" & Format$(.bTD1.triggerPrice, "000.000") & "on" & .bTD1.Date_
                .bTD1.Signal = R.SP.bTD1.Text
            End If
            If .bTD1.qHit Then
                .bTD1.qHiti = 255
            Else
                .bTD1.qHiti = 0
            End If
            .bTD1.Text = R.SP.bTD1.Text1
            .bTD1.Signal = .bTD1.Text & ":"
            .bTD2.Idx = R.SP.bTD2.Idx
            .bTD2.dbIdx = R.SP.bTD2.dbIdx
            .bTD2.db = R.SP.bTD2.dbText
            If .bTD2.Idx > 0 Then
                '               .bTD2.qHit = QTDBuySignal(thisDay, .bTD2)
                R.SP.bTD2.Text1 =
                    Strings.Left(R.SP.bTD2.Text, 12) & ":" &
                    Format$(.bTD2.executePrice, "000.000") & "on" & .bTD2.Text & "&" &
                    Format$(.bTD2.executePrice, "000.000") & "on" & .bTD2.Text
            Else
                R.SP.bTD2.Text = R.SP.bTD2.Text
                .bTD2.qHit = True
            End If
            If .bTD2.qHit Then
                .bTD2.qHiti = 255
            Else
                .bTD2.qHiti = 0
            End If
            .bTD2.Text = R.SP.bTD2.Text1
            ''            .bZscore.qHit = QZ(R, TS, R.SS.Main.Trades, 0)
            If .bZscore.qHit Then
                .bZscore.qHiti = 255
            Else
                .bZscore.qHiti = 0
            End If
            '           .bZscore.Text = R.SP.bZScore.Text & "ltpr=" & strSign(lastTrade.Profit0) & "!" & _
            '               strSign(lastTrade.Profit1) & "!" & strSign(lastTrade.Profit2) & "!" & _
            '              strSign(lastTrade.Profit3) & "!"
            If R.SP.bTrigger.Idx = 0 Then
                '                .qrig.qHit = True
                .bTrigger.Text = R.SP.bTrigger.Text
            Else
                '                .bTrigger.qHit = QBuyTrigger(trd, thisDay)
                '                .bTrig.Signal_ = R.SP.bTrigger.Text
            End If
            If .bTrigger.qHit Then
                .bTrigger.qHiti = 255
            Else
                .bTrigger.qHiti = 0
            End If
            If R.SP.bExecute.Idx = 0 Then
                .bExecute.qHit = True
                .bExecute.Text = R.SP.bExecute.Text & ":" & .bExecute.Text & "@" & Format$(0.0, "000.000")
            Else
                .bExecute.qHit = QBExeExecute(Counters.this_Day)
                .bExecute.Text = R.SP.bExecute.Text & ":" & .bExecute.Text & "@" & Format$(.bExecute.executePrice, "000.000")
            End If
            If .bExecute.qHit Then
                .bExecute.qHiti = 255
            Else
                .bExecute.qHiti = 0
            End If
            .bExecute.qHit = .bTrigger.qHit And .bExecute.qHit And .bDOW.qHit And .bTD1.qHit And .bTD2.qHit And .bZscore.qHit
            If .bExecute.qHit Then
                .bExecute.qHiti = 255
            Else
                .bExecute.qHiti = 0
            End If
            .bExecute.Text = .bEntry.Text
            '           If .bExecute.qHit Then Stop
            SetLastDayTrade1 = .bExecute.qHit
        End With
    End Function
    Public Function SetLastDayTradenozzz1(ByRef R As Results, ByRef TS() As Trades, thisDay As Integer) As Boolean
        Dim lastDayTrade As Trades
        lastDayTrade = New Trades
        With lastDayTrade
            .bDOW.Idx = R.SP.bDOW.Idx
            R.SP.bDOW.Text1 = RMain.SP.bDOW.Text & ":" &
             Format$(op(thisDay), "000.000") & "!" &
              Format$(cl(thisDay), "000.000")
            lastDayTrade.bDOW.Text = R.SP.bDOW.Text1 & "::" & dtStr1(thisDay) & ":" & Format(thisDay, "00000")
            lastDayTrade.bDOW.qHit = QExeBDOW(thisDay)
            'If lastDayTrade.bDOW.qHit Then Stop
            If .bDOW.qHit Then
                .bDOW.qHiti = 255
            Else
                .bDOW.qHiti = 0
            End If
            .bTD1.Idx = R.SP.bTD1.Idx
            .bTD1.dbIdx = R.SP.bTD1.dbIdx
            .bTD1.db = Val(R.SP.bTD1.dbText)
            If .bTD1.Idx > 0 Then
                '               .bTD1.qHit = QTDBuySignal(thisDay, .bTD1)
                R.SP.bTD1.Text = Strings.Left(R.SP.bTD1.Text, 12) & ":$" &
                 Format$(.bTD1.actualPrice, "000.000") &
                  "on" & .bTD1.Date_ & "&$" &
                   Format$(.bTD1.triggerPrice, "000.000") & "on" & .bTD1.Date_
                .bTD1.Signal = RMain.SP.bTD1.Text
            Else
                .bTD1.qHit = True
                R.SP.bTD1.Text = Strings.Left(R.SP.bTD1.Text, 12) & ":$" &
                 Format$(.bTD1.actualPrice, "000.000") &
                  "on" & .bTD1.Date_ & "&$" &
                   Format$(.bTD1.triggerPrice, "000.000") & "on" & .bTD1.Date_
                .bTD1.Signal = R.SP.bTD1.Text
            End If
            If .bTD1.qHit Then
                .bTD1.qHiti = 255
            Else
                .bTD1.qHiti = 0
            End If
            .bTD1.Text = R.SP.bTD1.Text1
            .bTD1.Signal = .bTD1.Text & ":"
            .bTD2.Idx = R.SP.bTD2.Idx
            .bTD2.dbIdx = R.SP.bTD2.dbIdx
            .bTD2.db = R.SP.bTD2.dbText
            If .bTD2.Idx > 0 Then
                '               .bTD2.qHit = QTDBuySignal(thisDay, .bTD2)
                '                R.SP.bTD2.Text1 = _
                '                   Strings.Left(R.SP.bTD2.Text, 12) & ":" & _
                '                  Format$(.bTD2.actualPrice, "000.000") & "on" & .bTD2.Text & "&" & _
                '   Format$(.bTD2.actualPrice, "000.000") & "on" & .bTD2.Text
            Else
                R.SP.bTD2.Text1 = R.SP.bTD2.Text
                .bTD2.qHit = True
            End If
            If .bTD2.qHit Then
                .bTD2.qHiti = 255
            Else
                .bTD2.qHiti = 0
            End If
            .bTD2.Text = R.SP.bTD2.Text1
            If R.SP.bTrigger.Idx = 0 Then
                .bTrigger.qHit = True
                .bTrigger.Text = R.SP.bTrigger.Text
            Else
                .bTrigger.qHit = QBuyTrigger(newTrade, thisDay)
                '                .bTrig.Signal_ = R.SP.bTrigger.Text
            End If
            If .bTrigger.qHit Then
                .bTrigger.qHiti = 255
            Else
                .bTrigger.qHiti = 0
            End If
            If R.SP.bExecute.Idx = 0 Then
                .bExecute.qHit = True
                .bExecute.Text = R.SP.bExecute.Text & ":" &
                 .bExecute.Text & "@" & Format$(0.0, "000.000")
            Else
                .bExecute.qHit = QBExeExecute(thisDay)
                .bExecute.Text = R.SP.bExecute.Text & ":" &
                 .bExecute.Text & "@" & Format$(.bExecute.actualPrice, "000.000")
            End If
            If .bExecute.qHit Then
                .bExecute.qHiti = 255
            Else
                .bExecute.qHiti = 0
            End If
            '            .bZscore.qHit = QZ_(R, R.SS.Trades, 0)
            '           .bZscore.qHit = QZ(R, TS, R.SS.Main.Trades, 0)
            If .bZscore.qHit Then
                .bZscore.qHiti = 255
            Else
                .bZscore.qHiti = 0
            End If
            .bZScoreStd.qHit = True
            If .bZscore.qHit Then
                .bZscore.qHiti = 255
            Else
                .bZscore.qHiti = 0
            End If

            .bInDay.qHit = QbuyInDay(thisDay)
            If .bInDay.qHit Then
                .bInDay.qHiti = 255
                '    Stop
            Else
                .bInDay.qHiti = 0
            End If
            lastDayTrade.bExecute.qHit = lastDayTrade.bDOW.qHit And
                .bInDay.qHit And
                .bZscore.qHit And
                .bZScoreStd.qHit And
                 .bTrigger.qHit And
                  .bTD1.qHit And
                  .bTD2.qHit
            If .bExecute.qHit Then
                .bExecute.qHiti = 255
            Else
                .bExecute.qHiti = 0
            End If
            .bExecute.Text = .bEntry.Text
            SetLastDayTradenozzz1 = lastDayTrade.bExecute.qHit
        End With
    End Function
    Public Function TradeString1(ByRef R As Results, ByRef TS() As Trades) As String
        ' Static MnLastTr As Integer
        'With R
        'MnLastTr = RMain.SS.Main.Trades
        TradeString1 = "d" ' "profitsMn^" & Format(MnLastTr, "0000") &
        '"tr" &
        '' StrSign(TS(MnLastTr).Profit0) & "^" &
        ' StrSign(TS(MnLastTr - 1).Profit0) & "#" & TS(MnLastTr - 1).bDate & "^" &
        '        StrSign(TS(MnLastTr - 2).Profit0) & "#" & TS(MnLastTr - 2).bDate & "^" &
        '       StrSign(TS(MnLastTr - 3).Profit0) & "#" & TS(MnLastTr - 3).bDate & "^" &
        '      StrSign(TS(MnLastTr - 4).Profit0) & "#" & TS(MnLastTr - 4).bDate & "^" &
        '     StrSign(TS(MnLastTr - 5).Profit0) & "#" & TS(MnLastTr - 5).bDate & "^" &
        '    StrSign(TS(MnLastTr - 6).Profit0) & "#" & TS(MnLastTr - 6).bDate & "^" &
        '   StrSign(TS(MnLastTr - 7).Profit0) & "#" & TS(MnLastTr - 7).bDate & "^" &
        '  StrSign(TS(MnLastTr - 8).Profit0) & "#" & TS(MnLastTr - 8).bDate & "^" &
        ' StrSign(TS(MnLastTr - 9).Profit0) & "#" & TS(MnLastTr - 9).bDate & "^" &
        'StrSign(TS(MnLastTr - 10).Profit0) & "#" & TS(MnLastTr - 10).bDate & "^" &
        '        StrSign(TS(MnLastTr - 11).Profit0) & "#" & TS(MnLastTr - 11).bDate & "^" &
        '       StrSign(TS(MnLastTr - 12).Profit0) & "#" & TS(MnLastTr - 12).bDate & "^" &
        '      StrSign(TS(MnLastTr - 13).Profit0) & "#" & TS(MnLastTr - 13).bDate
        ' End With
    End Function
    Public Sub Put_Signals1(ByRef R As Results, ByRef TS() As Trades, conNo As Integer)
        Dim Z_Trades As String, mnTrades As String
        Static OleDBC As New OleDbCommand, tr04 As Integer, tr05 As Integer, tr06 As Integer, tr07 As Integer
        Static tr00 As Integer, tr01 As Integer, tr02 As Integer, tr03 As Integer, tr08 As Integer
        Static tr09 As Integer, tr10 As Integer, tr11 As Integer, tr12 As Integer, tr13 As Integer
        Static tr14 As Integer, tr15 As Integer
        Static div As Single, trds As Integer, statStr As String
        Counters.iterationsWritten = Counters.iterationsWritten + 1
        Me.iterationsWritten.Text = Format$(Counters.iterationsWritten, "000000")
        Select Case conNo
            Case 1
                OleDBC.Connection = conn1
            Case 2
                OleDBC.Connection = conn2
            Case Else
                Stop
        End Select
        If Counters.qBuyandHold.qHit Then
            Counters.qBuyandHold.qHiti = 255
        Else
            Counters.qBuyandHold.qHiti = 0
        End If
        Me.lSymbol.Text = R.lSymbol
        With R
            trds = .SS.Main.Trades
            div = .SS.Main.Trades / 16
            .SP.bEntry.Text = .SP.bExecute.Text
            .SS.Main.bSignal.Pcntg = .SS.Main.bSignal.Hits / .SS.Main.Days
            tr00 = Int(div * 0 + 1)
            tr01 = Int(div * 1)
            tr02 = Int(div * 2)
            tr03 = Int(div * 3)
            tr04 = Int(div * 4)
            tr05 = Int(div * 5)
            tr06 = Int(div * 6)
            tr07 = Int(div * 7)
            tr08 = Int(div * 8)
            tr09 = Int(div * 9)
            tr10 = Int(div * 10)
            tr11 = Int(div * 11)
            tr12 = Int(div * 12)
            tr13 = Int(div * 13)
            tr14 = Int(div * 15)
            tr15 = Int(trds)
            R.SS.Main.profits00 = TS(tr00).totProfit
            R.SS.Main.profits01 = TS(tr01).totProfit
            R.SS.Main.profits02 = TS(tr02).totProfit
            R.SS.Main.profits03 = TS(tr03).totProfit
            R.SS.Main.profits04 = TS(tr04).totProfit
            R.SS.Main.profits05 = TS(tr05).totProfit
            R.SS.Main.profits06 = TS(tr06).totProfit
            R.SS.Main.profits07 = TS(tr07).totProfit
            R.SS.Main.profits08 = TS(tr08).totProfit
            R.SS.Main.profits09 = TS(tr09).totProfit
            R.SS.Main.profits10 = TS(tr10).totProfit
            R.SS.Main.profits11 = TS(tr11).totProfit
            R.SS.Main.profits12 = TS(tr12).totProfit
            .SS.Main.profits13 = TS(tr13).totProfit
            .SS.Main.profits14 = TS(tr14).totProfit
            .SS.Main.profits15 = TS(tr15).totProfit
            R.SS.Main.Profits.Pcntg = R.SS.Main.Profits.avg / 10
            .SP.bTD0.Text = .SP.bTD1.Text & .SP.bTD2.Text
            '            mnTrades = TradeString(RMain, RMainTS.tradeSeries)
            '           Z_Trades = TradeString(R, RMainTS.tradeSeries)
            Me.lSymbol.Text = R.lSymbol
            If R.SS.Main.Q.avg > .lastBuyDate1 Then
                Me.lSymbol.BackColor = SystemColors.MenuHighlight
            Else
                Me.lSymbol.BackColor = SystemColors.Info
            End If
            statStr = "QQ==" & StrSign(R.SS.Main.Q.avg) &
                "avgDH=" & Format(R.SS.Main.avgDH, "0.0") & "--DH=" & Format(R.SS.Main.DH.tot, "0.0") &
             "--avgPr=" & StrSign(R.SS.Main.Profits.avg) &
              "--Trades=" & Format(R.SS.Main.Trades, "00000") &
               "--ZScr=" & StrSign(R.SS.Main.zScore.V) & "-Runs=" & Format(R.SS.Main.Runs, "0000") & "-expR=" & Format(R.SS.Main.expRuns, "0000") &
                "--Correl=" & StrSign(R.SS.Main.Correlation.V) &
                 "--COV" & StrSign(R.SS.Main.COV.V)
            OleDBC.CommandText = "Insert Into Signals VALUES (
            '" & Counters.Iteration &
            "','" & Counters.start_Day &
            "','" & Counters.end_day &
            "','" & Counters.last_Date &
            "','" & Counters.SystemNumber &
            "','" & R.sSymbol &
            "','" & R.SS.Main.absQuantum &
            "','" & R.SS.Main.Q.avg &
            "','" & R.SS.Main.Q.min &
            "','" & R.SS.Main.Q.max &
            "','" & R.SS.Main.Q1.avg &
            "','" & R.SS.Main.Q1.min &
            "','" & R.SS.Main.Q1.max &
            "','" & R.SS.Main.Days &
            "','" & R.SP.bDOW.Text1 &
            "','" & RHitMisses.SS.Main.bDOW.Hits &
            "','" & R.SP.bInday.Text &
            "','" & RHitMisses.SS.Main.bInDay.Hits &
            "','" & lastDayTrade.bTrigger.triggerPrice &
            "','" & R.SP.bTrigger.Text1 &
            "','" & RHitMisses.SS.Main.bTrigger.Hits &
            "','" & R.SP.bTD1.Text1 &
            "','" & RHitMisses.SS.Main.bTD1.Hits &
            "','" & R.SP.bEntry.Text & "!!" & Format$("00", R.SP.sMaxDH.V) & ":" & Format$("0.0", R.SS.Main.avgDH) & "!!" & R.SP.sEntry.Text &
            "','" & statStr &
            "','" & Z_Trades &
            "','" & mnTrades &
            "','" & R.SP.bEntry.Text &
            "','" & R.SP.sEntry.Text &
            "','" & R.SS.Main.Correlation.V &
            "','" & R.SS.Main.zScore.V &
            "','" & TS(R.SS.Main.Trades).Profit0 &
            "','" & TS(R.SS.Main.Trades).bDate &
            "','" & Now() & "')"
            OleDBC.ExecuteNonQuery()
            '            "','" & " " & 'R.lSymbol &
            '           "','" & R.sSymbol & ":" & StrSign(R.SS.Main.Q.avg) & R.SP.bDOW.Text & ":" &
            '         "','" & R.sSymbol & ":" & .SP.bZScoreMode.Text & ":" & R.SP.bZScoreStd.Text &
            '        "','" & R.sSymbol & R.SP.bTrigger.Text &
            '       "','" & R.sSymbol & "!" & ' R.SP.bTD1.Text & "!" & R.SP.bTD2.Text & "!" &
            '      "','" & lastDayTrade.bExecute.qHiti &
            '     "','" & Counters.qBuyandHold.qHiti &
            '    "','" & Counters.qBuyandHold.Text &
            '"','" & R.SS.Main.ROR1 &
            '            "','" & R.SS.Main.ROR2 &
            '           "','" & R.SS.Main.ROR3 &
            '          "','" & R.SS.Main.ROR4 &
            '         "','" & R.SS.Main.Peak &
            '        "','" & R.SS.Main.Trough &
            '       "','" & R.SS.Main.trdPeak &
            '      "','" & R.SS.Main.trdTrough &
            '     "','" & R.SS.Main.drawDown.max &
            '    "','" & R.SS.Main.drawDown.avg &
            '            "','" & R.SS.Main.profits00 &
            '           "','" & R.SS.Main.profits01 &
            '          "','" & R.SS.Main.profits02 &
            '         "','" & R.SS.Main.profits03 &
            '        "','" & R.SS.Main.profits04 &
            '       "','" & R.SS.Main.profits05 &
            '      "','" & R.SS.Main.profits06 &
            '     "','" & R.SS.Main.profits07 &
            '            "','" & R.SS.Main.profits08 &
            '           "','" & R.SS.Main.profits09 &
            '          "','" & R.SS.Main.profits10 &
            '         "','" & R.SS.Main.profits11 &
            '        "','" & R.SS.Main.profits12 &
            '       "','" & R.SS.Main.profits13 &
            '      "','" & R.SS.Main.profits14 &
            '     "','" & R.SS.Main.profits15 &
            '    "','" & R.SS.Main.Q.avg &
            '   "','" & Math.Abs(R.SS.Main.Q.avg) &
            '  "','" & R.bhQ &
            ' "','" & R.bhSPr &
            '            "','" & R.bhEPr &
            '           "','" & R.bhProfit &
            '          "','" & Counters.start_Date &
            '         "','" & Counters.end_Date &
            '       "','" & R.SS.Main.bSignal.tot &
            '      "','" & R.SS.Main.bSignal.Hits &
            '     "','" & R.SS.Main.bSignal.Pcntg &
            '    "','" & R.SS.Main.Trades &
            '   "','" & R.SS.Main.W &
            '  "','" & R.SS.Main.L &
            ' "','" & R.SS.Main.wPcntg &
            '           "','" & R.SS.Main.COV.V &
            '          "','" & R.SS.Main.Correlation.V &
            '         "','" & lastDayTrade.bZscore.qHiti &
            '        "','" & R.SP.bZScoreStd.Text &
            '       "','" & R.SP.bZScoreMode.Text &
            '      "','" & R.SS.Main.zScrTrades &
            '     "','" & R.SS.Main.zScore.V &
            '    "','" & R.SS.Main.Runs &
            '   "','" & R.SS.Main.expRuns &
            '  "','" & RHitMisses.SS.Main.zScore.Hits &
            ' "','" & RHitMisses.SS.Main.zScore.Misses &
            '            "','" & R.SP.bZScoreMode.Text &
            '           "','" & R.SS.Main.zScore.min &
            '          "','" & R.SS.Main.zScore.max &
            '         "','" & R.TS(R.SS.Main.Trades).Profit0 &
            '        "','" & R.TS(R.SS.Main.Trades - 1).Profit0 &
            '       "','" & R.TS(R.SS.Main.Trades - 2).Profit0 &
            '      "','" & R.TS(R.SS.Main.Trades - 3).Profit0 &
            '     "','" & lastDayTrade.bEntry.qHiti &
            '    "','" & R.SP.bEntry.Text &
            '   "','" & RHitMisses.SS.Main.bEntry.Hits &
            '  "','" & RHitMisses.SS.Main.bEntry.Misses &
            ' "','" & lastDayTrade.bEntry.Text &
            '            "','" & lastDayTrade.bExecute.qHiti &
            '           "','" & R.SP.bExecute.Text &
            '          "','" & R.SS.Main.bExecute.Hits &
            '         "','" & R.SS.Main.bExecute.Misses &
            '        "','" & lastDayTrade.bExecute.Text &
            '       "','" & R.SP.sEntry.Text &
            '      "','" & R.SP.bDOW.Text & ":" & R.SP.bExecute.Text & "!" & Format$(R.SP.sMaxDH.V, "00.0") & "!" & R.SP.sEntry.Text &
            '     "','" & lastDayTrade.bDOW.qHiti &
            '    "','" & R.SP.bDOW.Text &
            '  "','" & RHitMisses.SS.Main.bDOW.Misses &
            ' "','" & lastDayTrade.bDOW.Text &
            '"','" & lastDayTrade.bInDay.qHiti &
            '          "','" & RHitMisses.SS.Main.bInDay.Misses &
            '         "','" & R.SP.bInday.Text1 &
            '        "','" & lastDayTrade.bTD0.qHiti &
            '       "','" & R.SP.bTD0.Text &
            '      "','" & RHitMisses.SS.Main.bTD0.Hits &
            '     "','" & RHitMisses.SS.Main.bTD0.Misses &
            '    "','" & lastDayTrade.bTD0.Text &
            '   "','" & lastDayTrade.bTD1.qHiti &
            '  "','" & R.SP.bTD1.Text1 &
            ' "','" & RHitMisses.SS.Main.bTD1.Hits &
            '"','" & RHitMisses.SS.Main.bTD1.Misses &
            '            "','" & lastDayTrade.bTD1.Signal &
            '           "','" & lastDayTrade.bTD2.qHiti &
            '          "','" & R.SP.bTD2.Text &
            '         "','" & RHitMisses.SS.Main.bTD2.Hits &
            '        "','" & RHitMisses.SS.Main.bTD2.Misses &
            '       "','" & lastDayTrade.bTD2.Text &
            '      "','" & lastDayTrade.bTrigger.qHiti &
            '   "','" & RHitMisses.SS.Main.bTrigger.Misses &
            '  "','" & lastDayTrade.bTrigger.Text &
            ' "','" & lastDayTrade.bSignal.signalPrice &
            '"','" & R.SS.Main.Profits.avg &
            '            "','" & R.SS.Main.Profits.tot &
            '           "','" & R.SS.Main.Profits.min &
            '          "','" & R.SS.Main.Profits.max &
            '         "','" & R.SS.Main.Winners.tot &
            '        "','" & R.SS.Main.Winners.avg &
            '       "','" & R.SS.Main.Losers.tot &
            '      "','" & R.SS.Main.Losers.avg &
            '     "','" & R.SS.Main.DH.avg &
            '    "','" & R.SS.Main.DH.tot &
            '   "','" & R.SP.sMaxDH.V &
            '  "','" & op(R.thisDay) &
            ' "','" & hi(R.thisDay) &
            '"','" & lo(R.thisDay) &
            '            "','" & cl(R.thisDay) &
            '            "','" & Now() & "')"
            '           OleDBC.ExecuteNonQuery()
        End With
    End Sub
    Private Sub Put_Base1(ByRef R As Results, conNo As Integer)
        With R
            Select Case conNo
                Case 1
                    OleDBC.Connection = conn1
                Case 2
                    OleDBC.Connection = conn2
                Case Else
                    Stop
            End Select
            OleDBC.CommandText = "Insert Into Base VALUES ('" & Counters.Iteration &
            "','" & Counters.SystemNumber &
            "','" & ";;" &
            "','" & R.sSymbol &
            "','" & Counters.this_Day &
            "','" & Counters.this_Date &
            "','" & Counters.first_Day &
            "','" & Counters.first_Date &
            "','" & Counters.start_Day &
            "','" & Counters.start_Date &
            "','" & Counters.last_Day &
            "','" & Counters.last_Date &
            "','" & Counters.end_day &
            "','" & Counters.end_Date &
            "','" & Now() & "')"
            OleDBC.ExecuteNonQuery()
            'End If
            'End If
        End With
        Me.SignalsGrid.Refresh()
        Application.DoEvents()
    End Sub
    Public Sub Put_LastDay1(ByRef R As Results, conNo As Integer, ByRef lt As Trades)
        Static bEntryStr As String
        Select Case conNo
            Case 1
                OleDBC.Connection = conn1
            Case 2
                OleDBC.Connection = conn2
            Case Else
                Stop
        End Select
        With R
            bEntryStr = ""
            Select Case R.SP.bSignal.Idx
                Case 0
                    lastTrade.bExecute.executePrice = 0.001
                    bEntryStr = "***.***"
                Case 1
                    lastTrade.bExecute.executePrice = op(Counters.end_day - 1)
                    bEntryStr = Format$(op(Counters.end_day - 1), "000.000")
                Case 2
                    lastTrade.bExecute.executePrice = cl(Counters.end_day - 1)
                    bEntryStr = Format$(cl(Counters.end_day - 1), "000.000")
                Case 3
                    lastTrade.bExecute.executePrice = op(Counters.end_day)
                    bEntryStr = Format$(op(Counters.end_day), "000.000")
                Case 4
                    lastTrade.bExecute.executePrice = cl(Counters.end_day)
                    bEntryStr = Format$(cl(Counters.end_day), "000.000")
                Case 5
                    lastTrade.bExecute.executePrice = op(Counters.end_day)
                    bEntryStr = Format$(op(Counters.end_day), "000.000")
                Case 6
                    lastTrade.bExecute.executePrice = cl(Counters.end_day)
                    bEntryStr = Format$(cl(Counters.end_day), "000.000")
                Case 7
                    lastTrade.bExecute.executePrice = op(Counters.end_day)
                    bEntryStr = Format$(lastTrade.bExecute.executePrice, "000.000")
                Case 8
                    lastTrade.bExecute.executePrice = cl(Counters.end_day)
                    bEntryStr = Format$(lastTrade.bExecute.executePrice, "000.000")
                Case Else
                    Stop
            End Select
            R.SP.Buy.Text = R.sSymbol & "-" & Counters.this_Date & "-" &
                "-tr=" & Strings.Left(.SP.bTrigger.Text, 13) & "@" & Format$(.SP.bTrigger.Price, "000.000") &
                "#" & lt.bSignal.Text & ":" &
                "td:" & lt.bTD1.Text & "@" & lt.bTD2.Text & "@" &
                .SP.bDOW.Text & ":" _
                & Strings.Left(.SP.bEntry.Text, 4) & ":" & .SP.bZScoreMode.Text & ":" &
                Format$(.SP.sMaxDH.V, "00") & "/" & Format$(.SS.Main.DH.avg, "00.00") & ":" &
                Strings.Left(.SP.sEntry.Text, 5) & "-on-" & dtStr1(Counters.end_day) &
                "avPr=" & StrSign(Format$(.SS.Main.Profits.avg, "00.000")) &
                "Q==" & StrSign(Format$(.SS.Main.Q.avg, "00.000")) &
                "DTWL%" & Format$(.SS.Main.Days, "0000") & "-" & Format$(.SS.Main.Trades, "0000") & "!" &
                Format$(.SS.Main.Winners.tot, "000") & "/" & Format$(.SS.Main.Losers.tot, "000") & "%" &
                Format$(.SS.Main.wPcntg, "0.00")
            OleDBC.CommandText = "Insert Into lastDay VALUES ('" & Counters.Iteration &
            "','" & Counters.SystemNumber &
            "','" & .SP.bDOW.Text &
            "','" & lt.bExecute.qHiti &
            "','" & lt.bTrigger.qHiti &
            "','" & lt.bZscore.qHiti &
            "','" & lt.bDOW.qHiti &
            "','" & lt.bTD1.qHiti &
            "','" & lt.bTD2.qHiti &
            "','" & lt.bEntry.qHiti &
            "','" & lt.bTrigger.Text & "-" &
            "','" & 0.0 &
            "','" & dtStr2(Counters.end_day) &
            "','" & op(Counters.end_day) &
            "','" & hi(Counters.end_day) &
            "','" & lo(Counters.end_day) &
            "','" & cl(Counters.end_day) &
            "','" & dtStr2(Counters.end_day - 1) &
            "','" & op(Counters.end_day - 1) &
            "','" & hi(Counters.end_day - 1) &
            "','" & lo(Counters.end_day - 1) &
            "','" & cl(Counters.end_day - 1) &
            "','" & lt.bExecute.Text &
            "','" & lt.Profit0 &
            "','" & op(Counters.end_day) &
            "','" & hi(Counters.end_day) &
            "','" & lo(Counters.end_day) &
            "','" & cl(Counters.end_day) &
            "','" & lt.Profit1 &
            "','" & lt.Profit2 &
            "','" & lt.Profit3 &
            "','" & Now() & "')"
            OleDBC.ExecuteNonQuery()
        End With
    End Sub
    Public Sub Put_Statistics021(ByRef R As Results)
        '        Dim OleDBC As New OleDbCommand
        '        Dim conn0 As New System.Data.OleDb.OleDbConnection
        With R
            OleDBC.Connection = conn0
            OleDBC.CommandText = "Insert Into Statistics02 VALUES ('" & Counters.totalIterations &
            "','" & Counters.end_day &
            "','" & lastTrade.bExecute.qHiti &
            "','" & lastTrade.bSignal.qHiti &
            "','" & lastTrade.bDOW.qHiti &
            "','" & lastTrade.bZscore.qHiti &
            "','" & lastTrade.bTD1.qHiti &
            "','" & lastTrade.bTD2.qHiti &
            "','" & lastTrade.bTD3.qHiti &
            "','" & lastTrade.bTD4.qHiti &
            "','" & lastTrade.bTD5.qHiti &
            "','" & R.TFSignal &
            "','" & Now() & "')"
            OleDBC.ExecuteNonQuery()
        End With
    End Sub
    Public Sub Put_distribution1(ByRef RR As Results, ByRef TS() As Trades)
        Static wid As Single, trx As Integer, iddx As Integer, trades_ As Single, wd2 As Single
        '       Dim OleDBC As New OleDbCommand
        '       Dim conn0 As New System.Data.OleDb.OleDbConnection
        With RR
            wid = 10
            wd2 = wid / 2
            trades_ = RR.SS.Main.Trades
            For iddx = 1 To 61
                .dist(iddx) = 0.0
            Next
            For trx = 1 To .SS.Main.Trades - 1
                iddx = Int(TS(trx).Quantum * wid + 1 / wd2)
                If iddx >= 30 Then iddx = 30
                If iddx <= -30 Then iddx = -30
                iddx = iddx + 30
                .dist(iddx) = .dist(iddx) + 1
            Next
            For trx = 1 To 61
                .dist(trx) = .dist(trx) / trades_
            Next
            OleDBC.Connection = conn0
            OleDBC.CommandText = "Insert Into dist VALUES ('" & Counters.Iteration &
            "','" & .sSymbol &
            "','" & .lSymbol &
             "','" & .SP.bDOW.Text &
           "','" & RR.SS.Main.Q.avg &
            "','" & .dist(1) &
            "','" & .dist(2) &
            "','" & .dist(3) &
            "','" & .dist(4) &
            "','" & .dist(5) &
            "','" & .dist(6) &
            "','" & .dist(7) &
            "','" & .dist(8) &
            "','" & .dist(9) &
            "','" & .dist(10) &
            "','" & .dist(11) &
            "','" & .dist(12) &
            "','" & .dist(13) &
            "','" & .dist(14) &
            "','" & .dist(15) &
            "','" & .dist(16) &
            "','" & .dist(17) &
            "','" & .dist(18) &
            "','" & .dist(19) &
            "','" & .dist(20) &
            "','" & .dist(21) &
            "','" & .dist(22) &
            "','" & .dist(23) &
            "','" & .dist(24) &
            "','" & .dist(25) &
            "','" & .dist(26) &
            "','" & .dist(27) &
            "','" & .dist(28) &
            "','" & .dist(29) &
            "','" & .dist(30) &
            "','" & .dist(31) &
            "','" & .dist(32) &
            "','" & .dist(33) &
            "','" & .dist(34) &
            "','" & .dist(35) &
            "','" & .dist(36) &
            "','" & .dist(37) &
            "','" & .dist(38) &
            "','" & .dist(39) &
            "','" & .dist(40) &
            "','" & .dist(41) &
            "','" & .dist(42) &
            "','" & .dist(43) &
            "','" & .dist(44) &
            "','" & .dist(45) &
            "','" & .dist(46) &
            "','" & .dist(47) &
            "','" & .dist(48) &
            "','" & .dist(49) &
            "','" & .dist(50) &
            "','" & .dist(51) &
            "','" & .dist(52) &
            "','" & .dist(53) &
            "','" & .dist(54) &
            "','" & .dist(55) &
            "','" & .dist(56) &
            "','" & .dist(57) &
            "','" & .dist(58) &
            "','" & .dist(59) &
            "','" & .dist(60) &
            "','" & .dist(61) &
            "','" & .SS.Main.Trades &
            "','" & Now() & "')"
            OleDBC.ExecuteNonQuery()
        End With
    End Sub
    Private Sub Form1_Load1(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'QFE_SignalsDataSet2.Symbols' table. You can move, or remove it, as needed.
        '       Me.SignalsGrid.DataSource = Me.SignalsGrid
        '        GetData("select * from Signals")
        'TODO: This line of code loads data into the 'DbActivity1DataSet.Signals' table.
        'You can move, or remove it, as needed.
        '        Me.SignalsTableAdapter.Fill(Me.DbActivity1DataSet.Signals)
        Static xxxx As Integer
        '        Call ConnectDatabase()
        xxxx = -1
        For Each foundFile As String In My.Computer.FileSystem.GetFiles(
        "c:\temp\")
            xxxx = xxxx + 1
            '           securitiesListBox0.Items.Add(foundFile)
        Next
        xxxx = -1
        '        For Each foundFile As String In My.Computer.FileSystem.GetFiles(
        '"\\WDMYCLOUD\Public\Gregory's data\QFE_Data\SectorX")
        '       xxxx = xxxx + 1
        '      securitiesListBox1.Items.Add(foundFile)
        '     Next
        '       For Each foundFile As String In My.Computer.FileSystem.GetFiles("C:\Users\Dad\OneDrive\QFE_Prices\")
        For Each foundFile As String In My.Computer.FileSystem.GetFiles("c:\temp")
            xxxx = xxxx + 1
            '           XSectorSecurities.Items.Add(foundFile)
        Next
        '    Me.Message.Text = "reading Parameters"
        '        Call Rdosecurities()
        '       Call Rdoparams1()
        '      Call Rdoparams2()
        '     Call Rdoparams3()
        '    Call Rdoparams4()
        '   Call Rdoparams5()
        '        Call write_Parameters1()
        My.Application.DoEvents()
        '    Me.Message.Text = "deleting Tables"
        '       Call delete_Tables()
        '   Me.Message.Text = "tables Deleted"
        '       start_Seconds = DateDiff(DateInterval.Second, Now.Date, Now)
        start_Time = Date.Now
        My.Application.DoEvents()
        '        Call Setlistview1cols()
        FileSystem.FileClose(1)
endd:
        'Me.quantumThreshHold.Value = 0.75
        '  Me.quantumThreshholdtxt1.Text = Me.quantumThreshHold.Value
        ' Counters.threshHold = Me.quantumThreshHold.Value
        My.Application.DoEvents()
    End Sub
    Public Sub DoBHTrades_OpnCl1(ByRef R As Results, ByRef TS() As Trades)
        Static xline As Integer
        With R
            R.SP.bDOW.Text = "b_h"
            .SP.BandH.Text = "bhOpnCl"
            .SP.bMaxDH.V = 1
            .SP.bEntry.Text = "bh__Op"
            .SP.bSignal.Text = "bh__Op"
            .SP.bExecute.Text = "bh__Op"
            .SP.sEntry.Text = "bh_nCl"
            .SP.sSignal.Text = "bh_nCl"
            .SP.sExecute.Text = "bh_nCl"
            .SP.bZScoreMode.Text = "0Z--"
            For xline = Counters.start_Day To Counters.end_day
                .SS.Main.Trades = .SS.Main.Trades + 1
                TS(.SS.Main.Trades).TradeNo = .SS.Main.Trades
                TS(.SS.Main.Trades).bDOW.qHit = True
                TS(.SS.Main.Trades).bSignal.qHit = True
                TS(.SS.Main.Trades).bExecute.qHit = True
                TS(.SS.Main.Trades).bDayNo = xline
                TS(.SS.Main.Trades).bDate = dtStr2(TS(.SS.Main.Trades).bDayNo)
                TS(.SS.Main.Trades).sDayNo = xline + 1
                TS(.SS.Main.Trades).sDate = dtStr2(TS(.SS.Main.Trades).sDayNo)
                TS(.SS.Main.Trades).bPrice = op(TS(.SS.Main.Trades).bDayNo)
                TS(.SS.Main.Trades).sPrice = cl(TS(.SS.Main.Trades).sDayNo)
                TS(.SS.Main.Trades).bEntry.Text = .SP.bEntry.Text
                TS(.SS.Main.Trades).sEntry.Text = .SP.sEntry.Text
                TS(.SS.Main.Trades).bAmt = 1000.0
                TS(.SS.Main.Trades).Shares = TS(.SS.Main.Trades).bAmt / TS(.SS.Main.Trades).bPrice
                TS(.SS.Main.Trades).sAmt = TS(.SS.Main.Trades).sPrice * TS(.SS.Main.Trades).Shares
                TS(.SS.Main.Trades).Profit0 = TS(.SS.Main.Trades).sAmt - TS(.SS.Main.Trades).bAmt
                TS(.SS.Main.Trades).Profit1 = TS(.SS.Main.Trades).Profit0
                TS(.SS.Main.Trades).Profit2 = TS(.SS.Main.Trades).Profit1
                TS(.SS.Main.Trades).Profit3 = TS(.SS.Main.Trades).Profit2
                TS(.SS.Main.Trades).Profit4 = TS(.SS.Main.Trades).Profit3
                .SS.Main.bDOW.Hits = .SS.Main.bDOW.Hits + 1
                .SS.Main.bSignal.Hits = .SS.Main.bSignal.Hits + 1
                .SS.Main.bExecute.Hits = .SS.Main.bExecute.Hits + 1
                TS(.SS.Main.Trades).Profit0 = TS(.SS.Main.Trades).sAmt - TS(.SS.Main.Trades).bAmt
                TS(.SS.Main.Trades).Profit1 = TS(.SS.Main.Trades).Profit0
                TS(.SS.Main.Trades).Profit2 = TS(.SS.Main.Trades).Profit1
                TS(.SS.Main.Trades).Profit3 = TS(.SS.Main.Trades).Profit2
                TS(.SS.Main.Trades).Profit4 = TS(.SS.Main.Trades).Profit3
                .SS.Main.Profits.tot = .SS.Main.Profits.tot + TS(.SS.Main.Trades).Profit0
                If TS(.SS.Main.Trades).Profit0 > 0.1 Then
                    .SS.Main.Winners.tot = .SS.Main.Winners.tot + TS(.SS.Main.Trades).Profit0
                    .SS.Main.W = .SS.Main.W + 1
                Else
                    .SS.Main.Losers.tot = .SS.Main.Losers.tot + TS(.SS.Main.Trades).Profit0
                    .SS.Main.L = .SS.Main.L + 1
                End If
                TS(.SS.Main.Trades).DH = 1
                .SS.Main.DH.tot = .SS.Main.DH.tot + TS(.SS.Main.Trades).DH
            Next xline
        End With
    End Sub
    Private Sub DoBHTrades_ClnCl1(ByRef R As Results, ByRef TS() As Trades)
        Static xline As Integer
        '        Call InitR(R)
        With R.SP
            .BandH.Text = "bhClnCl"
            .bEntry.Text = "bh__Cl"
            .bDOW.Idx = 0
            .bDOW.Text = "b_h"
            .bSignal.Idx = 1
            .bSignal.Text = "bh_Cl"
            .bExecute.Text = "bh_Cl"
            .sSignal.Idx = 1
            .sSignal.Text = "bh_nCl_"
            .sExecute.Text = "bh_nCl_"
            .sEntry.Text = "bh_nCl_"
            .bZScoreMode.Text = "0Z--"
            .bMaxDH.Text = "01"
        End With
        For xline = Counters.start_Day To Counters.end_day
            R.SS.Main.Trades = R.SS.Main.Trades + 1
            With TS(R.SS.Main.Trades)
                .bDOW.Text = R.SP.bDOW.Text
                .bDayNo = xline
                .sDayNo = xline + 1
                .DH = .sDayNo - .bDayNo + 1
                .bAmt = 1000.0
                .bPrice = cl(.bDayNo)
                .sPrice = cl(.sDayNo)
                .Shares = .bAmt / .bPrice
                .sAmt = .Shares * .sPrice
                .Profit0 = .sAmt - .bAmt
                .maxDH = 1
                .BuyOnOpen.qHit = True
                .BuyOnClose.qHit = False
                .bExecute.qHit = True
                .bExecute.qHit = True
                '                .BExeMiss.qHit = False
                .BuyOnLow.qHit = False
                .BuyOnNextClose.qHit = False
                '               .BExeTg.qHit = False
            End With
        Next xline
        '        lastTrade = TS(R.SS.Main.Trades)
        lastTrade.bExecute.qHit = True
        lastTrade.bDOW.qHit = True
        lastTrade.bSignal.qHit = True
    End Sub
    Public Sub DobhTrades1(ByRef R As Results, ByRef TS() As Trades, xxxx As Integer)
        If Me.qDoBHOptoCl.Checked Then
            Me.DoBHStatus.ForeColor = Color.Azure
            Me.DoBHStatus.Text = Format$(0.0, "000.000")
            '            Call DoBHTrades_Op_Cl(R)
            '            Call Calculate_BasicStatistics(R, 1, R.SS.Main.Trades)
            '           Call Put_Signals_bh(R, TS(), 2)
            '            Call Saves(R)
        End If
        If Me.qDoBHOptoNOp.Checked Then
            '           Call DoBHTrades_OpnOp(R)
            '           Call Calculate_BasicStatistics(R, 1, R.SS.Main.Trades)
            '           Call Put_Signals_bh(R, 2)
            '           Call Saves(R)
        End If
        If Me.qdoBHOptoNCl.Checked Then
            '           Call DoBHTrades_OpnCl(R)
            '          Call Calculate_BasicStatistics(R, 1, R.SS.Main.Trades)
            '           Call Put_Signals_bh(R, 2)
            '          Call Saves(R)
        End If
        If Me.qDoBHCltoNOp.Checked Then
            '         Call DoBHTrades_ClnOp(R)
            '        Call Calculate_BasicStatistics(R, 1, R.SS.Main.Trades)
            '           Call Put_Signals_bh(R, 2)
            '         Call Saves(R)
        End If
        '       If Me.qDoBHCltoNCl.Checked Then
        '       Call DoBHTrades_ClnCl(R)
        '       Call Calculate_BasicStatistics(R, 1, 0)
        '       Call Put_Signals_bh(R, 2)
        '    Call Saves(R)
        '      End If
        If Me.qDoBandHold.Checked Then
            Counters.qBuyandHold.qHit = True
            '           Call DoBuyandHold(R)
        End If
    End Sub
    Public Sub DoBHTrades_Op_Cl1(ByRef R As Results, ByRef TS() As Trades)
        Static xline As Integer
        '        Call InitR(R)
        With R.SP
            .BandH.Text = "bhOp_Cl"
            .bEntry.Text = "bh__Op"
            .bDOW.Idx = 0
            .bDOW.Text = "b_h"
            .bSignal.Idx = 1
            .bSignal.Text = "bh__Op"
            .bExecute.Text = "bh__Op"
            .sSignal.Idx = 1
            .sSignal.Text = "bh__Cl"
            .sExecute.Text = "bh__Cl"
            .sEntry.Text = .sSignal.Text
            .bZScoreMode.Text = "0Z--"
            .bMaxDH.Text = "01"
        End With
        For xline = Counters.start_Day To Counters.end_day
            R.SS.Main.Trades = R.SS.Main.Trades + 1
            With TS(R.SS.Main.Trades)
                .bDOW.Text = R.SP.bDOW.Text
                .bDayNo = xline
                .sDayNo = xline
                .DH = .sDayNo - .bDayNo + 1
                .bAmt = 1000.0
                .bPrice = op(xline)
                .sPrice = cl(.sDayNo)
                .Shares = .bAmt / .bPrice
                .sAmt = .Shares * .sPrice
                .Profit0 = .sAmt - .bAmt
                .maxDH = 1
                .BuyOnOpen.qHit = True
                .BuyOnClose.qHit = False
                .bExecute.qHit = True
                '                .BExeMiss.qHit = False
                .BuyOnLow.qHit = False
                .BuyOnNextClose.qHit = False
                '                .BExeTg.qHit = False
            End With
        Next xline
        '        lastTrade = TS(R.SS.Main.Trades)
        lastTrade.bExecute.qHit = True
        lastTrade.bDOW.qHit = True
        lastTrade.bSignal.qHit = True
    End Sub
    Public Sub DoBHTrades_OpnOp1(ByRef R As Results, ByRef TS() As Trades)
        Static xline As Integer
        '       Call InitR(R)
        With R.SP
            .BandH.Text = "bhOpnOp"
            .bEntry.Text = "bh__Op"
            .bDOW.Idx = 0
            .bDOW.Text = "b_h"
            .bSignal.Idx = 1
            .bSignal.Text = "bh___Op"
            .bExecute.Text = "bh___Op"
            .sSignal.Idx = 1
            .sSignal.Text = "bh_nOp"
            .sExecute.Text = "bh_nOp"
            .sEntry.Text = .bExecute.Text
            .bZScoreMode.Text = "0Z--"
            .bMaxDH.Text = "01"
        End With
        For xline = Counters.start_Day To Counters.end_day
            R.SS.Main.Trades = R.SS.Main.Trades + 1
            With TS(R.SS.Main.Trades)
                .bDOW.Text = R.SP.bDOW.Text
                .bDayNo = xline
                .sDayNo = xline + 1
                .DH = .sDayNo - .bDayNo + 1
                .bAmt = 1000.0
                .bPrice = op(.bDayNo)
                .sPrice = op(.sDayNo)
                .Shares = .bAmt / .bPrice
                .sAmt = .Shares * .sPrice
                .Profit0 = .sAmt - .bAmt
                .maxDH = 1
                .BuyOnOpen.qHit = True
                .BuyOnClose.qHit = False
                .bExecute.qHit = True
                .bExecute.qHit = True
                '                .BExeMiss.qHit = False
                .BuyOnLow.qHit = False
                .BuyOnNextClose.qHit = False
                .BuyOnNextOpen.qHit = False
                '                .BExeTg.qHit = False
            End With
        Next xline
        '        lastTrade = TS(R.SS.Main.Trades)
        lastTrade.bExecute.qHit = True
        lastTrade.bDOW.qHit = True
        lastTrade.bSignal.qHit = True
    End Sub
    Private Sub DoBuyandHold1(ByRef R As Results, ByRef TS() As Trades)
        Static incr As Integer, ddayt As Integer, dDay As Integer, totPr As Single, totDH As Integer
        With R
            .SP.BandH.Text = Counters.qBuyandHold.Text
            .SP.bDOW.Text = Counters.qBuyandHold.Text
            .SP.sDOW.Text = Counters.qBuyandHold.Text
            .SP.bSignal.Text = Counters.qBuyandHold.Text
            .SP.bExecute.Text = Counters.qBuyandHold.Text
            .SP.sSignal.Text = Counters.qBuyandHold.Text
            .SP.sEntry.Text = R.SP.sSignal.Text
            .SS.Main.Trades = Counters.end_day - Counters.start_Day + 1
            .SS.Main.Runs = 1
            .SS.Main.Days = Counters.end_day - Counters.start_Day + 1
            If .SS.Main.Days < 1 Then Stop
            incr = 0
            totDH = 0
            totPr = 0.0
            .SP.bTrigger.Text = Counters.qBuyandHold.Text
            .SP.bZScoreMode.Text = Counters.qBuyandHold.Text
            .SP.bZScoreStd.Text = Counters.qBuyandHold.Text
            .SP.bInday.Text = Counters.qBuyandHold.Text
            .SP.bInday.Text1 = Counters.qBuyandHold.Text
            .SP.bTD0.Text = Counters.qBuyandHold.Text
            .SP.bTD1.Text = Counters.qBuyandHold.Text
            .SP.bTD2.Text = Counters.qBuyandHold.Text
            Counters.this_Day = Counters.end_day
            dDay = 0
            For ddayt = Counters.start_Day To Counters.end_day - Counters.start_Day - 1
                incr = incr + 1
                dDay = dDay + 1
                TS(dDay).TradeNo = incr
                TS(dDay).bDOW.Text = "EvD"
                TS(dDay).sDOW.Text = "EvD"
                TS(dDay).bPrice = op(dDay)
                TS(dDay).sPrice = op(dDay + 1)
                TS(dDay).bAmt = 1000.0
                TS(dDay).Shares = TS(dDay).bAmt / TS(dDay).bPrice
                TS(dDay).sAmt = TS(dDay).Shares * TS(dDay).sPrice
                TS(dDay).Profit0 = TS(dDay).sAmt - TS(dDay).bAmt
                totDH = totDH + 1
                totPr = totPr + TS(dDay).Profit0
                TS(dDay).TradeNo = dDay
                TS(dDay).bDayNo = ddayt
                TS(dDay).bDate = dtStr2(TS(dDay).bDayNo)
                TS(dDay).sDayNo = ddayt + 1
                TS(dDay).sDate = dtStr2(TS(dDay).sDayNo)
                TS(dDay).DH = 1
                TS(dDay).maxDH = 1
                TS(dDay).dh1 = 1
                TS(dDay).totProfit = totPr
                TS(dDay).totDH = totDH
                TS(dDay).bTrigger.Text = Counters.qBuyandHold.Text
                TS(dDay).bDOW.Text = Counters.qBuyandHold.Text
                TS(dDay).sDOW.Text = Counters.qBuyandHold.Text
            Next ddayt
            .SS.Main.Trades = dDay
            '            lastDayTrade = TS(dDay - 1)
            lastDayTrade.bTD0.Text = "bh"
            lastDayTrade.bTD1.Text = "bh"
            lastDayTrade.bTD2.Text = "bh"
            lastDayTrade.bDOW.Text = "bh"
            lastDayTrade.bExecute.Text = "bh"
            lastDayTrade.bEntry.Text = "bh"
            '           Counters.Iteration = Counters.Iteration + 1
            '           Call Calculate_BasicStatistics(R, 2, R.SS.Main.Trades - 1)
            '            Call Put_Signals(R, TS, 2)
            '            Call saves(R)
        End With
    End Sub
    Private Sub DoDaytrades1(ByRef R As Results, ByRef TS() As Trades)
        R.SP.BandH.Text = "BH---Mon"
        R.SP.bDOW.Text = "1Mon"
        R.SP.bSignal.Text = "0Op"
        R.SP.bExecute.Text = "0Op"
        R.SP.sSignal.Text = "0CL"
        R.SP.sExecute.Text = "0CL"
        lastTrade.bDOW.qHit = (dowNo(Counters.end_day) = 1)
        lastTrade.bExecute.qHit = lastTrade.bDOW.qHit
        lastTrade.bSignal.qHit = lastTrade.bDOW.qHit
        ReDim R.qBuyOnLastDay(Counters.end_day)
        ReDim R.qBuyOnLastDayI(Counters.end_day)
        '        Call DoTradesMon(R, TS)
        '        Call Calculate_BasicStatistics(R, 1, 0)
        '      Call Write_Parameters0(R, TS)
        R.SP.BandH.Text = "BH---Tue"
        R.SP.bDOW.Text = "2Tue"
        R.SP.bSignal.Text = "0Op"
        R.SP.bExecute.Text = "0Op"
        R.SP.sSignal.Text = "0CL"
        R.SP.sExecute.Text = "bhOp"
        lastTrade.bDOW.qHit = (dowNo(Counters.end_day) = 2)
        lastTrade.bExecute.qHit = lastTrade.bDOW.qHit
        lastTrade.bSignal.qHit = lastTrade.bDOW.qHit
        '       Call DoTradesTue(R, TS)
        '       Call Calculate_BasicStatistics(R, 1, 0)
        '        r.qBuyOnLastDay(Counters.end_day) = lastTrade.bDOW.qHit
        '     Call Write_Parameters0(R, TS)
        R.SP.BandH.Text = "BH---Wed"
        R.SP.bDOW.Text = "3Wed"
        R.SP.bSignal.Text = "0Op"
        R.SP.bExecute.Text = "0Op"
        R.SP.sSignal.Text = "0CL"
        R.SP.sExecute.Text = "0CL"
        lastTrade.bDOW.qHit = (dowNo(Counters.end_day) = 3)
        lastTrade.bExecute.qHit = lastTrade.bDOW.qHit
        lastTrade.bSignal.qHit = lastTrade.bDOW.qHit
        '        Call DoTradesWed(R, TS)
        '      Call Calculate_BasicStatistics(R, 1, 0)
        '        r.qBuyOnLastDay(Counters.end_day) = lastTrade.bDOW.qHit
        '     Call Write_Parameters0(R, TS)
        R.SP.BandH.Text = "BH---Thr"
        R.SP.bDOW.Text = "4Thr"
        R.SP.bSignal.Text = "0Op"
        '        r.SP.Text = "0Op"
        R.SP.sSignal.Text = "0CL"
        R.SP.sExecute.Text = "0CL"
        lastTrade.bDOW.qHit = (dowNo(Counters.end_day) = 4)
        lastTrade.bExecute.qHit = lastTrade.bDOW.qHit
        lastTrade.bSignal.qHit = lastTrade.bDOW.qHit
        '       Call DoTradesThr(R, TS)
        '     Call Calculate_BasicStatistics(R, 1, 0)
        '        r.qBuyOnLastDay(Counters.end_day) = lastTrade.bDOW.qHit
        '     Call Write_Parameters0(R, TS)
        R.SP.BandH.Text = "BH---Fri"
        R.SP.bDOW.Text = "5Fri"
        R.SP.bSignal.Text = "0Op"
        R.SP.bExecute.Text = "0Op"
        R.SP.sSignal.Text = "0CL"
        R.SP.sExecute.Text = "0CL"
        lastTrade.bDOW.qHit = (dowNo(Counters.end_day) = 5)
        lastTrade.bExecute.qHit = lastTrade.bDOW.qHit
        lastTrade.bSignal.qHit = lastTrade.bDOW.qHit
        '        Call DoTradesFri(R, TS)
        '    Call Calculate_BasicStatistics(R, 1, 0)
        '       r.qBuyOnLastDay(Counters.end_day) = lastTrade.bDOW.qHit
        '     Call Write_Parameters0(R, TS)
    End Sub
    Private Sub Rdosecurities1()
        Static secno As Integer, xx As Integer
        FileSystem.FileOpen(1, "C:\Qfe\securitieslist0.txt", OpenMode.Input, OpenAccess.Read)
        secno = -1
        Do While Not EOF(1) And secno < Me.XSectorSecurities.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                XSectorSecurities.SetItemChecked(secno, True)
            Else
                XSectorSecurities.SetItemChecked(secno, False)
            End If
        Loop
        FileSystem.FileClose(1)
        '       FileSystem.FileOpen(1, "C:\Qfe\securitieslist1.txt", OpenMode.Input, OpenAccess.Read)
        '      secno = -1
        '     Do While Not EOF(1) And secno < Me.XSectorSecurities.Items.Count - 1
        'secno = secno + 1
        '      FileSystem.Input(1, xx)
        '     If xx = 1 Then
        'Me.XSectorSecurities.SetItemChecked(secno, True)
        '      Else
        '     Me.XSectorSecurities.SetItemChecked(secno, False)
        '    End If
        '   Loop
        '  FileSystem.FileClose(1)
    End Sub
    Private Sub Rdoparams11()
        Static secno As Integer, xx As Integer
        dbTextBox3.Text = "c\qfe:\parameterslist1.txt"
        FileSystem.FileOpen(1, "c:\qfe\parameterslist1.txt", OpenMode.Input, OpenAccess.Read)
        secno = -1
        Do While Not EOF(1) And secno < Me.buyDOW.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                buyDOW.SetItemChecked(secno, True)
            Else
                buyDOW.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < Me.buyTrigger.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                buyTrigger.SetItemChecked(secno, True)
            Else
                buyTrigger.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < Me.sellEntry.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                sellEntry.SetItemChecked(secno, True)
            Else
                sellEntry.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < Me.sellMaxDH.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                sellMaxDH.SetItemChecked(secno, True)
            Else
                sellMaxDH.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < Me.whichDates.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                whichDates.SetItemChecked(secno, True)
            Else
                whichDates.SetItemChecked(secno, False)
            End If
        Loop
        '        FileSystem.Input(1, xx)
        FileSystem.FileClose(1)
        Application.DoEvents()
    End Sub
    Private Sub Rdoparams21()
        Static xx As Integer
        dbTextBox4.Text = "c:\qfe\parameterslist2.txt"
        FileSystem.FileOpen(1, "c:\qfe\parameterslist2.txt", OpenMode.Input, OpenAccess.Read)
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qSaveTrades0.Checked = True
        Else
            qSaveTrades0.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qSaveTrades1.Checked = True
        Else
            qSaveTrades1.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qSaveTrades1a.Checked = True
        Else
            qSaveTrades1a.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qSaveTrades2.Checked = True
        Else
            qSaveTrades2.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qSaveTrades3.Checked = True
        Else
            qSaveTrades3.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qSaveTrades4.Checked = True
        Else
            qSaveTrades4.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qSaveTrades5.Checked = True
        Else
            qSaveTrades5.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qSaveParams0.Checked = True
        Else
            qSaveParams0.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qSaveParams1.Checked = True
        Else
            qSaveParams1.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qSaveParams2.Checked = True
        Else
            qSaveParams2.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qSaveParams3.Checked = True
        Else
            qSaveParams3.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qSaveStats0.Checked = True
        Else
            qSaveStats0.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qSaveStats1.Checked = True
        Else
            qSaveStats1.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qSaveStats2.Checked = True
        Else
            qSaveStats2.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qSaveStats3.Checked = True
        Else
            qSaveStats3.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qSaveDist.Checked = True
        Else
            qSaveDist.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            CheckBox4.Checked = True
        Else
            CheckBox4.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            cbXSector.Checked = True
        Else
            cbXSector.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qOnLastDay.Checked = True
        Else
            qOnLastDay.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qDoBHOptoCl.Checked = True
        Else
            qDoBHOptoCl.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qDoBHCltoNOp.Checked = True
        Else
            qDoBHCltoNOp.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qDoBHOptoNOp.Checked = True
        Else
            qDoBHOptoNOp.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qdoBHOptoNCl.Checked = True
        Else
            qdoBHOptoNCl.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qDoBHCltoNCl.Checked = True
        Else
            qDoBHCltoNCl.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            qDoBandHold.Checked = True
        Else
            qDoBandHold.Checked = False
        End If
        FileSystem.FileClose(1)
        Application.DoEvents()
    End Sub
    Private Sub Rdoparams31()
        Static secno As Integer, xx As Integer
        dbTextBox5.Text = "c:\qfe\parameterslist3.txt"
        FileSystem.FileOpen(1, "c:\qfe\parameterslist3.txt", OpenMode.Input, OpenAccess.Read)
        secno = -1
        Do While secno < Me.buyExecute.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                buyExecute.SetItemChecked(secno, True)
            Else
                buyExecute.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < Me.BuyInDay.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                BuyInDay.SetItemChecked(secno, True)
            Else
                BuyInDay.SetItemChecked(secno, False)
            End If
        Loop
exxit:
        FileSystem.FileClose(1)
        Application.DoEvents()
    End Sub
    Private Sub Rdoparams41()
        Static secno As Integer, xx As Integer
        dbTextBox3.Text = "c\qfe:\parameterslist4.txt"
        FileSystem.FileOpen(1, "c:\qfe\parameterslist4.txt", OpenMode.Input, OpenAccess.Read)
        secno = -1
        Do While Not EOF(1) And secno < BZScoreMode.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                BZScoreMode.SetItemChecked(secno, True)
            Else
                BZScoreMode.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < BZScoreStd.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                BZScoreStd.SetItemChecked(secno, True)
            Else
                BZScoreStd.SetItemChecked(secno, False)
            End If
        Loop
        FileSystem.FileClose(1)
        Application.DoEvents()
    End Sub
    Private Sub Rdoparams51()
        Static secno As Integer, xx As Integer
        TxtParameters05.Text = "c\qfe:\parameterslist5.txt"
        FileSystem.FileOpen(1, "c:\qfe\parameterslist5.txt", OpenMode.Input, OpenAccess.Read)
        secno = -1
        Do While Not EOF(1) And secno < Me.buyTDSignal1.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                buyTDSignal1.SetItemChecked(secno, True)
            Else
                buyTDSignal1.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < Me.buyTDSignal2.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                buyTDSignal2.SetItemChecked(secno, True)
            Else
                buyTDSignal2.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < Me.buyTDSignal3.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                buyTDSignal3.SetItemChecked(secno, True)
            Else
                buyTDSignal3.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < Me.buyTDSignal4.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                buyTDSignal4.SetItemChecked(secno, True)
            Else
                buyTDSignal4.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < Me.buyTDSignal5.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                buyTDSignal5.SetItemChecked(secno, True)
            Else
                buyTDSignal5.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < Me.buyTDDaysBack1.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                buyTDDaysBack1.SetItemChecked(secno, True)
            Else
                buyTDDaysBack1.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < Me.buyTDDaysBack2.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                buyTDDaysBack2.SetItemChecked(secno, True)
            Else
                buyTDDaysBack2.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < Me.buyTDDaysBack3.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                buyTDDaysBack3.SetItemChecked(secno, True)
            Else
                buyTDDaysBack3.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < Me.buyTDDaysBack4.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                buyTDDaysBack4.SetItemChecked(secno, True)
            Else
                buyTDDaysBack4.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < Me.buyTDDaysBack5.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                buyTDDaysBack5.SetItemChecked(secno, True)
            Else
                buyTDDaysBack5.SetItemChecked(secno, False)
            End If
        Loop
        '        FileSystem.Input(1, xx)
        Application.DoEvents()
        FileSystem.FileClose(1)
    End Sub
    Public Sub Delete_Tables1()
        Dim OleDBC As New OleDbCommand
        Dim conn0 As New System.Data.OleDb.OleDbConnection
        '        strConnString0 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Dad\Documents\dbactivity11.mdb"
        '        strConnString0 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=Z:\QFE_DB\QFE_StockData.mdb"
        ' strConnString0 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Dad\Desktop\QFE_StockData1.mdb"
        '  strConnString0 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=\\QFEDISK\QFEDisk\QFE_StockData1.mdb"
        'strConnString0 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Dad\Desktop\TradesSmall1.mdb"
        ' conn0.ConnectionString = strConnString0
        '  conn0.Open()
        ' dbTextBox0.Text = strConnString0
        ' OleDBC.Connection = conn0
        ' With OleDBC
        '           .CommandText = "DELETE FROM Parameters00"
        '           .ExecuteNonQuery()
        '           .CommandText = "DELETE FROM Parameters01"
        '            .ExecuteNonQuery()
        '            .CommandText = "DELETE FROM Parameters02"
        '           .ExecuteNonQuery()
        '            .CommandText = "DELETE FROM Parameters03"
        '            .ExecuteNonQuery()
        '            .CommandText = "DELETE FROM Parameters04"
        '           .ExecuteNonQuery()
        '          MsgBox("Records Params Deleted!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "SUCCESS")
        '           .CommandText = "DELETE FROM statistics00"
        '           .ExecuteNonQuery()
        '           .CommandText = "DELETE FROM statistics01"
        '            .ExecuteNonQuery()
        '           .CommandText = "DELETE FROM statistics02"
        '           .ExecuteNonQuery()
        '           MsgBox("Records Stats Deleted!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "SUCCESS")
        '            .CommandText = "DELETE FROM dist"
        '           .ExecuteNonQuery()
        '            .CommandText = "DELETE FROM Trades00"
        '           .ExecuteNonQuery()
        '.CommandText = "DELETE FROM Trades00"
        ' .ExecuteNonQuery()
        '           .CommandText = "DELETE FROM Tradesa"
        '           .ExecuteNonQuery()
        '            .CommandText = "DELETE FROM Trades02"
        '           .ExecuteNonQuery()
        '           .CommandText = "DELETE FROM Trades03"
        '            .ExecuteNonQuery()
        '            .CommandText = "DELETE FROM Trades04"
        '           .ExecuteNonQuery()
        '        MsgBox("Records Trades Deleted!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "SUCCESS")
        '            .CommandText = "DELETE FROM ldTrades"
        '            .ExecuteNonQuery()
        conn1.Close()
        conn1.Open()
        With OleDBC
            .Connection = conn1
            .CommandText = "DELETE FROM Trades00"
            '           .ExecuteNonQuery()
        End With
        With OleDBC
            .Connection = conn2
            .CommandText = "DELETE FROM Signal_"
            .ExecuteNonQuery()
        End With
        With OleDBC
            .Connection = conn2
            .CommandText = "DELETE FROM Signals"
            .ExecuteNonQuery()
        End With
        With OleDBC
            .Connection = conn2
            .CommandText = "DELETE FROM Base"
            .ExecuteNonQuery()
        End With
        With OleDBC
            .Connection = conn2
            .CommandText = "DELETE FROM lastDay"
            .ExecuteNonQuery()
        End With
        '  End With
        MsgBox("Record Deleted!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "SUCCESS")
    End Sub
    Public Sub Write_AllParameters1()
        '       Call Wdosecurities()
        '      Call Wdoparams01()
        '     Call Wdoparams02()
        '    Call Wdoparams03()
        '   Call Wdoparams04()
        '  Call Wdoparams05()
        My.Application.DoEvents()
    End Sub
    Private Sub Wdosecurities1()
        Static secno As Integer
        FileSystem.FileClose(1)
        My.Computer.FileSystem.DeleteFile("C:\Qfe\securitieslist0.txt")
        FileSystem.FileOpen(1, "C:\Qfe\securitiesList0.txt", OpenMode.Append, OpenAccess.ReadWrite)
        For secno = 0 To XSectorSecurities.Items.Count - 1
            If XSectorSecurities.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        '        FileSystem.FileClose(1)
        '        My.Computer.FileSystem.DeleteFile("C:\Qfe\securitieslist1.txt")
        '        FileSystem.FileOpen(1, "C:\Users\Dad\Documents\securitiesList1.txt", OpenMode.Append, OpenAccess.ReadWrite)
        '        For secno = 0 To XSectorSecurities.Items.Count - 1
        ' If XSectorSecurities.GetItemChecked(secno) Then
        '  FileSystem.Write(1, 1)
        '  Else
        '  FileSystem.Write(1, 0)
        '  End If
        '  Next secno
        ' FileSystem.FileClose(1)
    End Sub
    Private Sub Wdoparams011()
        FileSystem.FileClose(1)
        Static secno As Integer
        My.Computer.FileSystem.DeleteFile("c:\qfe\parameterslist1.txt")
        FileSystem.FileOpen(1, "c:\qfe\parameterslist1.txt", OpenMode.Append, OpenAccess.ReadWrite)
        For secno = 0 To buyDOW.Items.Count - 1
            If buyDOW.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To buyTrigger.Items.Count - 1
            If buyTrigger.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To sellEntry.Items.Count - 1
            If sellEntry.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To sellMaxDH.Items.Count - 1
            If sellMaxDH.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To whichDates.Items.Count - 1
            If whichDates.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        FileSystem.Write(1, Me.ThreshholdQuantum_.SelectedIndex)
        FileSystem.FileClose(1)
        Application.DoEvents()
    End Sub
    Private Sub Wdoparams021()
        My.Computer.FileSystem.DeleteFile("c:\qfe\parameterslist2.txt")
        FileSystem.FileOpen(1, "c:\qfe\parameterslist2.txt", OpenMode.Append, OpenAccess.ReadWrite)
        If qSaveTrades0.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qSaveTrades1.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qSaveTrades1a.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qSaveTrades2.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qSaveTrades3.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qSaveTrades4.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qSaveTrades5.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qSaveParams0.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qSaveParams1.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qSaveParams2.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qSaveParams3.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qSaveStats0.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qSaveStats1.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qSaveStats2.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qSaveStats3.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qSaveDist.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If CheckBox4.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If cbXSector.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qOnLastDay.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qDoBHOptoCl.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qDoBHOptoNOp.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qDoBHCltoNOp.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qDoBHCltoNCl.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qdoBHOptoNCl.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If qDoBandHold.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        FileSystem.FileClose(1)
        Application.DoEvents()
    End Sub
    Private Sub Wdoparams031()
        Static secno As Integer
        My.Computer.FileSystem.DeleteFile("c:\qfe\parameterslist3.txt")
        FileSystem.FileOpen(1, "c:\qfe\parameterslist3.txt", OpenMode.Append, OpenAccess.ReadWrite)
        For secno = 0 To buyExecute.Items.Count - 1
            If buyExecute.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To BuyInDay.Items.Count - 1
            If BuyInDay.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        TxtParameters03.BackColor = SystemColors.MenuHighlight
        FileSystem.FileClose(1)
        Application.DoEvents()
    End Sub
    Private Sub Wdoparams041()
        FileSystem.FileClose(1)
        Static secno As Integer
        My.Computer.FileSystem.DeleteFile("c:\qfe\parameterslist4.txt")
        FileSystem.FileOpen(1, "c:\qfe\parameterslist4.txt", OpenMode.Append, OpenAccess.ReadWrite)
        For secno = 0 To BZScoreMode.Items.Count - 1
            If BZScoreMode.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To BZScoreStd.Items.Count - 1
            If BZScoreStd.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        TxtParameters04.BackColor = SystemColors.MenuHighlight
        FileSystem.FileClose(1)
        Application.DoEvents()
    End Sub
    Private Sub Wdoparams051()
        FileSystem.FileClose(1)
        Static secno As Integer
        My.Computer.FileSystem.DeleteFile("c:\qfe\parameterslist5.txt")
        FileSystem.FileOpen(1, "c:\qfe\parameterslist5.txt", OpenMode.Append, OpenAccess.ReadWrite)
        For secno = 0 To buyTDSignal1.Items.Count - 1
            If buyTDSignal1.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To buyTDSignal2.Items.Count - 1
            If buyTDSignal2.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To buyTDSignal3.Items.Count - 1
            If buyTDSignal3.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To buyTDSignal4.Items.Count - 1
            If buyTDSignal4.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To buyTDSignal5.Items.Count - 1
            If buyTDSignal5.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To buyTDDaysBack1.Items.Count - 1
            If buyTDDaysBack1.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To buyTDDaysBack2.Items.Count - 1
            If buyTDDaysBack2.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To buyTDDaysBack3.Items.Count - 1
            If buyTDDaysBack3.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To buyTDDaysBack4.Items.Count - 1
            If buyTDDaysBack4.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To buyTDDaysBack5.Items.Count - 1
            If buyTDDaysBack5.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        FileSystem.FileClose(1)
        TxtParameters05.BackColor = SystemColors.MenuHighlight
        Application.DoEvents()
    End Sub
    Private Sub Run_Click1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Run.Click
        MsgBox("Starting . . ..", MsgBoxStyle.Exclamation, "Initialization")
        MessageBox.Show("Starting1 . . ..")
        Counters.totalIterations = 0
        Counters.THHits = 0
        Counters.THMisses = 0
        Counters.signalsSaved = 0
        Counters.signalsUnSaved = 0
        Counters.totTradeNo = 0
        Counters.SystemNumber = 0
        Counters.Iteration = 0
        Counters.iterationsWritten = 0
        Counters.elapsedSeconds = 0
        Counters.THHits = 0
        Counters.THMisses = 0
        Counters.signalsSaved = 0
        Counters.signalstradesth = 0
        Counters.Onn = 0
        Counters.Off = 0
        Counters.threshHold = Me.quantumThreshHold.Value
        Me.quantumThreshholdtxt1.Text = Format$(Counters.threshHold)
        start_Seconds = DateDiff(DateInterval.Second, Now.Date, Now)
        '        Call Write_AllParameters()
        Me.Run.BackColor = Color.Green
        Counters.THHits = 0
        OleDBC.Connection = conn2
        Application.EnableVisualStyles()
        '        ReDim RMainTS
        '       ReDim TS1(2500)
        '       ReDim RZA__TS(2500)
        '      ReDim RZA__TS1(2500)
        '        Call RunIt(RMain)
        Run.BackColor = Color.Red
    End Sub
    Public Sub RunIt()
        '        strConnString1 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Dad\Desktop\Documents\QFE_Signals.mdb"
        '       conn1.ConnectionString = strConnString1
        '      conn1.Open()
        '   Me.txtMessages.Text = "deleting Tables"
        ' Call delete_Tables()
        '  If QuantumThreshHolds.SelectedItem = -1 Then
        ' QuantumThreshHolds.SelectedIndex = 0
        'End If
        '       txtQuantumThreshHolds.Text = Val(Me.QuantumThreshHolds.SelectedItem)
        '  threshHoldTrades = Val(Me.txtQuantumThreshHolds.Text)
        ' threshHoldQuantum = Val(Me.ThreshholdQuantum_.Text)
        Counters.Iteration = 0
        Counters.SystemNumber = 0
        sysBuyOnLastDay0 = 0
        sysBuyOnLastDay1 = 0
        sysBuyOnLastDay2 = 0
        sysBuyOnLastDay3 = 0
        sysBuyOnLastDay4 = 0
        sysBuyOnLastDay5 = 0
        Counters.totalIterations = 0
        Counters.signalsSaved = 0
        Counters.totTradeNo = 0
        Call doSecAll()
        '       Call doSecSectorX()
        conn1.Close()
    End Sub
    Private Sub DoAllSecuritiesLoop1(ByRef R As Results)
        Static tmpDay As Integer, rowss As Integer
        Me.Message.Text = "processing Files"
        Me.SignalsGrid.BackgroundColor = Color.Bisque
        If Me.cbXSector.Checked = True Then
            With Counters
                rowss = 0
                For .SecNo = 0 To XSectorSecurities.Items.Count - 1
                    XSectorSecurities.SelectedIndex = .SecNo
                    If XSectorSecurities.GetItemChecked(.SecNo) Then
                        .currentSecurity = XSectorSecurities.Items(.SecNo).ToString
                        '                        R.sSymbol = .currentSecurity
                        Me.txtCurrent_Security.Text = .currentSecurity
                        '                        Call Process_Files(.currentSecurity, .SecNo, 1, R)
                        Me.endDate.Text = Counters.end_Date & DOWStr(Counters.end_day)
                        Me.startDate.Text = Counters.start_Date
                        Me.days.Text = Counters.Days
                        Me.startDay.Text = Counters.start_Day
                        Me.endDay.Text = Counters.end_day
                        Counters.this_Day = Counters.end_day
                        Counters.this_Date = Counters.end_Date
                        Counters.elapsedSeconds = current_Seconds - start_Seconds
                        Counters.iterationsPerSecond = Counters.totalIterations / Counters.elapsedSeconds
                        Me.iterpersecond.Text = Format$(Counters.iterationsPerSecond, "0.00")
                        Me.elapsedSeconds_.Text = Format$(Counters.elapsedSeconds, "0")
                        .end_Date = dtStr2(Counters.end_day)
                        .Days = .end_day - .start_Day + 1
                        .SystemNumber = 0
                        .xDayIdx = -1
                        '                        Call Put_Base(R, 2)
                        For tmpDay = .end_day To .last_Day - (Me.whichDates.Items.Count - 1) Step -1
                            .xDayIdx = .xDayIdx + 1
                            Me.whichDates.Items(.xDayIdx) = Format$(.xDayIdx, "000") & " " &
                                Format$(tmpDay, "0000") & " " & R.sSymbol
                        Next
                        Application.DoEvents()
                        Counters.qBuyandHold.Text = "bhNon"
                        Counters.qBuyandHold.Text1 = "bhNon"
                        Counters.qBuyandHold.qHit = False
                        '                        Call Do_WhichDayLoop(R)
                    End If
                    If XSectorSecurities.GetItemChecked(.SecNo) Then
                        If Me.qDoBandHold.Checked Then
                            Counters.qBuyandHold.Text = "bhEvD"
                            Counters.qBuyandHold.Text1 = "bhEvD"
                            Counters.qBuyandHold.qHit = True
                            Counters.SystemNumber = Counters.SystemNumber + 1
                            Counters.totalIterations = Counters.totalIterations + 1
                            '                           Call Me.DobhTrades(R, RMainTS.tradeSeries, -1)
                            If Me.qSaveTrades1a.Checked Then
                                '                               Call Me.Put_Trades1a(R, RMainTS.tradeSeries)
                            End If
                        End If
                    End If
                Next .SecNo
            End With
        End If
    End Sub
    Private Sub Do_WhichDayLoop1(ByRef R As Results)
        With Counters
            .last_Date = dtStr2(.last_Day)
            .xDayIdx = -1
            For .last_day1 = .last_Day To .last_Day - 4 Step -1
                .last_date1 = dtStr2(.last_day1)
                .xDayIdx = .xDayIdx + 1
                Me.whichDates.Items(.xDayIdx) = Format$(.xDayIdx, "000") & " " &
                    Format$(.last_day1, "0000") & " " & .last_Date & " " & .last_date1 & " " &
                    " op:" & Format$(op(.last_day1), "000.000") & " hi:" & Format$(hi(.last_day1), "000.000") &
                    " lo:" & Format$(lo(.last_day1), "000.000") & " cl:" & Format$(cl(.last_day1), "000.000") & " " &
                    Format$(Counters.Days, "00000") & " " & R.sSymbol
                If Me.whichDates.GetItemChecked(.xDayIdx) Then
                    Me.whichDates.SelectedIndex = .xDayIdx
                    Me.txtNowDates.Text = Now
                    '                    Res = New Results
                    R.sSymbol = .currentSecurity
                    Counters.qBuyandHold.qHit = False
                    Call ByDOW()
                End If
                Application.DoEvents()
            Next
        End With
    End Sub

    Private Sub SetBTrigger_Click1(sender As Object, e As EventArgs) Handles setBTrigger.Click
        Static xx As Integer
        For xx = 0 To buyTrigger.Items.Count - 1
            buyTrigger.SetItemChecked(xx, True)
        Next xx
    End Sub

    Private Sub ClearBTrigger_Click1(sender As Object, e As EventArgs) Handles clearBTrigger.Click
        Static xx As Integer
        For xx = 0 To buyTrigger.Items.Count - 1
            buyTrigger.SetItemChecked(xx, False)
        Next xx
    End Sub

    Private Sub InitBuyZScore_Click1(sender As Object, e As EventArgs) Handles initBuyZScore.Click
        Static xx As Integer
        For xx = 0 To BZScoreMode.Items.Count - 1
            BZScoreMode.SetItemChecked(xx, True)
        Next xx
    End Sub
    Private Sub ClearBuyZScore_Click1(sender As Object, e As EventArgs) Handles clearBuyZScore.Click
        Static xx As Integer
        For xx = 0 To BZScoreMode.Items.Count - 1
            BZScoreMode.SetItemChecked(xx, False)
        Next xx
    End Sub

    Private Sub Button2_Click1(sender As Object, e As EventArgs) Handles initBuyInDay.Click
        Static xx As Integer
        For xx = 0 To BuyInDay.Items.Count - 1
            BuyInDay.SetItemChecked(xx, True)
        Next xx
    End Sub

    Private Sub Button5_Click1(sender As Object, e As EventArgs) Handles Button5.Click
        Static xx As Integer
        buyTDDaysBack1.SetItemChecked(0, True)
        buyTDDaysBack2.SetItemChecked(0, True)
        buyTDDaysBack3.SetItemChecked(0, True)
        buyTDDaysBack4.SetItemChecked(0, True)
        buyTDDaysBack5.SetItemChecked(0, True)
        SellTDDaysBack1.SetItemChecked(0, True)
        buyTDSignal1.SetItemChecked(0, True)
        buyTDSignal2.SetItemChecked(0, True)
        buyTDSignal3.SetItemChecked(0, True)
        buyTDSignal4.SetItemChecked(0, True)
        buyTDSignal5.SetItemChecked(0, True)
        SellTDSignal1.SetItemChecked(0, True)
        For xx = 1 To buyTDDaysBack1.Items.Count - 1
            buyTDDaysBack1.SetItemChecked(xx, False)
        Next xx
        For xx = 1 To buyTDDaysBack2.Items.Count - 1
            buyTDDaysBack2.SetItemChecked(xx, False)
        Next xx
        For xx = 1 To buyTDDaysBack3.Items.Count - 1
            buyTDDaysBack3.SetItemChecked(xx, False)
        Next xx
        For xx = 1 To buyTDDaysBack4.Items.Count - 1
            buyTDDaysBack4.SetItemChecked(xx, False)
        Next xx
        For xx = 1 To buyTDDaysBack5.Items.Count - 1
            buyTDDaysBack5.SetItemChecked(xx, False)
        Next xx
        For xx = 1 To SellTDDaysBack1.Items.Count - 1
            SellTDDaysBack1.SetItemChecked(xx, False)
        Next xx
        For xx = 1 To buyTDSignal1.Items.Count - 1
            buyTDSignal1.SetItemChecked(xx, False)
            buyTDSignal2.SetItemChecked(xx, False)
            buyTDSignal3.SetItemChecked(xx, False)
            buyTDSignal4.SetItemChecked(xx, False)
            buyTDSignal5.SetItemChecked(xx, False)
            SellTDSignal1.SetItemChecked(xx, False)
        Next xx
    End Sub

    Private Sub QuantumThreshHold_ValueChanged1(sender As Object, e As EventArgs) Handles quantumThreshHold.ValueChanged
        Me.quantumThreshHold.ForeColor = Color.Aqua
        Me.quantumThreshholdText.Text = quantumThreshHold.Value
        Counters.threshHold = Me.quantumThreshHold.Value
        My.Application.DoEvents()
    End Sub

End Class

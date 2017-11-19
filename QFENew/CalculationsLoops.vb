Module CalculationsLoops
    Public Const xxx = 7000
    Public conn0 As New System.Data.OleDb.OleDbConnection, conn1 As New System.Data.OleDb.OleDbConnection
    Public conn2 As New System.Data.OleDb.OleDbConnection, conn4 As New System.Data.OleDb.OleDbConnection
    Public RHitMisses As Results
    Public RMain As Results, RMainTS(xxx) As Trades, RMainTS1 As TS
    Public RZA__ As Results, RZA__TS(xxx) As Trades, RZA__TS1 As TS
    Public RzP__ As Results, RZP__TS(xxx) As Trades
    Public RzN__ As Results, RZN__TS(xxx) As Trades
    Public RzPP_ As Results, RZPP_TS(xxx) As Trades
    Public RzNN_ As Results, RZNN_TS(xxx) As Trades
    Public RzPPP As Results, RzNNN As Results, RzPN_ As Results, RzNP_ As Results
    Public RzPNP As Results, RzNPN As Results, RzPPN As Results, RzNNP As Results
    Public RzNPP As Results, RzPNN As Results
    Public RBH___ As Results
    Public Counters As Counter, newTrade As Trades, newTrade1 As Trades, lastTrade As Trades, lastDayTrade As Trades
    Public DH As Integer, displaySecCnt As Integer, Peak As Single, Trough As Single
    Public bh0TotalProfit As Single, bh0AvgProfit As Single, bh0AvgWProfit As Single, bh0AvgLProfit As Single
    Public bh1TotalProfit As Single, bh1AvgProfit As Single, bh1AvgWProfit As Single, bh1AvgLProfit As Single
    Public bh2TotalProfit As Single, bh2AvgProfit As Single, bh2AvgWProfit As Single, bh2AvgLProfit As Single
    Public lastdayTrigger As Single, lastdayPrice As Single, lTrigger As Single, lPrice As Single
    Public lastBuyDay0 As Integer, lastBuyDay1 As Integer, lastBuyDay2 As Integer, lastBuyDay3 As Integer, lastBuyDay4 As Integer, lastBuyDay5 As Integer
    Public lastBuyDate0 As String, lastBuyDate1 As String, lastBuyDate2 As String, lastBuyDate3 As String, lastBuyDate4 As String, lastBuyDate5 As String
    Public qBuyOnLastDay0 As Boolean, qBuyOnLastDay1 As Boolean, qBuyOnLastDay2 As Boolean, qBuyOnLastDay3 As Boolean, qBuyOnLastDay4 As Boolean, qBuyOnLastDay5 As Boolean
    Public sysBuyOnLastDay0 As Integer, sysBuyOnLastDay1 As Integer, sysBuyOnLastDay2 As Integer, sysBuyOnLastDay3 As Integer, sysBuyOnLastDay4 As Integer, sysBuyOnLastDay5 As Integer
    Public ss As Array
    Public start_Seconds As Single, current_Seconds As Single, elapsed_Seconds As Single, upDated As String
    Public iterspersecond_ As Single
    Public start_Time As Date, current_Time As Date
    Public dtStr1(xxx) As String, dtStr2(xxx) As String
    Public op() As Single, hi() As Double, lo() As Double, cl() As Double, dyno() As Integer
    Public DOWStr() As String, dowNo() As Integer
    Public monthNo() As Integer, monthStr() As String, bhop(xxx) As Double, bhhi(xxx) As Double, bhlo(xxx) As Double, bhcl(xxx) As Double
    Public days_Records As Integer, xLine1 As Integer, qWhichDayAll As Boolean, qWhichDayOn As Boolean, qWhichDayOff As Boolean
    Public iterCountMax As Long, threshHoldTrades As Integer, threshHoldQuantum As Single

    '    Public Result As results
    Public Sub Process_Files(xsec As String, listNo As Integer, xLst As Integer, ByRef R As Results)
        Dim ssx As Integer, ssy As Integer, zero As Single
        Dim ss As String, ss0 As String, ss1 As String, ss2 As String, ss3 As String, ss4 As String, ln As Integer, lyr As String, ldy As String, lmo As String
        Dim lop As Single, lhi As Single, llo As Single, lcl As Single
        Dim comma0 As Integer, comma1 As Integer, comma2 As Integer, comma3 As Integer, comma4 As Integer
        ' Me.ToolStripContainer1.bottomtoolstrippanel.Text = TimeOfDay
        FileClose(1)
        FileOpen(1, xsec, OpenMode.Input, OpenAccess.Read)
        ss0 = LineInput(1)
        xLine1 = 0
        QFE1.DataGridView1.RowCount = 1
        Do While Not EOF(1)
            ss0 = LineInput(1)
            xLine1 = xLine1 + 1
            QFE1.DataGridView1.Rows.Add()
            QFE1.DataGridView1.Item(0, xLine1 - 1).Value = xLine1
        Loop
        ReDim dtStr1(xLine1)
        ReDim dtStr2(xLine1)
        ReDim op(xLine1)
        ReDim hi(xLine1)
        ReDim lo(xLine1)
        ReDim cl(xLine1)
        ReDim dyno(xLine1)
        ReDim DOWStr(xLine1)
        ReDim dowNo(xLine1)
        ReDim monthNo(xLine1)
        ReDim monthStr(xLine1)
        ReDim qbuyonlastday(xLine1)
        '     ReDim R.qBuyOnLastDayI(Counters.end_day)
        FileClose(1)
        FileOpen(1, xsec, OpenMode.Input, OpenAccess.Read)
        ss1 = LineInput(1)
        Select Case xLst
            Case 0
                ss = QFE_Base.XSectorSecurities.Items(listNo)
                ln = ss.Length
                ssx = ss.IndexOf(" ")
                Select Case ln
                    Case 52
                        ss1 = ss.Substring(47, 5)
                        Counters.currentSecurity = ss1.Substring(0, 1)
                    Case 53
                        ss1 = ss.Substring(47, 6)
                        Counters.currentSecurity = ss1.Substring(0, 2)
                    Case 54
                        ss1 = ss.Substring(47, 7)
                        Counters.currentSecurity = ss1.Substring(0, 3)
                    Case 55
                        ss1 = ss.Substring(47, 8)
                        Counters.currentSecurity = ss1.Substring(0, 4)
                    Case Else
                        Stop
                End Select
            Case 1
                ln = Len(xsec)
                ssx = xsec.IndexOf(".")
                ssy = xsec.LastIndexOf("\")
                If ssy > ssx Then Stop
                '                ss1 = ss.Substring(ssy + 1, ssx - ssy - 1) & ".csv"
                ss1 = xsec.Substring(ssy + 1, ssx - ssy - 1) & ".csv"
                Counters.currentSecurity = ss1.Substring(0, ssx - ssy - 1)
        End Select
        R.sSymbol = Counters.currentSecurity
        Counters.currentSecurity = Strings.Left(Counters.currentSecurity & "_____", 5)
        ''''        Form1.Current_Security.Text = Counters.currentSecurity
        R.SS.Main.DH.tot = 0
        xLine1 = 0
        ss0 = LineInput(1)
        QFE1.DataGridView1.Refresh()
        Do While Not EOF(1)
            ss0 = LineInput(1)
            If Strings.InStr(ss0, ",-,") Then GoTo skipp : 
            If Strings.InStr(ss0, ",0,") Then GoTo skipp : 
            xLine1 = xLine1 + 1
            '            dSplit = ss1.Split(",")
            comma0 = InStr(ss0, ",")
            lyr = Mid(ss0, 1, 4)
            lmo = Mid(ss0, 6, 2)
            ldy = Mid(ss0, 9, 2)
            ss1 = ss0.Substring(comma0)
            comma1 = InStr(ss1, ",") - 1
            lop = Val(ss1.Substring(0, comma1))
            ss2 = ss1.Substring(comma1 + 1, ss1.Length - comma1 - 1)
            comma2 = InStr(ss2, ",") - 1
            lhi = Val(ss2.Substring(0, comma2))
            ss3 = ss2.Substring(comma2 + 1, ss2.Length - comma2 - 1)
            comma3 = InStr(ss3, ",") - 1
            llo = Val(ss3.Substring(0, comma3))
            ss4 = ss3.Substring(comma3 + 1, ss3.Length - comma3 - 1)
            comma4 = InStr(ss4, ",") - 1
            lcl = Val(ss4.Substring(0, comma4))
            op(xLine1) = lop
            hi(xLine1) = lhi
            If lhi = 0.0 Then lhi = lop
            lo(xLine1) = llo
            If llo = 0.0 Then llo = lcl
            If llo = 0.0 Then Stop
            cl(xLine1) = lcl
            '            dtStr1(xLine) = lmo & "/" & ldy & "/" & lyr
            dtStr1(xLine1) = lyr & "/" & lmo & "/" & ldy
            dtStr2(xLine1) = lyr & lmo & ldy
            dyno(xLine1) = ldy
            dowNo(xLine1) = Weekday(dtStr1(xLine1)) - 1
            DOWStr(xLine1) = WeekdayName(dowNo(xLine1) + 1)
            monthNo(xLine1) = Month(dtStr1(xLine1)) - 1
            monthStr(xLine1) = MonthName(monthNo(xLine1) + 1)
            '    If lop < 1 Then Stop
            '   If lhi < 1 Then Stop
            '  If llo < 1 Then Stop
            ' If lcl < 1 Then Stop
            '   Stop
            zero = 0.0
            If Double.IsNaN(0 / op(xLine1)) Then Stop
            If Double.IsNaN(0 / hi(xLine1)) Then Stop
            If Double.IsNaN(0 / lo(xLine1)) Then Stop
            If Double.IsNaN(0 / cl(xLine1)) Then Stop
            QFE1.DataGridView1.Item(1, xLine1 - 1).Value = dtStr1(xLine1)
            QFE1.DataGridView1.Item(2, xLine1 - 1).Value = Format(op(xLine1), "000.000")
            QFE1.DataGridView1.Item(3, xLine1 - 1).Value = Format(hi(xLine1), "000.000")
            QFE1.DataGridView1.Item(4, xLine1 - 1).Value = Format(lo(xLine1), "000.000")
            QFE1.DataGridView1.Item(5, xLine1 - 1).Value = Format(cl(xLine1), "000.000")
            '            Form1.DataGridView1.Rows.Item(xLine).Cells(4).Value = Format(hi(xLine1 - xLine), "000.000")
            '            Form1.DataGridView1.Rows.Item(xLine).Cells(5).Value = Format(lo(xLine1 - xLine), "000.000")
            '            Form1.DataGridView1.Rows.Item(xLine).Cells(6).Value = Format(cl(xLine1 - xLine), "000.000")

next1t:
            Application.DoEvents()
skipp:
        Loop
        FileClose(1)
        days_Records = Val(xLine1)
        Counters.end_day = xLine1
        Counters.start_Day = xLine1 - QFE1.daysBack.Value
        If Counters.start_Day < 100 Then Counters.start_Day = 100
        Counters.first_Day = Counters.start_Day
        Counters.last_Day = Counters.end_day
        Counters.first_Date = dtStr2(Counters.first_Day)
        Counters.start_Date = dtStr2(Counters.start_Day)
        Counters.last_Date = dtStr2(Counters.last_Day)
        Counters.end_Date = dtStr1(Counters.end_day)
        Counters.Days = xLine1
        QFE1.startDay.Text = Counters.start_Date & "--" &
            Format$(Counters.start_Day, "000000") & ":" & DOWStr(Counters.start_Day)
        QFE1.firstDateTxt.Text = Counters.first_Date & "--" &
            Format$(Counters.first_Day, "000000") & ":" & DOWStr(Counters.first_Day)
        QFE1.endDateText.Text = Counters.end_Date & "--" &
            Format$(Counters.end_day, "000000") & ":" & DOWStr(Counters.end_day)
        QFE1.lastDateText.Text = Counters.last_Date & "--" &
            Format$(Counters.last_Day, "000000") & ":" & DOWStr(Counters.last_Day)
        'QFE_Base.whichDates.Items.Add("00000000000000000")
        Application.DoEvents()
    End Sub
    Public Sub Process_Files1(xsec As String, listNo As Integer, ByRef R As Results)
        Static ss As String, ss0 As String, ss1 As String, ss2 As String, ss3 As String, ss4 As String
        Static lyr As String, ldy As String, lmo As String, ln As Integer
        Static lop As Single, lhi As Single, llo As Single, lcl As Single, xline As Integer
        Static comma0 As Integer, comma1 As Integer, comma2 As Integer, comma3 As Integer, comma4 As Integer
        ReDim dtStr1(0)
        ReDim dtStr2(0)
        ReDim op(0)
        ReDim hi(0)
        ReDim lo(0)
        ReDim cl(0)
        ReDim dyno(0)
        ReDim DOWStr(0)
        ReDim dowNo(0)
        ss = QFE_Base.securitiesListBox0.Items(listNo)
        ln = ss.Length
        Select Case ln
            Case 48
                ss1 = ss.Substring(43, 5)
                R.sSymbol = ss1.Substring(0, 1)
            Case 49
                ss1 = ss.Substring(43, 6)
                R.sSymbol = ss1.Substring(0, 2)
            Case 50
                ss1 = ss.Substring(43, 7)
                R.sSymbol = ss1.Substring(0, 3)
            Case 51
                ss1 = ss.Substring(43, 8)
                R.sSymbol = ss1.Substring(0, 4)
            Case Else
                Stop
        End Select
        '        R.lSymbol = R.sSymbol
        ''''        Current_Security.Text = R.sSymbol
        FileClose(1)
        FileOpen(1, ss, OpenMode.Input, OpenAccess.Read)
        xline = 0
        ss1 = LineInput(1)
        R.securityNumber = R.securityNumber + 1
        Do While Not EOF(1)
            ss0 = LineInput(1)
            xline = xline + 1
        Loop
        ReDim dtStr1(xline)
        ReDim dtStr2(xline)
        ReDim op(xline)
        ReDim hi(xline)
        ReDim lo(xline)
        ReDim cl(xline)
        ReDim dyno(xline)
        ReDim DOWStr(xline)
        ReDim dowNo(xline)
        FileClose(1)
        FileOpen(1, ss, OpenMode.Input, OpenAccess.Read)
        xline = 0
        ss0 = LineInput(1)
        ss0 = LineInput(1)
        With R.SS
            .Main.DH.tot = 0
            .Main.Winners.V = 0
            .Main.Losers.V = 0
            .Main.Profits.tot = 0.0
            .Main.Winners.tot = 0.0
            .Main.Losers.tot = 0.0
            Do While Not EOF(1)
                ss0 = LineInput(1)
                xline = xline + 1
                comma0 = InStr(ss0, ",")
                lyr = Mid(ss0, 1, 4)
                lmo = Mid(ss0, 6, 2)
                ldy = Mid(ss0, 9, 2)
                ss1 = ss0.Substring(comma0)
                comma1 = InStr(ss1, ",") - 1
                lop = Val(ss1.Substring(0, comma1))
                ss2 = ss1.Substring(comma1 + 1, ss1.Length - comma1 - 1)
                comma2 = InStr(ss2, ",") - 1
                lhi = Val(ss2.Substring(0, comma2))
                ss3 = ss2.Substring(comma2 + 1, ss2.Length - comma2 - 1)
                comma3 = InStr(ss3, ",") - 1
                llo = Val(ss3.Substring(0, comma3))
                ss4 = ss3.Substring(comma3 + 1, ss3.Length - comma3 - 1)
                comma4 = InStr(ss4, ",") - 1
                lcl = Val(ss4.Substring(0, comma4))
                op(xline) = lop
                hi(xline) = lhi
                lo(xline) = llo
                cl(xline) = lcl
                dtStr1(xline) = lmo & "/" & ldy & "/" & lyr
                dtStr2(xline) = lyr & lmo & ldy
                dyno(xline) = ldy
                dowNo(xline) = Weekday(dtStr1(xline)) - 1
                DOWStr(xline) = WeekdayName(dowNo(xline) + 1)
                '   Stop
            Loop
            Counters.Days = xline
            op(xline + 1) = lop
            hi(xline + 1) = lhi
            lo(xline + 1) = llo
            cl(xline + 1) = lcl
            dtStr1(xline + 1) = lmo & "/" & ldy & "/" & lyr
            '            dtStr2(xline + 1) = lyr & lmo & ldCounters.first_Day
            dyno(xline + 1) = ldy
            dowNo(xline + 1) = Weekday(dtStr1(xline + 1)) - 1
            FileClose(1)
            days_Records = Val(xline)
            Counters.first_Day = 100
            Counters.first_Date = dtStr2(Counters.first_Day)
            Counters.last_Day = xline
            Counters.last_Date = dtStr2(Counters.last_Day)
            QFE_Base.startDayDate.Text = Counters.first_Date & "--" & Format$(Counters.first_Day, "000000") & ":" & DOWStr(Counters.first_Day)
        End With
    End Sub
    Public Sub InitR(ByRef R As Results)
        Counters.Iteration = Counters.Iteration + 1
        R.SS.Main.qThreshhold = QFE1.quantumThreshHold.Text
        R.profitThreshhold = 75.55 'Me.profitThreshhold.Text
        '  ReDim R.TS(xxx)
        With R.SS.Main
            .Profits.tot = 0.0
            .Profits.avg = 0.0
            .Losers.tot = 0.0
            .Winners.tot = 0.0
            .DH.avg = 0.0
            .DH.tot = 0
            .Trades = 0
            .W = 0
            .L = 0
            .wPcntg = 0.0
            .bSignal.Hits = 0
            .bSignal.tot = 0.0
            .bSignal.Misses = 0
            .bEntryStats.Hits = 0
            .bMonth.Hits = 0
            .bMonth.Misses = 0
            .bSignal.Misses = 0
            .bSignal.Hits = 0
            .bSignal.Misses = 0
            .zScore.Hits = 0
            .zScore.Misses = 0
            .zScore.V = 0.0
            .zScore.min = 0.0
            .zScore.max = 0.0
            .Q.avg = 0.0
            .Trades = 0
            .bTrigger.tot = 0.0
            .bTrigger.Hits = 0
            .bTrigger.Misses = 0
            .bEntryStats.tot = 0.0
            .bEntryStats.Hits = 0
            .bEntryStats.Misses = 0
            .sEntryStats.tot = 0.0
            .sEntryStats.Hits = 0
            .sEntryStats.Misses = 0
            .sSignal.tot = 0.0
            .sSignal.Hits = 0
            .sSignal.Misses = 0
            .sTD1.tot = 0.0
            .sTD1.Hits = 0
            .sTD1.Misses = 0
            .Winners.tot = 0.0
            .Winners.avg = 0.0
            .Losers.tot = 0.0
            .Losers.avg = 0.0
            .drawDown.tot = 0
            .drawUp.tot = 0
            R.posOutliers = 0
            R.negOutliers = 0
            .zScore.V = 0.0
            .drawDown.avg = 0.0
            .drawDown.tot = 0.0
            .drawDown.min = 0.0
            .drawDown.max = 0.0
            .drawUp.avg = 0.0
            .drawUp.tot = 0.0
            .drawUp.min = 0.0
            .drawUp.max = 0.0
        End With
    End Sub
    Public Sub DoBHTrades_OpnCl(ByRef R As Results, ByRef TS() As Trades)
        Static xline As Integer
        Call InitR(R)
        With R
            .SP.bDOW.Text = "b_h"
            .SP.BandH.Text = "bhOpnCl"
            .SP.bMaxDH.V = 1
            .SP.bSignal.Text = "bh_Op"
            .SP.sSignal.Text = "bhnCl"
            .SP.bZScoreMode.Text = "0Z--"
        End With
        For xline = 1 To Counters.Days - 1
            R.SS.Main.Trades = R.SS.Main.Trades + 1
            TS(xline).bDOW.qHit = True
            TS(xline).bSignal.qHit = True
            TS(xline).bEntry.qHit = True
            R.SS.Main.bDOW.Hits = R.SS.Main.bDOW.Hits + 1
            R.SS.Main.bSignal.Hits = R.SS.Main.bSignal.Hits + 1
            R.SS.Main.bEntryStats.Hits = R.SS.Main.bEntryStats.Hits + 1
            TS(R.SS.Main.Trades).bDayNo = xline
            TS(R.SS.Main.Trades).sDayNo = xline + 1
            TS(R.SS.Main.Trades).bPrice = op(TS(R.SS.Main.Trades).bDayNo)
            TS(R.SS.Main.Trades).sPrice = cl(TS(R.SS.Main.Trades).sDayNo)
        Next xline
        lastTrade = TS(R.SS.Main.Trades)
    End Sub
    Public Sub DoBHTrades_ClnOp(ByRef R As Results, ByRef TS() As Trades)
        Static xline As Integer
        Call InitR(R)
        With R
            .SP.bDOW.Text = "b_h"
            .SP.BandH.Text = "bhClnOp"
            .SP.bMaxDH.V = 1
            .SP.bSignal.Text = "bh_Cl"
            .SP.sSignal.Text = "bhnOp"
            .SP.bZScoreMode.Text = "0Z--"
        End With
        For xline = 1 To Counters.Days - 1
            R.SS.Main.Trades = R.SS.Main.Trades + 1
            TS(xline).bDOW.qHit = True
            TS(xline).bSignal.qHit = True
            TS(xline).bEntry.qHit = True
            R.SS.Main.bDOW.Hits = R.SS.Main.bDOW.Hits + 1
            R.SS.Main.bSignal.Hits = R.SS.Main.bSignal.Hits + 1
            R.SS.Main.bEntryStats.Hits = R.SS.Main.bEntryStats.Hits + 1
            TS(R.SS.Main.Trades).bDayNo = xline
            TS(R.SS.Main.Trades).sDayNo = xline + 1
            TS(R.SS.Main.Trades).bPrice = cl(TS(R.SS.Main.Trades).bDayNo)
            TS(R.SS.Main.Trades).sPrice = op(TS(R.SS.Main.Trades).sDayNo)
        Next xline
        lastTrade = TS(R.SS.Main.Trades)
    End Sub
    Public Sub DoBHTrades_ClnCl(ByRef R As Results, ByRef TS() As Trades)
        Static xline As Integer
        Call InitR(R)
        With R
            .SP.bDOW.Text = "b_h"
            .SP.BandH.Text = "bhClnCl"
            .SP.bMaxDH.V = 1
            .SP.bSignal.Text = "bh_Cl"
            .SP.bEntry.Text = "bh_Cl"
            .SP.sSignal.Text = "bhnCl"
            .SP.sEntry.Text = "bhnCl"
            .SP.bZScoreMode.Text = "0Z--"
        End With
        For xline = 1 To Counters.Days - 1
            R.SS.Main.Trades = R.SS.Main.Trades + 1
            TS(xline).bDOW.qHit = True
            TS(xline).bSignal.qHit = True
            TS(xline).bEntryS.qHit = True
            R.SS.Main.bDOW.Hits = R.SS.Main.bDOW.Hits + 1
            R.SS.Main.bSignal.Hits = R.SS.Main.bSignal.Hits + 1
            R.SS.Main.bEntryStats.Hits = R.SS.Main.bEntryStats.Hits + 1
            TS(R.SS.Main.Trades).bDayNo = xline
            TS(R.SS.Main.Trades).sDayNo = xline + 1
            TS(R.SS.Main.Trades).bPrice = cl(TS(R.SS.Main.Trades).bDayNo)
            TS(R.SS.Main.Trades).sPrice = cl(TS(R.SS.Main.Trades).sDayNo)
        Next xline
        lastTrade = TS(R.SS.Main.Trades)
    End Sub
    Public Sub DoTradesMon(ByRef R As Results, ByRef TS() As Trades)
        Static xline As Integer
        Call InitR(R)
        R.SS.Main.bDOW.Hits = 0
        R.SS.Main.bDOW.Misses = 0
        For xline = 1 To Counters.Days
            R.SS.Main.Days = R.SS.Main.Days + 1
            If dowNo(xline) = 1 Then
                R.SS.Main.Trades = R.SS.Main.Trades + 1
                TS(xline).bDOW.qHit = True
                TS(xline).bSignal.qHit = True
                TS(xline).bEntry.qHit = True
                R.SS.Main.bDOW.Hits = R.SS.Main.bDOW.Hits + 1
                R.SS.Main.bSignal.Hits = R.SS.Main.bSignal.Hits + 1
                R.SS.Main.bEntryStats.Hits = R.SS.Main.bEntryStats.Hits + 1
                TS(R.SS.Main.Trades).bDayNo = xline
                TS(R.SS.Main.Trades).sDayNo = xline
                TS(R.SS.Main.Trades).bPrice = op(TS(R.SS.Main.Trades).bDayNo)
                TS(R.SS.Main.Trades).sPrice = cl(TS(R.SS.Main.Trades).sDayNo)
                '            Call doTrd(R)
            Else
                TS(xline).bDOW.qHit = False
                TS(xline).bSignal.qHit = False
                TS(xline).bEntry.qHit = False
                R.SS.Main.bDOW.Misses = R.SS.Main.bDOW.Misses + 1
                R.SS.Main.bSignal.Misses = R.SS.Main.bSignal.Misses + 1
                R.SS.Main.bEntryStats.Misses = R.SS.Main.bEntryStats.Misses + 1
            End If
        Next xline
        lastTrade = TS(R.SS.Main.Trades)
    End Sub
    Public Sub DoTradesTue(ByRef R As Results, ByRef TS() As Trades)
        Static xline As Integer
        Call InitR(R)
        For xline = 1 To Counters.Days
            R.SS.Main.Days = R.SS.Main.Days + 1
            If dowNo(xline) = 2 Then
                R.SS.Main.Trades = R.SS.Main.Trades + 1
                TS(xline).bDOW.qHit = True
                TS(xline).bSignal.qHit = True
                TS(xline).bEntry.qHit = True
                R.SS.Main.bDOW.Hits = R.SS.Main.bDOW.Hits + 1
                R.SS.Main.bSignal.Hits = R.SS.Main.bSignal.Hits + 1
                R.SS.Main.bEntryStats.Hits = R.SS.Main.bEntryStats.Hits + 1
                TS(R.SS.Main.Trades).bDayNo = xline
                TS(R.SS.Main.Trades).sDayNo = xline
                TS(R.SS.Main.Trades).bPrice = op(TS(R.SS.Main.Trades).bDayNo)
                TS(R.SS.Main.Trades).sPrice = cl(TS(R.SS.Main.Trades).sDayNo)
                '    Call doTrd(R)
            Else
                TS(xline).bDOW.qHit = False
                TS(xline).bSignal.qHit = False
                TS(xline).bEntry.qHit = False
                R.SS.Main.bDOW.Misses = R.SS.Main.bDOW.Misses + 1
                R.SS.Main.bSignal.Misses = R.SS.Main.bSignal.Misses + 1
                R.SS.Main.bEntryStats.Misses = R.SS.Main.bEntryStats.Misses + 1
            End If
        Next xline
        lastTrade = TS(R.SS.Main.Trades)
    End Sub
    Public Sub DoTradesWed(ByRef R As Results, ByRef TS() As Trades)
        Static xline As Integer
        Call InitR(R)
        For xline = 1 To Counters.Days
            R.SS.Main.Days = R.SS.Main.Days + 1
            If dowNo(xline) = 3 Then
                R.SS.Main.Trades = R.SS.Main.Trades + 1
                TS(xline).bDOW.qHit = True
                TS(xline).bSignal.qHit = True
                TS(xline).bEntry.qHit = True
                R.SS.Main.bDOW.Hits = R.SS.Main.bDOW.Hits + 1
                R.SS.Main.bSignal.Hits = R.SS.Main.bSignal.Hits + 1
                R.SS.Main.bEntryStats.Hits = R.SS.Main.bEntryStats.Hits + 1
                TS(R.SS.Main.Trades).bDayNo = xline
                TS(R.SS.Main.Trades).sDayNo = xline
                TS(R.SS.Main.Trades).bPrice = op(TS(R.SS.Main.Trades).bDayNo)
                TS(R.SS.Main.Trades).sPrice = cl(TS(R.SS.Main.Trades).sDayNo)
                '     Call doTrd(R)
            Else
                TS(xline).bDOW.qHit = False
                TS(xline).bSignal.qHit = False
                TS(xline).bEntry.qHit = False
                R.SS.Main.bDOW.Misses = R.SS.Main.bDOW.Misses + 1
                R.SS.Main.bSignal.Misses = R.SS.Main.bSignal.Misses + 1
                R.SS.Main.bEntryStats.Misses = R.SS.Main.bEntryStats.Misses + 1
            End If
        Next xline
        lastTrade = TS(R.SS.Main.Trades)
    End Sub
    Public Sub DoTradesThr(ByRef R As Results, ByRef TS() As Trades)
        Static xline As Integer
        Call InitR(R)
        For xline = 1 To Counters.Days
            R.SS.Main.Days = R.SS.Main.Days + 1
            If dowNo(xline) = 4 Then
                R.SS.Main.Trades = R.SS.Main.Trades + 1
                TS(xline).bDOW.qHit = True
                TS(xline).bSignal.qHit = True
                TS(xline).bEntry.qHit = True
                R.SS.Main.bDOW.Hits = R.SS.Main.bDOW.Hits + 1
                R.SS.Main.bSignal.Hits = R.SS.Main.bSignal.Hits + 1
                R.SS.Main.bEntryStats.Hits = R.SS.Main.bEntryStats.Hits + 1
                TS(R.SS.Main.Trades).bDayNo = xline
                TS(R.SS.Main.Trades).sDayNo = xline
                TS(R.SS.Main.Trades).bPrice = op(TS(R.SS.Main.Trades).bDayNo)
                TS(R.SS.Main.Trades).sPrice = cl(TS(R.SS.Main.Trades).sDayNo)
                '          Call doTrd(R)
            Else
                TS(xline).bDOW.qHit = False
                TS(xline).bSignal.qHit = False
                TS(xline).bEntry.qHit = False
                R.SS.Main.bDOW.Misses = R.SS.Main.bDOW.Misses + 1
                R.SS.Main.bSignal.Misses = R.SS.Main.bSignal.Misses + 1
                R.SS.Main.bEntryStats.Misses = R.SS.Main.bEntryStats.Misses + 1
            End If
        Next xline
        lastTrade = TS(R.SS.Main.Trades)
    End Sub
    Public Sub DoTradesFri(ByRef R As Results, ByRef TS() As Trades)
        Static xline As Integer
        Call InitR(R)
        For xline = 1 To Counters.Days
            R.SS.Main.Days = R.SS.Main.Days + 1
            If dowNo(xline) = 5 Then
                R.SS.Main.Trades = R.SS.Main.Trades + 1
                TS(xline).bDOW.qHit = True
                TS(xline).bSignal.qHit = True
                TS(xline).bEntry.qHit = True
                R.SS.Main.bDOW.Hits = R.SS.Main.bDOW.Hits + 1
                R.SS.Main.bSignal.Hits = R.SS.Main.bSignal.Hits + 1
                R.SS.Main.bEntryStats.Hits = R.SS.Main.bEntryStats.Hits + 1
                TS(R.SS.Main.Trades).bDayNo = xline
                TS(R.SS.Main.Trades).sDayNo = xline
                TS(R.SS.Main.Trades).bPrice = op(TS(R.SS.Main.Trades).bDayNo)
                TS(R.SS.Main.Trades).sPrice = cl(TS(R.SS.Main.Trades).sDayNo)
                '           Call doTrd(R)
            Else
                TS(xline).bDOW.qHit = False
                TS(xline).bSignal.qHit = False
                TS(xline).bEntry.qHit = False
                R.SS.Main.bDOW.Misses = R.SS.Main.bDOW.Misses + 1
                R.SS.Main.bSignal.Misses = R.SS.Main.bSignal.Misses + 1
                R.SS.Main.bEntryStats.Misses = R.SS.Main.bEntryStats.Misses + 1
            End If
        Next xline
        lastTrade = TS(R.SS.Main.Trades)
    End Sub
    Private Sub PreLoop(ByRef r As Results)
        Counters.Days = 0
        newTrade.TPHits = 0
        newTrade.TPMisses = 0
        With r
            .SS.Main.Trades = 0
            .bhSPr = op(Counters.start_Day)
            .bhEPr = op(Counters.end_day)
            .Days = Counters.end_day - Counters.start_Day + 1
            .bhGrossPr = .bhEPr - .bhSPr
            .bhShares = 1000 / .bhSPr
            .bhProfit = .bhGrossPr * .bhShares
            .bhProfitDay = .bhProfit / .Days
            .bhQ = .bhProfitDay / 10
            .bhAveProfit = .bhProfit / .Days
        End With
        '        With Form1
        '        .Iterations.Text = Format$(Counters.Iteration, "000000")
        '       '          .Iters.Text = Format$(Counters.Iteration, "000000")
        '        .threshholdTxt.Text = Format$(Counters.threshHold, "0.000")
        '        .startDayDate.Text = Counters.start_Day & "--" & Counters.start_Date
        '       '            .endDayDate.Text = Counters.end_day & "--" & Counters.end_Date & ":" & DOWStr(Counters.end_day)
        '       '          DataGridView1.Columns(2).Width = 150
        '        .seconds.Text = DateDiff(DateInterval.Second, Now.Date, Now)
        '        End With
        Call InitR(r)
    End Sub
    Public Sub Do_LoopsBuyExecuteTrigger(ByRef thisDay As Integer)
        Static oldTxt As String, bt As Integer, trig As Single, be As Integer
        QFE1.Message.Text = "looping " & Format$(Counters.Iteration, "000000")
        With RMain
            For bt = 0 To QFE1.buyTrigger.Items.Count - 1
                .SP.bTrigger.Idx = bt
                oldTxt = Strings.Left(QFE1.buyTrigger.Items(.SP.bTrigger.Idx), 10)
                trig = TrigPrice_(.SP.bTrigger.Idx, Counters.end_day)
                QFE1.buyTrigger.Items(.SP.bTrigger.Idx) = Strings.Left(oldTxt, 10) & "@" & Format$(trig, "000.000")
                QFE1.TxtBuyTrigger.BackColor = SystemColors.ActiveCaption
                .SP.bTrigger.Text = oldTxt
                .SP.bTrigger.Text1 = (QFE1.buyTrigger.Items(.SP.bTrigger.Idx).ToString)
                QFE1.buyTriggerTxt.Text = .SP.bTrigger.Text
                QFE1.buyTrigger.SelectedIndex = .SP.bTrigger.Idx
                If QFE1.buyTrigger.GetItemChecked(.SP.bTrigger.Idx) Then
                    .SP.bTrigger.Text = Strings.Left(QFE1.buyTrigger.Items(.SP.bTrigger.Idx).ToString, 6)
                    .SP.bTrigger.Text1 = .SP.bTrigger.Text
                    QFE1.buyTriggerTxt.Text = .SP.bTrigger.Text
                    For be = 0 To QFE1.buyEntry.Items.Count - 1
                        .SP.bEntry.Idx = be
                        If QFE1.buyEntry.GetItemChecked(.SP.bEntry.Idx) Then
                            QFE1.buyEntry.SelectedIndex = .SP.bEntry.Idx
                            .SP.bEntry.Text = Strings.Left(QFE1.buyEntry.Items(.SP.bEntry.Idx).ToString, 5)
                            .SP.bEntry.Text1 = Strings.Left(QFE1.buyEntry.Items(.SP.bEntry.Idx).ToString, 5)
                            '                            Application.DoEvents()
                            QFE1.buyEntryTxt.Text = .SP.bEntry.Text
                            QFE1.buyExecuteText_.Text = .SP.bEntry.Text
                            Call SellEntry1_(thisDay)
                        End If
                    Next be
                End If
            Next bt
        End With
        QFE1.seconds.Text = DateDiff(DateInterval.Second, Now.Date, Now)
    End Sub
    Public Sub SellEntry1_(ByRef thisDay As Integer)
        With RMain.SP
            For RMain.SP.sEntry.Idx = 0 To QFE1.sellEntry.Items.Count - 1
                If QFE1.sellEntry.GetItemChecked(.sEntry.Idx) Then
                    QFE1.sellEntry.SelectedIndex = RMain.SP.sEntry.Idx
                    .sEntry.Text = Strings.Left(QFE1.sellEntry.Items(.sEntry.Idx).ToString, 3)
                    .sEntry.Text1 = Strings.Left(QFE1.sellEntry.Items(.sEntry.Idx).ToString, 3)
                    QFE1.txtSEntry.Text = .sEntry.Text
                    QFE1.sellEntryTxt_.Text = .sEntry.Text
                    '                                    If .SP.bentry.Idx = 1 And .SP.sExecute.Idx = 1 Then GoTo skippp
                    '                                    If .SP.bentry.Idx = 2 And .SP.sExecute.Idx = 1 Then GoTo skippp
                    '                                    If .SP.bentry.Idx = 2 And .SP.sExecute.Idx = 2 Then GoTo skippp
                    '                                   If .SP.bentry.Idx = 3 And .SP.sExecute.Idx = 1 Then GoTo skippp
                    '                                   If .SP.bentry.Idx = 3 And .SP.sExecute.Idx = 3 Then GoTo skippp
                    '                                If r.SP.bentry.Idx = 3 And r.SP.sExecute.Idx = 2 Then GoTo skippp
                    '                                   If .SP.bentry.Idx = 4 And .SP.sExecute.Idx = 1 Then GoTo skippp
                    '                                   If .SP.bentry.Idx = 4 And .SP.sExecute.Idx = 2 Then GoTo skippp
                    Call TDLoop1_(thisDay)
skippp:
                End If
            Next
        End With
    End Sub
    Public Function TrigPrice_(trigIdx As Integer, ByRef dDay As Integer) As Single
        Select Case trigIdx
            Case 0
                TrigPrice_ = 0.0
            Case 1
                TrigPrice_ = op(dDay)
            Case 2
                TrigPrice_ = (op(dDay) + hi(dDay)) / 2
            Case 3
                TrigPrice_ = (op(dDay) + lo(dDay)) / 2
            Case 4
                TrigPrice_ = (op(dDay) + cl(dDay)) / 2
            Case 5
                TrigPrice_ = hi(dDay)
            Case 6
                TrigPrice_ = (hi(dDay) + lo(dDay)) / 2
            Case 7
                TrigPrice_ = (hi(dDay) + cl(dDay)) / 2
            Case 8
                TrigPrice_ = lo(dDay)
            Case 9
                TrigPrice_ = (lo(dDay) + cl(dDay)) / 2
            Case 10
                TrigPrice_ = cl(dDay)
            Case 11
                TrigPrice_ = op(dDay - 1)
            Case 12
                TrigPrice_ = (op(dDay - 1) + hi(dDay - 1)) / 2
            Case 13
                TrigPrice_ = (op(dDay - 1) + lo(dDay - 1)) / 2
            Case 14
                TrigPrice_ = (op(dDay - 1) + cl(dDay - 1)) / 2
            Case 15
                TrigPrice_ = hi(dDay - 1)
            Case 16
                TrigPrice_ = (hi(dDay - 1) + lo(dDay - 1)) / 2
            Case 17
                TrigPrice_ = (hi(dDay - 1) + cl(dDay - 1)) / 2
            Case 18
                TrigPrice_ = lo(dDay - 1)
            Case 19
                TrigPrice_ = (lo(dDay - 1) + cl(dDay - 1)) / 2
            Case 20
                TrigPrice_ = cl(dDay - 1)
            Case Else
                Stop
                TrigPrice_ = -1.0
        End Select
    End Function
    Private Sub TDLoop1_(ByRef thisDay As Integer)
        With RMain.SP.bTD1
            For RMain.SP.bTD1.dbIdx = 0 To QFE1.buyTDDaysBack1.Items.Count - 1
                If QFE1.buyTDDaysBack1.GetItemChecked(RMain.SP.bTD1.dbIdx) Then
                    RMain.SP.bTD1.dbText = QFE1.buyTDDaysBack1.Items(RMain.SP.bTD1.dbIdx)
                    RMain.SP.bTD1.dbV = Val(RMain.SP.bTD1.dbText)
                    QFE1.buyTDDaysBack1.SelectedIndex = RMain.SP.bTD1.dbIdx
                    For RMain.SP.bTD1.Idx = 0 To QFE1.buyTDSignal1.Items.Count - 1
                        If QFE1.buyTDSignal1.GetItemChecked(RMain.SP.bTD1.Idx) Then
                            RMain.SP.bTD1.Text = Strings.Left(QFE1.buyTDSignal1.Items(RMain.SP.bTD1.Idx), 9)
                            If .Idx = 0 Then
                                .Text = .Text & "___"
                            Else
                                RMain.SP.bTD1.Text = .Text & .dbText & "]"
                            End If
                            RMain.SP.bTD1.Text1 = RMain.SP.bTD1.Text & ":" & Format$(RMain.SP.bTD1.dbV, "00")
                            QFE1.bTDtxt1.Text = RMain.SP.bTD1.Text1
                            QFE1.buyTD1txt.Text = RMain.SP.bTD1.Text1
                            If RMain.SP.bTD1.Idx = 0 And RMain.SP.bTD1.dbIdx > 0 Then GoTo skipp
                            QFE1.buyTDSignal1.SelectedIndex = .Idx
                            lastDayTrade.bTD1.Idx = .Idx
                            lastDayTrade.bTD1.qHit = QTD_(lastDayTrade.bTD1, Counters.end_day, RMain.SP.bTD1.dbV)
                            If lastDayTrade.bTD1.qHit Then
                                Call TDLoop2(thisDay)
                            Else
                                Counters.Off = Counters.Off + 1
                            End If

skipp:
                        End If
                        QFE1.bTD1_.Text = .Text
                    Next .Idx
                End If
            Next .dbIdx
        End With
    End Sub
    Private Sub TDLoop2(ByRef thisDay As Integer)
        With RMain.SP.bTD2
            For RMain.SP.bTD2.dbIdx = 0 To QFE1.buyTDDaysBack2.Items.Count - 1
                If QFE1.buyTDDaysBack2.GetItemChecked(RMain.SP.bTD2.dbIdx) Then
                    RMain.SP.bTD2.dbText = QFE1.buyTDDaysBack2.Items(.dbIdx)
                    RMain.SP.bTD2.dbV = Val(RMain.SP.bTD2.dbText)
                    QFE1.buyTDDaysBack2.SelectedIndex = RMain.SP.bTD2.dbIdx
                    For RMain.SP.bTD2.Idx = 0 To QFE1.buyTDSignal2.Items.Count - 1
                        If QFE1.buyTDSignal2.GetItemChecked(.Idx) Then
                            RMain.SP.bTD2.Text = Strings.Left(QFE1.buyTDSignal2.Items(.Idx), 9)
                            If RMain.SP.bTD2.Idx = 0 Then
                                RMain.SP.bTD2.Text = RMain.SP.bTD2.Text & "___"
                            Else
                                RMain.SP.bTD2.Text = RMain.SP.bTD2.Text & .dbText & "]"
                            End If
                            RMain.SP.bTD2.Text1 = RMain.SP.bTD2.Text & ":" & Format$(RMain.SP.bTD2.dbV, "00")
                            QFE1.bTDtxt2.Text = RMain.SP.bTD2.Text1
                            QFE1.buyTD2txt.Text = RMain.SP.bTD2.Text1
                            If RMain.SP.bTD2.Idx = 0 And RMain.SP.bTD2.dbIdx > 0 Then GoTo skipp
                            QFE1.buyTDSignal2.SelectedIndex = RMain.SP.bTD2.Idx
                            lastDayTrade.bTD2.Idx = RMain.SP.bTD2.Idx
                            lastDayTrade.bTD2.qHit = QTD_(lastDayTrade.bTD2, thisDay, RMain.SP.bTD2.dbV)
                            If lastDayTrade.bTD2.qHit Then
                                Call Max(thisDay)
                            Else
                                Counters.Off = Counters.Off + 1
                            End If
skipp:
                        End If
skiptd2:
                        QFE1.bTD2_.Text = .Text1
                    Next .Idx
                End If
            Next .dbIdx
        End With
    End Sub
    Private Sub Max(ByRef thisDay As Integer)
        Static xmax As Integer
        With RMain.SP
            For RMain.SP.TP.Idx = 0 To QFE1.targetProfit.Items.Count - 1
                If QFE1.targetProfit.GetItemChecked(RMain.SP.TP.Idx) Then
                    QFE1.targetProfit.SelectedIndex = RMain.SP.TP.Idx
                    .TP.Text = QFE1.targetProfit.Items(.TP.Idx)
                    .TP.V = Val(.TP.Text)
                    QFE1.TP.Text = .TP.Text
                    For xmax = 0 To QFE1.sellMaxDH.Items.Count - 1
                        .sMaxDH.Idx = xmax
                        If QFE1.sellMaxDH.GetItemChecked(.sMaxDH.Idx) Then
                            QFE1.sellMaxDH.SelectedIndex = .sMaxDH.Idx
                            .sMaxDH.Text = QFE1.sellMaxDH.Items(.sMaxDH.Idx)
                            RMain.SP.sMaxDH.V = Val(.sMaxDH.Text)
                            QFE1.maxDH_.Text = RMain.SP.sMaxDH.Text
                            QFE1.maxDHText_.Text = RMain.SP.sMaxDH.Text
                            '                         If Not (RMain.SP.sMaxDH.V = 0 And RMain.SP.TP.Idx) > 0 Then
                            If .bEntry.Idx = 2 And .sEntry.Idx = 2 And .sMaxDH.V = 0 Then GoTo skipp
                                If .bEntry.Idx = 4 And .sEntry.Idx = 4 And .sMaxDH.Idx = 0 Then GoTo skipp
                                '                           If .bEntry.Idx = 4 And .sEntry.Idx = 3 And .sMaxDH.Idx = 0 Then GoTo skipp
                                '                                    If r.SP.sMaxDH.V = 0 And r.SP.bEntry.Idx = 4 Then GoTo skipp
                                If .sMaxDH.Idx = 0 And .sTD1.Idx >= 1 Then GoTo skipp
                                '                                            If .sMaxDH.Idx = 0 And (.ssignal.Idx = 1 Or .ssignal.Idx = 2) Then GoTo skipp
                                If .sMaxDH.Idx = 0 And .sTD1.Idx >= 1 Then GoTo skipp
                                '                   If .bTD3.Idx = 0 And .bTD3.dbIdx > 0 Then GoTo skipp
                                '                                   If .bTD4.Idx = 0 And .bTD4.dbIdx > 0 Then GoTo skipp
                                '                                  If .bTD5.Idx = 0 And .bTD5.dbIdx > 0 Then GoTo skipp
                                Application.DoEvents()
                            Call ByZScoreMode_Loop(thisDay)
skipp:
                                QFE1.maxDH_.Text = RMain.SP.sMaxDH.Text
                                'End If
                            End If
                    Next xmax
                End If
            Next RMain.SP.TP.Idx
        End With
    End Sub
    Public Sub ByDOW(ByRef thisDay As Integer)
        Static xx As Integer
        With RMain.SP
            For RMain.SP.bMonth.Idx = 0 To QFE1.buyMonth.Items.Count - 1
                If QFE1.buyMonth.GetItemChecked(.bMonth.Idx) Then
                    .bMonth.Text = Strings.Left(QFE1.buyMonth.Items(.bMonth.Idx).ToString, 8)
                    QFE1.txtBuyMonth.Text = .bMonth.Text
                    QFE1.buyMonth.SelectedIndex = .bMonth.Idx
                    For xx = 0 To QFE1.buyDOW.Items.Count - 2
                        RMain.SP.bDOW.Idx = xx
                        .bDOW.Text = Strings.Left(QFE1.buyDOW.Items(.bDOW.Idx).ToString, 4)
                        QFE1.DOW.Text = .bDOW.Text1
                        QFE1.buyDOW.SelectedIndex = .bDOW.Idx
                        QFE1.txtBuyDOW.Text = .bDOW.Text
                        QFE1.whichDates.Items(Counters.xDayIdx) = QFE1.whichDates.Items(Counters.xDayIdx) & " " & DOWStr(Counters.last_day1)
                        Application.DoEvents()
                        If QFE1.buyDOW.GetItemChecked(RMain.SP.bDOW.Idx) Then
                            .bDOW.Text1 = .bDOW.Text & ":" & dtStr1(thisDay)
                            QFE1.buyDowTxt.Text = RMain.SP.bDOW.Text1
                            If (.bDOW.Idx = 0) Then
                                'Stop
                                Call Do_InsideDay(thisDay)
                            Else
                                If QExeBDOW(thisDay) Then
                                    Call Do_InsideDay(thisDay)
                                End If
                            End If
                        End If
                    Next xx
                End If
            Next
        End With
    End Sub
    Private Sub Do_InsideDay(ByRef thisDay As Integer)
        Static q As Boolean
        With RMain.SP
            For .bInday.Idx = 0 To QFE1.BuyInDay.Items.Count - 1
                RMain.SP.bInday.Text = Strings.Left(QFE1.BuyInDay.Items(.bInday.Idx).ToString, 17)
                .bInday.qP = .bInday.Idx > 0
                If QFE1.BuyInDay.GetItemChecked(.bInday.Idx) Then
                    .bInday.Text1 = .bInday.Text & ":" & dtStr1(thisDay)
                    QFE1.BuyInDay.SelectedIndex = .bInday.Idx
                    QFE1.buyInDayText.Text = RMain.SP.bInday.Text
                    '                   Form1.buyInDayText.Text = Strings.Left(Form1.BuyInDay.Items(.bInday.Idx).ToString, 8)
                    q = QbuyInDay(thisDay)
                    If q Then
                        Counters.Onn = Counters.Onn + 1
                        Call Do_LoopsBuyExecuteTrigger(thisDay)
                        QFE1.On_.Text = Format$(Counters.Onn, "000000")
                    Else
                        Counters.Off = Counters.Off + 1
                        QFE1.Off_.Text = Format$(Counters.Off, "000000")
                    End If
                End If
                QFE1.InsideDay_.Text = .bInday.Text
            Next .bInday.Idx
        End With
    End Sub
    Private Sub CalcZScoreEachTrade(ByRef R As Results,
                                     ByRef TS() As Trades)
        Dim rtmp As Single, trd As Integer
        R.SS.Main.zScore.min = 1000.0
        R.SS.Main.zScore.max = -1000.0
        rtmp = 0.0
        For trd = 22 To R.SS.Main.Trades
            R.SS.Main.zScore.V = TS(trd).bZscoreMode.V
            If R.SS.Main.zScore.V > R.SS.Main.zScore.max Then
                R.SS.Main.zScore.max = R.SS.Main.zScore.V
            End If
            If R.SS.Main.zScore.V < R.SS.Main.zScore.min Then
                R.SS.Main.zScore.min = R.SS.Main.zScore.V
            End If
        Next trd
    End Sub
    Private Sub ZModeSetMinMax(ByRef R As Results)
        With R.SP.bZScoreValue
            '         Form1.txtBZScoreStd.Text = .Text1
            '          .Text = Form1.ZScoreStd.Items(.Idx)
            '           .Text1 = Form1.ZScoreStd.Items(.Idx)
            Select Case .Idx
                Case 0
                    .min = -1000.0
                    .max = +1000.0
                Case 1
                    .min = -1000.0
                    .max = -3.0
                Case 2
                    .min = -3.0
                    .max = -2.5
                Case 3
                    .min = -2.5
                    .max = -2.0
                Case 4
                    .min = -2.0
                    .max = -1.5
                Case 5
                    .min = -1.5
                    .max = -1.0
                Case 6
                    .min = -1.0
                    .max = -0.5
                Case 7
                    .min = -0.5
                    .max = -0.0
                Case 8
                    .min = +0.0
                    .max = +0.5
                Case 9
                    .min = +0.5
                    .max = +1.0
                Case 10
                    .min = +1.0
                    .max = +1.5
                Case 11
                    .min = +1.5
                    .max = +2.0
                Case 12
                    .min = +2.0
                    .max = +2.5
                Case 13
                    .min = +2.5
                    .max = +3.0
                Case 14
                    .min = +3.0
                    .max = +1000.0
                Case Else
                    Stop
            End Select
        End With
    End Sub
    Public Function QbuyInDay(ByRef tday As Integer) As Boolean
        Static q As Boolean
        q = False
        Select Case RMain.SP.bInday.Idx
            Case 0
                q = True
            Case 1
                q = (hi(tday) < hi(tday - 1) And lo(tday) > lo(tday - 1))
            Case 2
                q = Not ((hi(tday) < hi(tday - 1) And lo(tday) > lo(tday - 1)))
            Case 3
                q = (hi(tday) < hi(tday - 2) And lo(tday) > lo(tday - 2))
            Case 4
                q = Not (hi(tday) < hi(tday - 2) And lo(tday) > lo(tday - 2))
            Case 5
                q = (hi(tday) < hi(tday - 3) And lo(tday) > lo(tday - 3))
            Case 6
                q = Not (hi(tday) < hi(tday - 3) And lo(tday) > lo(tday - 3))
            Case 7
                q = (hi(tday) < hi(tday - 4) And lo(tday) > lo(tday - 4))
            Case 8
                q = Not (hi(tday) < hi(tday - 4) And lo(tday) > lo(tday - 4))
            Case Else
                Stop
        End Select
        QbuyInDay = q
        '       If Not q Then Stop
    End Function
    Private Sub Calc_HitsMisses(ByRef R As Results)
        Dim strr As String
        Static q As Boolean, q1 As Boolean, q2 As Boolean, tDay As Integer, qd As Boolean
        RHitMisses.SP = R.SP
        With RHitMisses.SS.Main
            .bMonth.Hits = 0
            .bMonth.Misses = 0
            .bMonth.tot = 0
            .bDOW.Pcntg = 0.0
            .bDOW.Hits = 0
            .bDOW.Misses = 0
            .bDOW.tot = 0
            .bDOW.Pcntg = 0.0
            .zScore.V = 0
            .zScore.Hits = 0
            .zScore.Misses = 0
            .zScore.tot = 0
            .zScoreValue.Hits = 0
            .zScoreValue.Misses = 0
            .bInDay.tot = 0
            .bInDay.Hits = 0
            .bInDay.Misses = 0
            .bInDay.Pcntg = 0.0
            .bTD0.tot = 0
            .bTD0.Hits = 0
            .bTD0.Misses = 0
            .bDOW.Pcntg = 0.0
            .bTD1.tot = 0
            .bTD1.Hits = 0
            .bTD1.Misses = 0
            .bTD1.Pcntg = 0.0
            .bTD2.tot = 0
            .bTD2.Hits = 0
            .bTD2.Misses = 0
            .bTD2.Pcntg = 0.0
            .bTrigger.tot = 0
            .bTrigger.Hits = 0
            .bTrigger.Misses = 0
            .bTrigger.Pcntg = 0.0
            For tDay = Counters.start_Day To Counters.end_day
                qd = QExeBDOW(tDay)
                If qd Then
                    .bDOW.Hits = .bDOW.Hits + 1
                Else
                    .bDOW.Misses = .bDOW.Misses + 1
                End If
                q = QbuyInDay(tDay)
                If q Then
                    .bInDay.Hits = RHitMisses.SS.Main.bInDay.Hits + 1
                Else
                    .bInDay.Misses = .bInDay.Misses + 1
                End If
                Call CountTD(tDay)
                q = newTrade.bTD1.qHit And newTrade.bTD2.qHit
                If q Then
                    .bTD0.Hits = .bTD0.Hits + 1
                Else
                    .bTD0.Misses = .bTD0.Misses + 1
                End If
                q1 = newTrade.bTD1.qHit
                If q1 Then
                    .bTD1.Hits = .bTD1.Hits + 1
                Else
                    .bTD1.Misses = .bTD1.Misses + 1
                End If
                q2 = newTrade.bTD2.qHit
                If q2 Then
                    .bTD2.Hits = .bTD2.Hits + 1
                Else
                    .bTD2.Misses = .bTD2.Misses + 1
                End If
                '                q = qHitPatternZ(newTrade)
                q = QBuyTrigger(newTrade, tDay)
                If q Then
                    .bTrigger.Hits = .bTrigger.Hits + 1
                Else
                    .bTrigger.Misses = .bTrigger.Misses + 1
                End If
                q = q
            Next tDay
            QFE1.buyDOW.Items(RHitMisses.SP.bDOW.Idx) =
         Strings.Left(QFE1.buyDOW.Items(RHitMisses.SP.bDOW.Idx), 7) & ":" &
          Format(RHitMisses.SS.Main.bDOW.Hits, "0000")
            strr =
            Strings.Left(QFE1.BuyInDay.Items(RHitMisses.SP.bInday.Idx), 9) & ":::" &
            Format(RHitMisses.SS.Main.bInDay.Hits, "0000")
            '            Form1.BuyInDay.Items(RHitMisses.SP.bInday.Idx) = strr
            QFE1.buyInDayText.Text = strr
        End With
        Application.DoEvents()
    End Sub
    REM
    Public Sub display(ByRef R As Results, ByRef TS() As Trades, q As Boolean)
        Static xTrade As Integer, tmptxt As String
        R.SP.bZScoreMode.Text1 = "td=" & Format$(R.SS.Main.Trades, "0000") &
                                  "q=" & StrSign(R.SS.Main.Q.avg) &
                                   "w%=" & Format$(R.SS.Main.wPcntg, "0.00") &
                                    "w=" & Format$(R.SS.Main.W, "0000") &
                                     "l=" & Format$(R.SS.Main.L, "0000") &
                                      "zScr=" & StrSign(R.SS.Main.zScore.V)
        If q Then
            For xTrade = 1 To R.SS.Main.Trades
                Call insertTrade(R, TS(xTrade))
            Next xTrade
            tmptxt = Strings.Left(QFE1.buyZScoreMode.Items(R.SP.bZScoreMode.Idx), 8) & R.SP.bZScoreMode.Text1
            QFE1.buyZScoreMode.Items(R.SP.bZScoreMode.Idx) = tmptxt
            current_Seconds = DateDiff(DateInterval.Second, Now.Date, Now)
            elapsed_Seconds = current_Seconds - start_Seconds
            iterspersecond_ = Counters.Iteration / elapsed_Seconds
            QFE1.iterpersecond.Text = Format$(iterspersecond_, "00.000")
        End If
    End Sub
    Public Sub ByZScoreMode_Loop(ByRef thisDay As Integer)
        Dim t0 As Integer, t1 As Integer, t2 As Integer, t3 As Integer, t4 As Integer
        Dim tmptxt0 As String, tmptxt1 As String, tmptxt2 As String, tmptxt3 As String, tmptxt4 As String
        Counters.ZIterBase = Counters.ZIterBase + 1
        Counters.ZIterNo = 0
        ReDim RMainTS(0)
        ReDim RMainTS(Counters.end_day)
        RMain.SP.bZScoreMode.Idx = 0
        Counters.totalIterations = Counters.totalIterations + 1
        Call PreLoop(RMain)
        RMain.SP.bZScoreMode.Text = QFE1.buyZScoreMode.Items(RZA__.SP.bZScoreMode.Idx) 'buyZScore_.Items(RMain.SP.bZScoremode.Idx)
        RMain.SS.Main.Trades = BsDayLoopCase00()
        RZA__ = RMain
        If RMain.SS.Main.Trades > 13 Then
            RZA__TS = RMainTS
            Call ZModeSetMinMax(RZA__)
            QFE1.BuyZScoreModetxt.Text = RMain.SP.bZScoreMode.Text
            RZA__.SP.bZScoreMode.Text = Strings.Left(QFE1.buyZScoreMode.Items(RZA__.SP.bZScoreMode.Idx), 8)
            t0 = RZA__.SS.Main.Trades
            RZA__.SS.Main.zScoreMode.V = CalcZScore(RZA__, RZA__TS, 2, RZA__.SS.Main.Trades)
            '            RZA__.SP.bZScoreMode.Text1 = "td=" & Format$(RZA__.SS.Main.Trades, "0000") &
            '                              "q=" & StrSign(RZA__.SS.Main.Q.avg) &
            '                              "w%=" & Format$(RZA__.SS.Main.wPcntg, "0.00") &
            '                              "w=" & Format$(RZA__.SS.Main.W, "0000") &
            '                               "l=" & Format$(RZA__.SS.Main.L, "0000") &
            '                               "zScr=" & StrSign(RZA__.SS.Main.zScore.V)
            tmptxt0 = Strings.Left(QFE1.buyZScoreMode.Items(RZA__.SP.bZScoreMode.Idx), 8) & RZA__.SP.bZScoreMode.Text1
            QFE1.buyZScoreMode.Items(RZA__.SP.bZScoreMode.Idx) = tmptxt0
            QFE1.BZScrVal_.Text = RZA__.SP.bZScoreValue.Text1
            Call byZScoreValue_Loop(RZA__, RZA__TS, thisDay)
            Counters.Onn = Counters.Onn + 1
            QFE1.On_.Text = Format(Counters.Onn, "0000000")
            Call display(RZA__, RZA__TS, QFE1.qSaveTrades0.Checked)
        Else
            Counters.Off = Counters.Off + 1
            QFE1.Off_.Text = Format$(Counters.Off, "0000000")
        End If
        For RMain.SP.bZScoreMode.Idx = 1 To QFE1.buyZScoreMode.Items.Count - 1
            QFE1.buyZScoreMode.SelectedIndex = RMain.SP.bZScoreMode.Idx
            Select Case RMain.SP.bZScoreMode.Idx
                Case 0
                    Stop
                Case 1
                    If QFE1.buyZScoreMode.GetItemChecked(RMain.SP.bZScoreMode.Idx) Then
                        ReDim RZP__TS(0)
                        ReDim RZP__TS(Counters.end_day)
                        Counters.totalIterations = Counters.totalIterations + 1
                        Call PreLoop(RzP__)
                        RzP__.SP = RMain.SP
                        RzP__.SP.bZScoreMode.Text = Strings.Left(QFE1.buyZScoreMode.Items(RzP__.SP.bZScoreMode.Idx), 8)
                        Call byZScoreValue_Loop(RzP__, RZP__TS, thisDay)
                        t1 = RzP__.SS.Main.Trades
                        tmptxt1 = Strings.Left(QFE1.buyZScoreMode.Items(RzP__.SP.bZScoreMode.Idx), 8) & RzP__.SP.bZScoreMode.Text1
                        RzP__.SP.bZScoreValue.Text1 = tmptxt1
                        QFE1.txtZBP____.Text = Format(RzP__.SS.Main.Trades, "00000")
                        QFE1.On_.Text = Format(Counters.Onn, "0000000")
                        Counters.Onn = Counters.Onn + 1
                        Call display(RzP__, RZP__TS, QFE1.qSaveTrades0.Checked)
                    Else
                        Counters.Off = Counters.Off + 1
                        QFE1.Off_.Text = Format$(Counters.Off, "0000000")
                    End If
                Case 2
                    If QFE1.buyZScoreMode.GetItemChecked(RMain.SP.bZScoreMode.Idx) Then
                        Counters.totalIterations = Counters.totalIterations + 1
                        Call PreLoop(RzN__)
                        RzN__.SP = RMain.SP
                        RzN__.SP.bZScoreMode.Text = Strings.Left(QFE1.buyZScoreMode.Items(RzN__.SP.bZScoreMode.Idx), 8)
                        Call byZScoreValue_Loop(RzN__, RZN__TS, thisDay)
                        t2 = RzN__.SS.Main.Trades
                        tmptxt2 = Strings.Left(QFE1.buyZScoreMode.Items(RzN__.SP.bZScoreMode.Idx), 8) & RzN__.SP.bZScoreMode.Text1
                        QFE1.buyZScoreMode.Items(RzN__.SP.bZScoreMode.Idx) = tmptxt2
                        QFE1.BZScrVal_.Text = RzN__.SP.bZScoreValue.Text1
                        QFE1.txtZCN____.Text = Format(RzN__.SS.Main.Trades, "00000")
                        QFE1.On_.Text = Format(Counters.Onn, "0000000")
                        Counters.Onn = Counters.Onn + 1
                        Call display(RzN__, RZN__TS, QFE1.qSaveTrades0.Checked)
                        QFE1.txtZCN____.Text = Format(RzN__.SS.Main.Trades, "00000")
                    Else
                        Counters.Off = Counters.Off + 1
                        QFE1.Off_.Text = Format$(Counters.Off, "0000000")
                    End If
                Case 3
                    If QFE1.buyZScoreMode.GetItemChecked(RMain.SP.bZScoreMode.Idx) Then
                        Counters.totalIterations = Counters.totalIterations + 1
                        Call PreLoop(RzPP_)
                        RzPP_.SP = RMain.SP
                        RzPP_.SP.bZScoreMode.Text = Strings.Left(QFE1.buyZScoreMode.Items(RzPP_.SP.bZScoreMode.Idx), 8)
                        Call byZScoreValue_Loop(RzPP_, RZPP_TS, thisDay)
                        t3 = RzPP_.SS.Main.Trades
                        tmptxt3 = Strings.Left(QFE1.buyZScoreMode.Items(RzPP_.SP.bZScoreMode.Idx), 8) & RzPP_.SP.bZScoreMode.Text1
                        QFE1.buyZScoreMode.Items(RzPP_.SP.bZScoreMode.Idx) = tmptxt3
                        QFE1.BZScrVal_.Text = RzPP_.SP.bZScoreValue.Text1
                        QFE1.On_.Text = Format(Counters.Onn, "0000000")
                        QFE1.txtZDPP___.Text = Format(RzPP_.SS.Main.Trades, "00000")
                        Counters.Onn = Counters.Onn + 1
                        Call display(RzPP_, RZPP_TS, QFE1.qSaveTrades0.Checked)
                    Else
                        Counters.Off = Counters.Off + 1
                        QFE1.Off_.Text = Format$(Counters.Off, "0000000")
                    End If
                Case 4
                    If QFE1.buyZScoreMode.GetItemChecked(RMain.SP.bZScoreMode.Idx) Then
                        Counters.totalIterations = Counters.totalIterations + 1
                        Call PreLoop(RzNN_)
                        RzNN_.SP = RMain.SP
                        RzNN_.SP.bZScoreMode.Text = Strings.Left(QFE1.buyZScoreMode.Items(RzPP_.SP.bZScoreMode.Idx), 8)
                        Call byZScoreValue_Loop(RzNN_, RZNN_TS, thisDay)
                        t4 = RzNN_.SS.Main.Trades
                        tmptxt4 = Strings.Left(QFE1.buyZScoreMode.Items(RzNN_.SP.bZScoreMode.Idx), 8) & RzNN_.SP.bZScoreMode.Text1
                        QFE1.buyZScoreMode.Items(RzNN_.SP.bZScoreMode.Idx) = tmptxt4
                        QFE1.BZScrVal_.Text = RzNN_.SP.bZScoreValue.Text1
                        QFE1.On_.Text = Format(Counters.Onn, "0000000")
                        QFE1.txtZENN___.Text = Format(RzNN_.SS.Main.Trades, "00000")
                        Counters.Onn = Counters.Onn + 1
                        Call display(RzNN_, RZNN_TS, QFE1.qSaveTrades0.Checked)
                    Else
                        Counters.Off = Counters.Off + 1
                        QFE1.Off_.Text = Format$(Counters.Off, "0000000")
                    End If
                Case 5
                    If QFE1.buyZScoreMode.GetItemChecked(RMain.SP.bZScoreMode.Idx) Then
                        Counters.totalIterations = Counters.totalIterations + 1
                        Call PreLoop(RzPN_)
                        RzPN_.SP = RMain.SP
                        '                                Call Doozee(RzPN_)
                        '        If RzNN_.SS.Main.Trades > 0 Then Stop
                        QFE1.txtZFPN___.Text = Format(RzPN_.SS.Main.Trades, "00000")
                        Counters.Onn = Counters.Onn + 1
                        '            Call display(RzPN_, RZpn_ts(xTrade), QFE1.qSaveTrades0.Checked)
                    End If
                Case 6
                    If QFE1.buyZScoreMode.GetItemChecked(RMain.SP.bZScoreMode.Idx) Then
                        Counters.totalIterations = Counters.totalIterations + 1
                        Call PreLoop(RzNP_)
                        RzNP_.SP = RMain.SP
                        '                               Call Doozee(RzPPP)
                        QFE1.txtZHPPP__.Text = Format(RzNP_.SS.Main.Trades, "00000")
                        Counters.Onn = Counters.Onn + 1
                        '             Call display()
                    End If
                Case 7
                    If QFE1.buyZScoreMode.GetItemChecked(RMain.SP.bZScoreMode.Idx) Then
                        Counters.totalIterations = Counters.totalIterations + 1
                        Call PreLoop(RzPPP)
                        RzPPP.SP = RMain.SP
                        '                               Call Doozee(RzPPP)
                        QFE1.txtZHPPP__.Text = Format(RzPPP.SS.Main.Trades, "00000")
                        Counters.Onn = Counters.Onn + 1
                        '             Call display()
                    End If
                Case 8
                    If QFE1.buyZScoreMode.GetItemChecked(RMain.SP.bZScoreMode.Idx) Then
                        Counters.totalIterations = Counters.totalIterations + 1
                        Call PreLoop(RzNNN)
                        RzNNN.SP = RMain.SP
                        '                               Call Doozee(RzNNN)
                        QFE1.txtZINNN__.Text = Format(RzNNN.SS.Main.Trades, "00000")
                        Counters.Onn = Counters.Onn + 1
                        '              Call display()
                    End If
                    '          Case 9
                    '             Call PreLoop(RzPNN)
                    '            RzPNN.SP = RMain.SP
                    '           RzPNN.SP.bZScoreStd.Idx = 9
                    '          If Form1.buyZScoreMode.GetItemChecked(RzPNN.SP.bZScoreStd.Idx) Then
                    ' '                               Call Doozee(RzPNN)
                    ' End 'If
                    '   Case 10
                    '      Call PreLoop(RzPNP)
                    '     RzPNP.SP = RMain.SP
                    '    RzPNP.SP.bZScoreStd.Idx = 10
                    '   If Form1.buyZScoreMode.GetItemChecked(RzPNP.SP.bZScoreStd.Idx) Then
                    ''                               Call Doozee(RzPNP)
                    'End If
                    '  Case 11
                    '     Call PreLoop(RzNPP)
                    '    RzNPP.SP = RMain.SP
                    '   RzNPP.SP.bZScoreStd.Idx = 11
                    '  If Form1.buyZScoreMode.GetItemChecked(RzNPP.SP.bZScoreStd.Idx) Then
                    ''                               Call Doozee(RzNPP)
                    'End If
                    '       Case 12
                    '          Call PreLoop(RzNPN)
                    '         RzNPN.SP = RMain.SP
                    '        RzNPN.SP.bZScoreStd.Idx = 12
                    '       If Form1.buyZScoreMode.GetItemChecked(RzNPN.SP.bZScoreStd.Idx) Then
                    ''                               Call Doozee(RzNPN)
                    'End If
                    '       Case 13
                    '          Call PreLoop(RzPPN)
                    '         RzPPN.SP = RMain.SP
                    '        RzPPN.SP.bZScoreStd.Idx = 13
                    '       If Form1.buyZScoreMode.GetItemChecked(RzPPN.SP.bZScoreStd.Idx) Then
                    ''                          Call Doozee(RzPPN)
                    'End If
                    '       Case 14
                    '          Call PreLoop(RzNNP)
                    '         RzNNP.SP = RMain.SP
                    '        RzNPP.SP.bZScoreStd.Idx = 14
                    '       If Form1.buyZScoreMode.GetItemChecked(RzNPP.SP.bZScoreStd.Idx) Then
                    ''                            Call Doozee(RzNNP)
                    '       End If
                Case Else
                    Stop
            End Select

            QFE1.Iterations.Text = Format(Counters.totalIterations, "0000000")
            '            Form1.buyZScoreMode.Items(RMain.SP.bZScoreMode.Idx) =
            '           Strings.Left(Form1.buyZScoreMode.Items(RMain.SP.bZScoreMode.Idx), 5) &
            '          RMain.SP.bZScoreMode.Text1
            '         RMain.SP.bZScoreMode.Text1 = "td=" & Format$(RMain.SS.Main.Trades, "0000") &
            '                      "!h=" & Format$(RMain.SS.Main.zScore.Hits, "0000") &
            '                      "m=" & Format$(RMain.SS.Main.zScore.Misses, "0000") &
            '                      "zScr=" & StrSign(RMain.SS.Main.zScore.V)
            '     Form1.BuyZSCRModeTxt.Text = RMain.SP.bZScoreMode.Text1
            '    Form1.BuyZScoreModetxt.Text = RMain.SP.bZScoreMode.Text1
            '   Form1.bZScr_.Text = RMain.SP.bZScoreMode.Text1
            Application.DoEvents()
        Next RMain.SP.bZScoreMode.Idx
exitt:
        '        Else
        '        Counters.Off = Counters.Off + 1
        '        Form1.offLastDay.Text = Format$(Counters.Off, "000000")
        '        End If
    End Sub
    Public Sub byZScoreValue_Loop(ByRef R As Results, ByRef dstTS() As Trades, ByRef thisDay As Integer)
        For R.SP.bZScoreValue.Idx = 0 To 4
            Counters.ZIterNo = Counters.ZIterNo + 1
            If QFE1.buyZScoreValue.GetItemChecked(R.SP.bZScoreValue.Idx) Then
                ReDim dstTS(0)
                ReDim dstTS(5000)
                QFE1.BuyZScoreValueTxt.Text = QFE1.buyZScoreMode.Items(RMain.SP.bZScoreValue.Idx)
                R.SP.bZScoreValue.Text = Strings.Left(QFE1.buyZScoreValue.Items(R.SP.bZScoreValue.Idx), 13)
                R.SP.bZScoreValue.Text1 = QFE1.buyZScoreValue.Items(R.SP.bZScoreValue.Idx)
                Call DOZXXXX(R, dstTS, 1, RMain.SS.Main.Trades, RMainTS)
                If R.SS.Main.Trades >= 13 Then
                    lastDayTrade.bZscoreMode.Signal = R.SP.bZScoreMode.Text & "!" & R.SP.bZScoreValue.Text & "=" &
                          StrSign(dstTS(R.SS.Main.Trades).Profit0) & "!" &
                           StrSign(dstTS(R.SS.Main.Trades - 1).Profit0) & "!" &
                            StrSign(dstTS(R.SS.Main.Trades - 2).Profit0)
                    R.SS.Main.absQuantum = Math.Abs(R.SS.Main.Q.avg)
                    lastDayTrade.bEntry.qHit = SetLastDayTradenozzz(R, dstTS, thisDay)
                    QFE1.buyZScoreValue.Text = R.SP.bZScoreValue.Text & "=" & StrSign(dstTS(R.SS.Main.Trades).Profit0) & "!"
                    Call Calc_HitsMisses(R)
                    Call Write_Parameters0(R, dstTS, thisDay)
                    '                Form1.buyZScoreValue.Items(R.SP.bZScoreMode.Idx) = Strings.Left(Form1.buyZScoreValue.Items(R.SP.bZScoreValue.Idx), 15)
                End If
            End If
        Next RMain.SP.bZScoreValue.Idx
        Application.DoEvents()
    End Sub
    Private Function qHitValueZ(ByRef R As Results, ByRef trd As Trades) As Boolean
        Static q As Boolean
        With R
            Select Case R.SP.bZScoreValue.Idx
                Case 0
                    q = True
                Case 1
                    q = trd.bZScoreValue.V < -3.0
                Case 2
                    q = trd.bZScoreValue.V >= -3.0 And trd.bZScoreValue.V < -2.5
                Case 3
                    q = trd.bZScoreValue.V >= -2.5 And trd.bZScoreValue.V < -2.0
                Case 4
                    q = trd.bZScoreValue.V >= -2.0 And trd.bZScoreValue.V < -1.5
                Case 5
                    q = trd.bZScoreValue.V >= -1.5 And trd.bZScoreValue.V < -1.0
                Case 6
                    q = trd.bZScoreValue.V >= -1.0 And trd.bZScoreValue.V < -0.5
                Case 7
                    q = trd.bZScoreValue.V >= -0.5 And trd.bZScoreValue.V < 0.0
                Case 8
                    q = trd.bZScoreValue.V >= 0.0 And trd.bZScoreValue.V < +0.5
                Case 9
                    q = trd.bZScoreValue.V >= 0.5 And trd.bZScoreValue.V < +1.0
                Case 10 ' a
                    q = trd.bZScoreValue.V >= 1.0 And trd.bZScoreValue.V < +1.5
                Case 11 ' b
                    q = trd.bZScoreValue.V >= 1.5 And trd.bZScoreValue.V < +2.0
                Case 12 ' c
                    q = trd.bZScoreValue.V >= 2.0 And trd.bZScoreValue.V < +2.5
                Case 13 ' d
                    q = trd.bZScoreValue.V >= 2.5 And trd.bZScoreValue.V < +3.0
                Case 14 ' e
                    q = trd.bZScoreValue.V >= 3.0
                Case Else
                    q = False
                    Stop
            End Select
        End With
        qHitValueZ = q
    End Function
    Public Sub DOZXXXX(ByRef R As Results, ByRef dstTS() As Trades,
                       startTr As Integer, endTr As Integer, srcTS() As Trades)
        Dim zCnt As Integer, trd As Integer
        Dim totalprofit As Single
        With R.SS.Main
            .Days = RMain.SS.Main.Days
            .bSignal.Hits = RMain.SS.Main.bSignal.Hits
            .bTrigger.Hits = RMain.SS.Main.bTrigger.Hits
            .bTrigger.Misses = RMain.SS.Main.bTrigger.Misses
            .zScoreMode.Hits = 0
            .zScoreMode.Misses = 0
            .zScoreMode.tot = 0
            .zScoreValue.Hits = 0
            .zScoreValue.Misses = 0
            .zScoreValue.tot = 0
            'Call CalcZScoreEachTrade(R, dstTS)
            'If R.SS.Main.Trades < 35 Then Stop
            If R.SP.bZScoreMode.Idx = 0 Then
                startTr = 13
            Else
                startTr = 13
            End If
            zCnt = 0
            For trd = startTr + 1 To endTr
                .zScore.tot = .zScore.tot + 1
                newTrade = srcTS(trd)
                'If newTrade.Profit0 = 0.0 Then Stop
                newTrade.bZScoreValue.V = CalcZScore(R, dstTS, trd - 12, trd)
                If newTrade.bZScoreValue.V = 0.0 Then Stop
                newTrade.bZscoreMode.Text = R.SP.bZScoreMode.Text
                newTrade.bZscoreMode.qHit = qHitPatternZ(R, dstTS, trd, 1)
                If newTrade.bZscoreMode.qHit Then
                    .zScoreMode.Hits = .zScoreMode.Hits + 1
                    QFE1.qbZScrMode.Text = "qbZScrMode=T"
                Else
                    .zScoreMode.Misses = .zScoreMode.Misses + 1
                    QFE1.qbZScrMode.Text = "qbZScrMode=F"
                End If
                newTrade.bZScoreValue.qHit = qHitValueZ(R, dstTS(trd))
                If newTrade.bZScoreValue.qHit Then
                    .zScoreValue.Hits = .zScoreValue.Hits + 1
                    QFE1.qbZScrValue.Text = "qbZScrV=T"
                Else
                    .zScoreValue.Misses = .zScoreValue.Misses + 1
                    QFE1.qbZScrValue.Text = "qbZScrV=F"
                End If
                newTrade.bZScore.qHit = newTrade.bZscoreMode.qHit And newTrade.bZScoreValue.qHit
                If newTrade.bZScore.qHit Then
                    .zScore.Hits = .zScore.Hits + 1
                    zCnt = zCnt + 1
                    'If destR.SP.bZScore.Idx = 2 And R.TS(trd - 1).Profit0 > 0 Then Stop
                    'If trds.Profit0 > 0.01 Then Stop
                    If zCnt = 1 Then
                        totalprofit = newTrade.Profit0
                        newTrade.totProfit = totalprofit
                        newTrade.totProfit = totalprofit
                    Else
                        totalprofit = totalprofit + newTrade.Profit0
                        newTrade.totProfit = totalprofit
                        newTrade.avgProfit = newTrade.totProfit / zCnt
                        newTrade.oldTradeNo = newTrade.TradeNo
                        newTrade.TradeNo = zCnt
                    End If
                    dstTS(zCnt) = newTrade
                End If
            Next trd
            R.SS.Main.Trades = zCnt
        End With
    End Sub
    Public Function StrLD(ByRef R As Results, ByRef t() As Trades) As String
        Dim tstr1 As String
        tstr1 = R.sSymbol & "|" & R.SP.bMonth.Text & "|"
        If Counters.qBuyandHold.qHit Then
            tstr1 = tstr1 & "bh="
        Else
            tstr1 = tstr1 & "tr="
        End If
        If lastDayTrade.bEntry.qHit Then
            tstr1 = tstr1 & "|on" & Counters.last_Date
        Else
            tstr1 = tstr1 & "|of" & Counters.last_Date
        End If
        tstr1 = tstr1 & "|dy" & Format$(R.SS.Main.Days, "0000") & "|tr=" & Format(R.SS.Main.Trades, "0000")
        tstr1 = tstr1 & "|Q=" & StrSign((R.SS.Main.Q.avg))
        tstr1 = tstr1 & "|nd=" & R.SP.bInday.Text & ":" & Format$(R.SS.Main.bInDay.Hits, "000")
        tstr1 = tstr1 & Strings.Left(R.SP.bDOW.Text, 3) & ":" & Format(R.SS.Main.bDOW.Hits, "0000")
        tstr1 = tstr1 & "|" & Left(R.SP.bEntry.Text, 6) & "on" & lastTrade.bDate & "@" &
         Format$(lastTrade.bSignal.actualPrice, "000.000") & "dh:" &
          Format$(R.SS.Main.DH.avg, "0.00") & "|" & Format$(R.SP.sMaxDH.V, "00") & "|" &
           R.SP.sEntry.Text & Left(R.SP.sSignal.Text, 8) & "|" &
            R.SP.bZScoreMode.Text & ":TP=" & Format(R.SP.TP.V, "0.00") &
             "o=" & Format(op(Counters.this_Day - 1), "000.000") &
              "h:" & Format(hi(Counters.this_Day - 1), "000.000") &
               "l:" & Format(lo(Counters.this_Day - 1), "000.000") &
                "c:" & Format(cl(Counters.this_Day - 1), "000.000") & ">>" &
                  RMain.SP.TP.Text & "|tr" & "|tr" & Format(R.SS.Main.Trades, "0000") &
                   "w%=" & Format(R.SS.Main.wPcntg, "0.00") &
                    "|w=" & Format(R.SS.Main.W, "0000") &
                     "|l=" & Format(R.SS.Main.L, "0000")
        tstr1 = tstr1 & "|Z=" & R.SP.bZScoreMode.Text & ":" & Format(R.SS.Main.zScore.Hits, "0000") &
          ":" & Strings.Left(R.SP.bZScoreValue.Text, 8) & ":" &
           Format(R.SS.Main.zScore.Hits, "0000") & "||"
        'tstr12 = "" '"TD1:" & R.SP.bTD1.Text & ":" & Format$(RHitMisses.SS.Main.bTD1.Hits, "0000") & "|" &
        '"TD2:" & R.SP.bTD2.Text & ":" & Format$(RHitMisses.SS.Main.bTD2.Hits, "0000")
        'tstr13 = "||TRIGR:" & R.SP.bTrigger.Text & "@" & Format(R.SP.bTrigger.Price, "000.000") &
        'R.SP.bTrigger.Text & Format(R.SP.bTrigger.Price, "000.000") & "h=" &
        'Format$(RHitMisses.SS.Main.bTrigger.Hits, "0000") & "||" &
        'R.SP.sEntry.Text
        '     '                        "@" & Format$(lastTrade.sSignal.trigPrice, "000.000") & _
        '    '                       "-" & _
        '   '                        "on" & (R.SP.bTD1.Text) & ":" & _
        '  '                        "@" & Format$(lastTrade.bTD1.actualPrice, "000.000") & _
        StrLD = tstr1
    End Function
    Public Function mainTradesStr(ByRef tr() As Trades) As String
        With RMain
            Static t As Integer, str As String
            t = RMain.SS.Main.Trades
            Select Case t
                Case >= 13
                    str = "profitsMn^" & Format(t, "0000") & "tr:" &
         StrSign(tr(t).Profit0) & "^" & StrSign(tr(t - 1).Profit0) &
          "^" & StrSign(tr(t - 2).Profit0) & "^" & StrSign(tr(t - 3).Profit0) &
           "^" & StrSign(tr(t - 4).Profit0) & "^" & StrSign(tr(t - 5).Profit0) &
            "^" & StrSign(tr(t - 6).Profit0) & "^" & StrSign(tr(t - 7).Profit0) &
             "^" & StrSign(tr(t - 8).Profit0) & "^" & StrSign(tr(t - 9).Profit0) &
              "^" & StrSign(tr(t - 10).Profit0) & "^" & StrSign(tr(t - 11).Profit0) &
               "^" & StrSign(tr(t - 12).Profit0) & "^" & StrSign(tr(t - 13).Profit0)
                Case Else
                    str = "less than 13"
            End Select
            mainTradesStr = str
        End With
    End Function
    Public Function zScoreTradesStr(ByRef r As Results, ByRef tr() As Trades) As String
        With RMain
            Static t As Integer, str As String
            t = RMain.SS.Main.Trades
            Select Case t
                Case >= 13
                    str = "profitsMn^" & Format(t, "0000") & "tr:" &
         StrSign(tr(t).Profit0) & "^" & StrSign(tr(t - 1).Profit0) &
          "^" & StrSign(tr(t - 2).Profit0) & "^" & StrSign(tr(t - 3).Profit0) &
           "^" & StrSign(tr(t - 4).Profit0) & "^" & StrSign(tr(t - 5).Profit0) &
            "^" & StrSign(tr(t - 6).Profit0) & "^" & StrSign(tr(t - 7).Profit0) &
             "^" & StrSign(tr(t - 8).Profit0) & "^" & StrSign(tr(t - 9).Profit0) &
              "^" & StrSign(tr(t - 10).Profit0) & "^" & StrSign(tr(t - 11).Profit0) &
               "^" & StrSign(tr(t - 12).Profit0) & "^" & StrSign(tr(t - 13).Profit0)
                Case Else
                    str = "less than 13"
            End Select
            zScoreTradesStr = str
        End With
    End Function
    Public Function StrTr(ByRef R As Results, Tr As Trades) As String
        StrTr = Counters.currentSecurity &
        "!" & R.SP.bZScoreMode.Text & StrSign(Tr.Profit0) & "!" &
        "on" & Strings.Right(Tr.bDate, 4) & ":" & Tr.bDOW.Date_ &
        "||sg" & Tr.bSignal.Text & "!" &
        "@bex" & Format$(Tr.bEntry.executePrice, "000.000") & "-" & Strings.Right(Tr.bEntry.Date_, 4) & ":" &
        "@btg" & Format$(Tr.bTrigger.triggerPrice, "000.000") & "-" & Strings.Right(Tr.bTrigger.Date_, 4) & ":" &
        "@bsg" & Format$(Tr.bSignal.signalPrice, "000.000") & "-" & Strings.Right(Tr.bSignal.Date_, 4) &
        "||" & Tr.sSignal.Text & "!" &
        "@sex" & Format$(Tr.sEntry.executePrice, "000.000") & "-" & Strings.Right(Tr.sEntry.Date_, 4) & ":" &
        "dh:" & Format$(Tr.maxDH, "00") &
        Format$(Tr.sPrice, "000.000") & "-" & Format$(Tr.bPrice, "000.000") & "%" &
        StrSign(Tr.Profit0)
    End Function
    Public Function QTD_(ByRef td As Sig, ByRef xDay As Integer, ByRef db As Integer) As Boolean
        Static qTD As Boolean
        With td
            Select Case .Idx
                Case 0
                    .triggerPrice = cl(xDay)
                    .actualPrice = cl(xDay)
                    .Text = dtStr1(xDay)
                    qTD = True
                Case 1
                    td.triggerPrice = op(xDay - db)
                    td.actualPrice = op(xDay)
                    If td.actualPrice < td.triggerPrice Then
                        td.executePrice = op(xDay)
                        qTD = True
                    Else
                        .executePrice = -1.0
                        qTD = False
                    End If
                Case 2
                    .triggerPrice = hi(xDay - db)
                    .actualPrice = op(xDay)
                    If .actualPrice < .triggerPrice Then
                        .executePrice = op(xDay)
                        qTD = True
                    Else
                        .executePrice = -1.0
                        qTD = False
                    End If
                Case 3
                    .triggerPrice = lo(xDay - db)
                    .actualPrice = op(xDay)
                    If .actualPrice < .triggerPrice Then
                        .executePrice = op(xDay)
                        qTD = True
                    Else
                        .executePrice = -1.0
                        qTD = False
                    End If
                Case 4
                    .triggerPrice = cl(xDay - db)
                    .actualPrice = op(xDay)
                    If .actualPrice < .triggerPrice Then
                        .executePrice = op(xDay)
                        qTD = True
                    Else
                        .executePrice = -1.0
                        qTD = False
                    End If
                Case 5
                    .triggerPrice = op(xDay - db)
                    .actualPrice = hi(xDay)
                    newTrade.bTD1.Text = "O<" & Format(.triggerPrice, "000.000")
                    If .actualPrice < .triggerPrice Then
                        .executePrice = cl(xDay)
                        qTD = True
                    Else
                        .executePrice = -1.0
                        qTD = False
                    End If
                Case 6
                    .triggerPrice = hi(xDay - db)
                    .actualPrice = hi(xDay)
                    If .actualPrice < .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 7
                    .triggerPrice = lo(xDay - db)
                    .actualPrice = hi(xDay)
                    If .actualPrice < .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 8
                    .triggerPrice = cl(xDay - db)
                    .actualPrice = hi(xDay)
                    If .actualPrice < .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 9
                    .triggerPrice = op(xDay - db)
                    .actualPrice = lo(xDay)
                    If .actualPrice < .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 10
                    .triggerPrice = hi(xDay - db)
                    .actualPrice = lo(xDay)
                    If .actualPrice < .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 11
                    .triggerPrice = lo(xDay - db)
                    .actualPrice = lo(xDay)
                    If .actualPrice < .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 12
                    .triggerPrice = cl(xDay - db)
                    .actualPrice = lo(xDay)
                    If td.actualPrice < .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 13
                    .triggerPrice = op(xDay - db)
                    .actualPrice = cl(xDay)
                    If .actualPrice < .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 14
                    .triggerPrice = hi(xDay - db)
                    .actualPrice = cl(xDay)
                    If .actualPrice < .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 15
                    .triggerPrice = lo(xDay - db)
                    .actualPrice = cl(xDay)
                    If .actualPrice < .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 16
                    .triggerPrice = cl(xDay - db)
                    .actualPrice = cl(xDay)
                    If .actualPrice < .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 17
                    .triggerPrice = op(xDay - db)
                    .actualPrice = op(xDay)
                    If .actualPrice > .triggerPrice Then
                        qTD = True
                        .executePrice = op(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 18
                    .triggerPrice = hi(xDay - db)
                    .actualPrice = op(xDay)
                    If .actualPrice > .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 19
                    .triggerPrice = lo(xDay - db)
                    .actualPrice = op(xDay)
                    If .actualPrice > .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 20
                    .triggerPrice = cl(xDay - db)
                    .actualPrice = op(xDay)
                    If .actualPrice > .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 21
                    .triggerPrice = op(xDay - db)
                    .actualPrice = hi(xDay)
                    If .actualPrice > .triggerPrice Then
                        qTD = True
                        td.executePrice = op(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 22
                    .triggerPrice = hi(xDay - db)
                    .actualPrice = hi(xDay)
                    If .actualPrice > .triggerPrice Then
                        qTD = True
                        td.executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 23
                    .triggerPrice = lo(xDay - db)
                    .actualPrice = hi(xDay)
                    If .actualPrice > .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 24
                    .triggerPrice = cl(xDay - db)
                    .actualPrice = hi(xDay)
                    If .actualPrice > .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 25
                    .triggerPrice = op(xDay - db)
                    .actualPrice = lo(xDay)
                    If .actualPrice > .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 26
                    .triggerPrice = hi(xDay - db)
                    .actualPrice = lo(xDay)
                    If td.actualPrice > .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 27
                    .triggerPrice = lo(xDay - db)
                    .actualPrice = lo(xDay)
                    If .actualPrice > .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        td.executePrice = -1.0
                    End If
                Case 28
                    .triggerPrice = cl(xDay - db)
                    .actualPrice = lo(xDay)
                    If .actualPrice > .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 29
                    .triggerPrice = op(xDay - db)
                    .actualPrice = cl(xDay)
                    If .actualPrice > .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 30
                    .triggerPrice = hi(xDay - db)
                    .actualPrice = cl(xDay)
                    If .actualPrice > .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 31
                    .triggerPrice = lo(xDay - db)
                    .actualPrice = cl(xDay)
                    If .actualPrice > .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case 32
                    .triggerPrice = cl(xDay - db)
                    .actualPrice = cl(xDay)
                    If .actualPrice > .triggerPrice Then
                        qTD = True
                        .executePrice = cl(xDay)
                    Else
                        qTD = False
                        .executePrice = -1.0
                    End If
                Case Else
                    Stop
            End Select
        End With
        QTD_ = qTD
    End Function
    Public Function QTD_Text1(ByRef td As Sig, ByRef xDay As Integer) As String
        Static qTD As Boolean, qTD_Text As String, str1 As String, str2 As String
        qTD_Text = " "
        Select Case td.Idx
            Case 0
                qTD = True
                qTD_Text = "none"
            Case 1
                str1 = Format(op(xDay - td.db), "000.000")
                str2 = Format(op(xDay), "000.000")
                qTD_Text = str2 & "<" & str1
            Case 2
                str1 = Format(hi(xDay - td.db), "000.000")
                str2 = Format(op(xDay), "000.000")
                qTD_Text = str2 & "<" & str1
            Case 3
                str1 = Format(lo(xDay - td.db), "000.000")
                str2 = Format(op(xDay), "000.000")
                qTD_Text = str2 & "<" & str1
            Case 4
                str1 = Format(cl(xDay - td.db), "000.000")
                str2 = Format(op(xDay), "000.000")
                qTD_Text = str2 & "<" & str1
            Case 5
                str1 = Format(op(xDay - td.db), "000.000")
                str2 = Format(hi(xDay), "000.000")
                qTD_Text = str2 & "<" & str1
            Case 6
                str1 = Format(hi(xDay - td.db), "000.000")
                str2 = Format(hi(xDay), "000.000")
                qTD_Text = str2 & "<" & str1
            Case 7
                str1 = Format(lo(xDay - td.db), "000.000")
                str2 = Format(hi(xDay), "000.000")
                qTD_Text = str2 & "<" & str1
            Case 8
                str1 = Format(cl(xDay - td.db), "000.000")
                str2 = Format(hi(xDay), "000.000")
                qTD_Text = str2 & "<" & str1
            Case 9
                str1 = Format(op(xDay - td.db), "000.000")
                str2 = Format(lo(xDay), "000.000")
                qTD_Text = str2 & "<" & str1
            Case 10
                str1 = Format(hi(xDay - td.db), "000.000")
                str2 = Format(lo(xDay), "000.000")
                qTD_Text = str2 & "<" & str1
            Case 11
                str1 = Format(lo(xDay - td.db), "000.000")
                str2 = Format(lo(xDay), "000.000")
                qTD_Text = str2 & "<" & str1
            Case 12
                str1 = Format(cl(xDay - td.db), "000.000")
                str2 = Format(lo(xDay), "000.000")
                qTD_Text = str2 & "<" & str1
            Case 13
                str1 = Format(op(xDay - td.db), "000.000")
                str2 = Format(cl(xDay), "000.000")
                qTD_Text = str2 & "<" & str1
            Case 14
                str1 = Format(hi(xDay - td.db), "000.000")
                str2 = Format(cl(xDay), "000.000")
                qTD_Text = str2 & "<" & str1
            Case 15
                str1 = Format(lo(xDay - td.db), "000.000")
                str2 = Format(cl(xDay), "000.000")
                qTD_Text = str2 & "<" & str1
            Case 16
                str1 = Format(cl(xDay - td.db), "000.000")
                str2 = Format(cl(xDay), "000.000")
                qTD_Text = str2 & "<" & str1
            Case 17
                str1 = Format(op(xDay - td.db), "000.000")
                str2 = Format(op(xDay), "000.000")
                qTD_Text = str2 & ">" & str1
            Case 18
                str1 = Format(hi(xDay - td.db), "000.000")
                str2 = Format(op(xDay), "000.000")
                qTD_Text = str2 & ">" & str1
            Case 19
                str1 = Format(lo(xDay - td.db), "000.000")
                str2 = Format(op(xDay), "000.000")
                qTD_Text = str2 & ">" & str1
            Case 20
                str1 = Format(cl(xDay - td.db), "000.000")
                str2 = Format(op(xDay), "000.000")
                qTD_Text = str2 & ">" & str1
            Case 21
                str1 = Format(op(xDay - td.db), "000.000")
                str2 = Format(hi(xDay), "000.000")
                qTD_Text = str2 & ">" & str1
            Case 22
                str1 = Format(hi(xDay - td.db), "000.000")
                str2 = Format(hi(xDay), "000.000")
                qTD_Text = str2 & ">" & str1
            Case 23
                str1 = Format(lo(xDay - td.db), "000.000")
                str2 = Format(hi(xDay), "000.000")
                qTD_Text = str2 & ">" & str1
            Case 24
                str1 = Format(cl(xDay - td.db), "000.000")
                str2 = Format(hi(xDay), "000.000")
                qTD_Text = str2 & ">" & str1
            Case 25
                str1 = Format(op(xDay - td.db), "000.000")
                str2 = Format(lo(xDay), "000.000")
                qTD_Text = str2 & ">" & str1
            Case 26
                str1 = Format(hi(xDay - td.db), "000.000")
                str2 = Format(lo(xDay), "000.000")
                qTD_Text = str2 & ">" & str1
            Case 27
                str1 = Format(lo(xDay - td.db), "000.000")
                str2 = Format(lo(xDay), "000.000")
                qTD_Text = str2 & ">" & str1
            Case 28
                str1 = Format(cl(xDay - td.db), "000.000")
                str2 = Format(lo(xDay), "000.000")
                qTD_Text = str2 & ">" & str1
            Case 29
                str1 = Format(op(xDay - td.db), "000.000")
                str2 = Format(cl(xDay), "000.000")
                qTD_Text = str2 & ">" & str1
            Case 30
                str1 = Format(hi(xDay - td.db), "000.000")
                str2 = Format(cl(xDay), "000.000")
                qTD_Text = str2 & ">" & str1
            Case 31
                str1 = Format(lo(xDay - td.db), "000.000")
                str2 = Format(cl(xDay), "000.000")
                qTD_Text = str2 & ">" & str1
            Case 32
                str1 = Format(cl(xDay - td.db), "000.000")
                str2 = Format(cl(xDay), "000.000")
                qTD_Text = str2 & ">" & str1
            Case Else
                Stop
        End Select
        QTD_Text1 = qTD_Text
    End Function
    Public Function QBuySignal(ByRef thisDay As Integer) As Boolean
        Static q As Boolean
        'Stop
        '     newTrade.BExeTg.qHit = False
        newTrade.BuyOnOpen.qHit = False
        '      newTrade.BExeLo.qHit = False
        newTrade.BuyOnClose.qHit = False
        '       newTrade.BExeNOp.qHit = False
        newTrade.BuyOnNextClose.qHit = False
        '        newTrade.BExeMiss.qHit = False
        '        newTrade.qPastEndDay = newTrade.bDayNo > Counters.end_day
        q = False
        Select Case RMain.SP.bSignal.Idx
            Case 0
                '           db = 0
                newTrade.bSignal.Day = thisDay
                newTrade.bSignal.actualPrice = op(newTrade.bSignal.Day)
                newTrade.bTrigger.Day = thisDay
                newTrade.bTrigger.triggerPrice = 0.001
                newTrade.bDayNo = thisDay
                '                newTrade.BExeOp.qHit = True
                q = True
            Case 1
                '            db = 2
                newTrade.bActual.Day = thisDay
                newTrade.bDayNo = newTrade.bActual.Day
                newTrade.bSignal.actualPrice = op(newTrade.bActual.Day)
                newTrade.bTrigger.Day = newTrade.bSignal.Day - 2
                newTrade.bTrigger.signalPrice = op(newTrade.bTrigger.Day)
                newTrade.bSignal.Day = thisDay
                newTrade.bSignal.executePrice = 0.001
                newTrade.BuyOnOpen.qHit = newTrade.bSignal.actualPrice <= newTrade.bTrigger.triggerPrice
                If newTrade.BuyOnOpen.qHit Then
                    q = True
                    '                    newTrade.BExeMiss.qHit = False
                    newTrade.bSignal.executePrice = newTrade.bSignal.actualPrice
                    If newTrade.bSignal.executePrice = 0.0 Then Stop
                Else
                    newTrade.BuyOnLow.qHit = (lo(newTrade.bSignal.actualDay) <= newTrade.bSignal.triggerPrice)
                    If newTrade.BuyOnLow.qHit Then
                        q = True
                        '                        newTrade.BExeMiss.qHit = False
                        newTrade.bSignal.executePrice = newTrade.bSignal.triggerPrice
                        If newTrade.bSignal.executePrice = 0.0 Then Stop
                    Else
                        q = False
                        newTrade.Bexemiss.qHit = True
                        newTrade.bSignal.executePrice = 0.0
                    End If
                End If
            Case 2
                '             db = 2
                newTrade.bSignal.actualDay = thisDay
                newTrade.bDayNo = newTrade.bSignal.actualDay
                newTrade.bSignal.executeDay = thisDay
                newTrade.bSignal.actualPrice = op(newTrade.bSignal.actualDay)
                newTrade.bSignal.triggerDay = newTrade.bSignal.actualDay - 2
                newTrade.bSignal.triggerPrice = cl(newTrade.bSignal.triggerDay)
                newTrade.bSignal.executePrice = 0.002
                newTrade.BuyOnOpen.qHit = newTrade.bSignal.actualPrice <= newTrade.bSignal.triggerPrice
                If newTrade.BuyOnOpen.qHit Then
                    q = True
                    '                    newTrade.BExeMiss.qHit = False
                    newTrade.bSignal.executePrice = newTrade.bSignal.actualPrice
                    If newTrade.bSignal.executePrice = 0.0 Then Stop
                Else
                    newTrade.BuyOnLow.qHit = (lo(newTrade.bSignal.actualDay) <= newTrade.bSignal.triggerPrice)
                    If newTrade.BuyOnLow.qHit Then
                        q = True
                        newTrade.Bexemiss.qHit = False
                        newTrade.bSignal.executePrice = newTrade.bSignal.triggerPrice
                        If newTrade.bSignal.executePrice = 0.0 Then Stop
                    Else
                        q = False
                        newTrade.Bexemiss.qHit = True
                        newTrade.bSignal.executePrice = 0.0
                    End If
                End If
            Case 3
                '              db = 1
                newTrade.bSignal.actualDay = thisDay
                newTrade.bDayNo = newTrade.bSignal.actualDay
                newTrade.bSignal.executeDay = thisDay
                newTrade.bSignal.actualPrice = op(newTrade.bSignal.actualDay)
                newTrade.bSignal.triggerDay = newTrade.bSignal.actualDay - 1
                newTrade.bSignal.triggerPrice = op(newTrade.bSignal.triggerDay)
                newTrade.bSignal.executePrice = 0.003
                newTrade.BuyOnOpen.qHit = newTrade.bSignal.actualPrice <= newTrade.bSignal.triggerPrice
                If newTrade.BuyOnOpen.qHit Then
                    q = True
                    newTrade.Bexemiss.qHit = False
                    newTrade.bSignal.executePrice = newTrade.bSignal.actualPrice
                    If newTrade.bSignal.executePrice = 0.0 Then Stop
                Else
                    newTrade.BuyOnLow.qHit = (lo(newTrade.bSignal.actualDay) <= newTrade.bSignal.triggerPrice)
                    If newTrade.BuyOnLow.qHit Then
                        q = True
                        newTrade.Bexemiss.qHit = False
                        newTrade.bSignal.executePrice = newTrade.bSignal.triggerPrice
                        If newTrade.bSignal.executePrice = 0.0 Then Stop
                    Else
                        q = False
                        newTrade.Bexemiss.qHit = True
                        newTrade.bSignal.executePrice = 0.0
                    End If
                End If
            Case 4
                '               db = 1
                newTrade.bSignal.actualDay = thisDay
                newTrade.bDayNo = newTrade.bSignal.actualDay
                newTrade.bSignal.executeDay = thisDay
                newTrade.bSignal.actualPrice = op(newTrade.bSignal.actualDay)
                newTrade.bSignal.triggerDay = newTrade.bSignal.actualDay - 1
                newTrade.bSignal.triggerPrice = cl(newTrade.bSignal.triggerDay)
                newTrade.bSignal.executePrice = 0.004
                newTrade.BuyOnOpen.qHit = newTrade.bSignal.actualPrice <= newTrade.bSignal.triggerPrice
                q = False
                If newTrade.BuyOnOpen.qHit Then
                    q = True
                    '                    newTrade.BExeMiss.qHit = False
                    newTrade.bSignal.executePrice = newTrade.bSignal.actualPrice
                    If newTrade.bSignal.executePrice = 0.0 Then Stop
                Else
                    newTrade.BuyOnLow.qHit = (lo(newTrade.bSignal.actualDay) <= newTrade.bSignal.triggerPrice)
                    If newTrade.BuyOnLow.qHit Then
                        q = True
                        '                        newTrade.BExeMiss.qHit = False
                        newTrade.bSignal.executePrice = newTrade.bSignal.triggerPrice
                        If newTrade.bSignal.executePrice = 0.0 Then Stop
                    Else
                        q = False
                        '                       newTrade.BExeMiss.qHit = True
                        newTrade.bSignal.executePrice = 0.0
                    End If
                End If
            Case 5
                '                db = 0
                newTrade.bSignal.actualDay = thisDay
                newTrade.bSignal.executeDay = thisDay
                newTrade.bDayNo = thisDay
                newTrade.bSignal.actualPrice = op(newTrade.bSignal.actualDay)
                newTrade.bSignal.triggerDay = thisDay
                newTrade.bSignal.triggerPrice = newTrade.bSignal.actualPrice
                newTrade.bSignal.executePrice = newTrade.bSignal.actualPrice
                newTrade.BuyOnOpen.qHit = True
                q = True
                If newTrade.bSignal.executePrice = 0.0 Then Stop
            Case 6
                '                db = 0
                newTrade.bSignal.actualDay = thisDay
                newTrade.bSignal.executeDay = thisDay
                newTrade.bDayNo = thisDay
                newTrade.bSignal.actualPrice = cl(newTrade.bSignal.actualDay)
                newTrade.bSignal.triggerDay = thisDay
                newTrade.bSignal.triggerPrice = newTrade.bSignal.actualPrice
                newTrade.bSignal.executePrice = newTrade.bSignal.actualPrice
                newTrade.BuyOnClose.qHit = True
                q = True
                If newTrade.bSignal.executePrice = 0.0 Then Stop
            Case 7
                '                    Stop
                q = True
                newTrade.bSignal.actualDay = thisDay + 1
                newTrade.bDayNo = newTrade.bSignal.actualDay
                '                newTrade.bActualPrice = op(newTrade.bSignal.actualDay)
                '              newTrade.bSignal.actualPrice = newTrade.bActualPrice
                newTrade.bTrigger.Day = newTrade.bDayNo
                newTrade.bSignal.executePrice = newTrade.bSignal.actualPrice
                newTrade.bSignal.triggerPrice = op(newTrade.bTrigger.Day)
                '               newTrade.bPrice = newTrade.bActualPrice
                newTrade.BuyOnOpen.qHit = True
            Case 8
                '                   Stop
                q = True
                newTrade.bSignal.actualDay = thisDay + 1
                newTrade.bDayNo = newTrade.bSignal.actualDay
                '             newTrade.bActualPrice = cl(newTrade.bSignal.actualDay)
                '            newTrade.bSignal.actualPrice = newTrade.bActualPrice
                newTrade.bTrigger.Day = newTrade.bDayNo
                newTrade.bSignal.executePrice = newTrade.bSignal.actualPrice
                newTrade.bSignal.triggerPrice = cl(newTrade.bTrigger.Day)
                '           newTrade.bPrice = newTrade.bActualPrice
                newTrade.BuyOnClose.qHit = True
            Case 9, 10
                Stop
                '                q = True
                '               If thisDay + 2 > Counters.end_day Then
                '              SP.bSignal.Text1 = SP.bSignal.Text1 & "@@" & dtStr1(Counters.end_day) & "+2Days"
                '             Else
                '            SP.bSignal.Text1 = R.SP.bSignal.Text1 & "@@" & dtStr1(R.thisDay + 2) & "+2Days"
     '           End If
            Case 11, 12
                Stop
                '                q = True
                '               If R.thisDay + 3 > Counters.end_day Then
                '              R.SP.bSignal.Text1 = R.SP.bSignal.Text1 & "@@" & dtStr1(Counters.end_day) & "+3Days"
                '             Else
                '            R.SP.bSignal.Text1 = R.SP.bSignal.Text1 & "@@" & dtStr1(R.thisDay + 3) & "+3Days"
                '           End If
            Case Else
                Stop
        End Select
        QBuySignal = q
    End Function
    Public Function QBuyTrigger(ByRef nt As Trades, ByRef thisDay As Integer) As Boolean
        Static q As Boolean, db As Integer
        With nt
            .BuyOnClose.qHit = False
            .BuyOnTrigger.qHit = False
            .BuyOnLow.qHit = False
            .BuyOnClose.qHit = False
            .Bexemiss.qHit = False
            .qPastEndDay = .bDayNo > Counters.end_day
            .bSignal.actualDay = thisDay
            '       If nt.bSignal.actualDay > Counters.last_Day Then Stop
            nt.bDayNo = nt.bSignal.actualDay
            .bSignal.executeDay = .bSignal.actualDay
            .BuyOnLow.qHit = False
            Select Case RMain.SP.bTrigger.Idx
                Case 0
                    db = 0
                    q = True
                    GoTo exitt
                Case 1 To 10
                    db = 1
                Case 11 To 20
                    db = 2
                Case Else
                    Stop
            End Select
            .bSignal.triggerDay = .bSignal.actualDay - db
            .bSignal.Date_ = dtStr1(.bSignal.triggerDay)
            .bSignal.triggerPrice = ExeTriggerPrice(RMain.SP.bTrigger.Idx, .bSignal.triggerDay)
            q = False
            .bSignal.actualPrice = op(.bSignal.actualDay)
            .bSignal.executePrice = 0.001
            .BuyOnOpen.qHit = (.bSignal.actualPrice <= .bSignal.triggerPrice)
            If .BuyOnOpen.qHit Then
                q = True
                '                .BExeMiss.qHit = False
                .bSignal.executePrice = .bSignal.actualPrice
            Else
                .BuyOnLow.qHit = (lo(.bSignal.actualDay) <= .bSignal.triggerPrice)
                If .BuyOnLow.qHit Then
                    q = True
                    '                    .BExeMiss.qHit = False
                    .bSignal.executePrice = .bSignal.triggerPrice
                Else
                    q = False
                    '                    .BExeMiss.qHit = True
                End If
            End If
            If .bSignal.executePrice = 0.0 Then Stop
exita:
            nt.bTrigger.triggerPrice = .bSignal.triggerPrice
exitt:
            .bTrigger.Signal = RMain.SP.bTrigger.Text & ":" & Format$(.bTrigger.triggerPrice, "000.000")
            QBuyTrigger = q
            '    If q Then Stop
        End With
    End Function
    Public Function ExeTriggerPrice(ByRef idx As Integer, ByRef d As Integer) As Single
        Static tp As Single
        Select Case idx
            Case 0
                tp = 0.0
            Case 1, 11
                tp = op(d)
            Case 2, 12
                tp = (op(d) + hi(d)) / 2
            Case 3, 13
                tp = (op(d) + lo(d)) / 2
            Case 4, 14
                tp = (op(d) + cl(d)) / 2
            Case 5, 15
                tp = hi(d)
            Case 6, 16
                tp = (hi(d) + lo(d)) / 2
            Case 7, 17
                tp = (hi(d) + cl(d)) / 2
            Case 8, 18
                tp = lo(d)
            Case 9, 19
                tp = (lo(d) + cl(d)) / 2
            Case 10, 20
                tp = cl(d)
            Case Else
                Stop
        End Select
        ExeTriggerPrice = tp
    End Function
    Public Function BsDayLoopCase00() As Integer
        Dim xx As String, str1 As String, str2 As String, xTrade As Integer
        Static addDBforSell As Integer, addDBforBuy As Integer
        Select Case RMain.SP.bSignal.Idx
            Case Is <= 4
                If RMain.SP.sSignal.Idx = 3 Or RMain.SP.sSignal.Idx = 4 Then
                    addDBforBuy = 1
                Else
                    addDBforBuy = 0
                End If
            Case Is = 5
                addDBforBuy = 1
            Case Is = 6
                addDBforBuy = 1
            Case Else
                Stop
        End Select
        Select Case RMain.SP.sSignal.Idx
            Case 0
                addDBforSell = addDBforBuy + RMain.SP.sMaxDH.V
            Case 1
                addDBforSell = addDBforBuy + addDBforSell + 0
            Case 2
                addDBforSell = addDBforBuy + addDBforSell + 0
            Case 3
                addDBforSell = addDBforBuy + RMain.SP.sMaxDH.V + 1
            Case 4
                addDBforSell = addDBforBuy + RMain.SP.sMaxDH.V + 1
            Case 5
                addDBforSell = addDBforBuy + RMain.SP.sMaxDH.V + 2
            Case 6
                addDBforSell = addDBforBuy + RMain.SP.sMaxDH.V + 2
            Case 7
                addDBforSell = addDBforBuy + RMain.SP.sMaxDH.V + 3
            Case 8
                addDBforSell = addDBforBuy + RMain.SP.sMaxDH.V + 3
            Case Else
                Stop
        End Select
        Counters.finalDay = Counters.end_day - addDBforSell
        Counters.finalDate = dtStr2(Counters.finalDay)
        QFE1.TxtFinalDay.Text = Counters.finalDate
        QFE1.TxtStartDate.Text = Counters.start_Date
        QFE1.TxtEndDate.Text = Counters.end_Date
        With RMain
            .SS.Main.bInDay.tot = 0
            .SS.Main.bInDay.Hits = 0
            .SS.Main.bInDay.Misses = 0
            .SS.Main.bDOW.tot = 0
            .SS.Main.bDOW.Hits = 0
            .SS.Main.bDOW.Misses = 0
            .SS.Main.bInDay.tot = 0
            .SS.Main.bInDay.Hits = 0
            .SS.Main.bInDay.Misses = 0
            .SS.Main.bTD1.tot = 0
            .SS.Main.bTD1.Hits = 0
            .SS.Main.bTD1.Misses = 0
            .SS.Main.bTD2.tot = 0
            .SS.Main.bTD2.Hits = 0
            .SS.Main.bTD2.Misses = 0
            .SS.Main.bTrigger.tot = 0
            .SS.Main.bTrigger.Hits = 0
            .SS.Main.bTrigger.Misses = 0
            .SS.Main.sTP.Hits = 0
            .SS.Main.Days = 0
            xx = Strings.InStr(RMain.sSymbol, " ")
            If xx <> 0 Then
                RMain.sSymbol = Strings.Left(RMain.sSymbol, xx) & "_____"
            Else
                RMain.sSymbol = RMain.sSymbol & "_____"
            End If
            RMain.sSymbol = Strings.Left(RMain.sSymbol, 5)
            str1 = RMain.sSymbol & ":" & Format$(Counters.SystemNumber, "00000")
            RMain.sSymbol = str1 & ":" & RMain.SP.bDOW.Text & "!" &
             RMain.SP.bInday.Text & "!" & RMain.SP.bZScoreMode.Text &
              "!" &
               RMain.SP.bEntry.Text & ":" &
                Format(RMain.SP.bMaxDH.V, "00") & ":" & RMain.SP.sEntry.Text
            str2 = RMain.sSymbol & ":" & Format$(Counters.SystemNumber, "0000") & ":" &
                RMain.SP.bDOW.Text & RMain.SP.bTD1.Text & ":" & RMain.SP.bTD2.Text & ":" &
                 RMain.SP.sMaxDH.Text & "=" & Format$(RMain.SS.Main.Q.avg, "0.00")
            '        R.qPosLTr = TSa.(Rmain.SS.Main.Trades). > 0.0
            RMain.SS.Main.Profits.tot = 0.0
            RMain.SS.Main.DH.tot = 0
            Call makeNewTrades()
            If RMain.SS.Main.Trades > 13 Then
                RMain.Symbol = Counters.currentSecurity
                RMain.sSymbol = RMain.Symbol & "|" & StrSign(RMain.SS.Main.Q.avg)
                RMain.lSymbol = RMain.sSymbol & "|" & RMain.SP.bZScoreMode.Text & "|" &
                    RMain.SP.bEntry.Text & " : " &
                    Format(RMain.SP.sMaxDH.V, "00") & ":" & RMain.SP.sEntry.Text
                RMain.tradesString = mainTradesStr(RMainTS)
                If QFE1.qSaveTrades0.Checked Then
                    For xTrade = 1 To RMain.SS.Main.Trades
                        Call insertTrade(RMain, RMainTS(xTrade))
                    Next xTrade
                End If
                RMain.SS.Main.bInDay.Pcntg = RMain.SS.Main.bInDay.Hits / (RMain.SS.Main.bInDay.Hits + RMain.SS.Main.bInDay.Misses)
                QFE1.Quantum_.Text = Format$(RMain.SS.Main.Q.tot, "0.000")
                QFE1.Trades_.Text = Format(RMain.SS.Main.Trades, "00000")
            End If
        End With
        BsDayLoopCase00 = RMain.SS.Main.Trades
        'QFE_Base.txtZATrades.Text = Format(RMain.SS.Main.Trades, "000000")5
    End Function
    Public Sub makeNewTrades()
        For Counters.this_Day = Counters.start_Day To Counters.end_day
            '            If Counters.this_Day = Counters.end_day Then Stop
            'If Counters.this_Day = Counters.end_day Then Stop
            'If R.thisDay > Counters.end_day - R.SP.sMaxDH.V Then Stop
            newTrade = New Trades
            RMain.SS.Main.Days = RMain.SS.Main.Days + 1
            Counters.this_Date = dtStr2(Counters.this_Day)
            newTrade.maxDH = RMain.SP.sMaxDH.V
            newTrade.Profit0 = 0
            newTrade.tradeDate = dtStr2(Counters.this_Day)
            newTrade.tradeDayNo = Counters.this_Day
            newTrade.bDate = dtStr2(Counters.this_Day)
            newTrade.bZscoreMode.Text = RMain.SP.bZScoreMode.Text
            newTrade.bMonth.Text = RMain.SP.bMonth.Text
            newTrade.bMonth.qHit = QExeMonth(Counters.this_Day)
            If newTrade.bMonth.qHit Then
                newTrade.bInDay.Text = RMain.SP.bInday.Text
                newTrade.bInDay.qHit = QbuyInDay(Counters.this_Day)
                If newTrade.bInDay.qHit Then
                    RMain.SS.Main.bSignal.Hits = RMain.SS.Main.bSignal.Hits + 1
                    newTrade.bDOW.qHit = QExeBDOW(Counters.this_Day)
                    '    If Counters.this_Day = Counters.end_day Then Stop
                    If newTrade.bDOW.qHit Then
                        newTrade.bDOW.Text = Strings.Left(DOWStr(Counters.this_Day), 1)
                        Call CountTD(Counters.this_Day)
                        newTrade.bTD0.qHit = newTrade.bTD1.qHit And newTrade.bTD2.qHit
                        If newTrade.bTD0.qHit Then
                            newTrade.bTrigger.qHit = QBuyTrigger(newTrade, Counters.this_Day)
                            '          Form1.TxtBuyTrigger.Text = R.SP.bTrigger.Text & "@" &
                            '         Format(lo(.bSignal.actualDay), "000.000") & "<<" &
                            'Format(newTrade.bTrigger.triggerPrice, " 000.000")
                            If newTrade.bTrigger.qHit Then
                                RMain.SS.Main.bTrigger.Hits = RMain.SS.Main.bTrigger.Hits + 1
                                '               Form1.TxtBuyTrigger.BackColor = SystemColors.ControlDark
                                '             Call T.bentry.rade()
                                newTrade.bEntry.qHit = True
                                If newTrade.bEntry.qHit Then
                                    newTrade.bEntry.qHit = QBExeEntry(Counters.this_Day)
                                    If newTrade.bPrice = 0 Then newTrade.bEntry.qHit = False
                                    If newTrade.bEntry.qHit Then
                                        qbuyonlastday(Counters.this_Day) = True
                                        RMain.SS.Main.bSignal.tot = RMain.SS.Main.bSignal.tot + 1
                                        Select Case RMain.SP.TP.Idx
                                            Case 0
                                                Call SellIt(RMain, RMainTS)
                                            Case 1 To 11
                                                Call SellItWithTargetProfit(RMain, RMainTS)
                                            Case Else
                                                Stop
                                        End Select
                                        'If Not newTrade.qPastEndDay Then
                                        newTrade.bEntry.qHit = True
                                        If newTrade.DH = 0 Then newTrade.DH = 1
                                        'If Math.Abs(newTrade.Profit0) <0.1 Then Stop
                                        RMain.SS.Main.Profits.tot = RMain.SS.Main.Profits.tot + newTrade.Profit0
                                        RMain.SS.Main.DH.tot = RMain.SS.Main.DH.tot + newTrade.DH
                                        newTrade.totProfit = RMain.SS.Main.Profits.tot
                                        newTrade.totDH = RMain.SS.Main.DH.tot
                                        newTrade.Quantum = newTrade.Profit0 / newTrade.DH / 10.0
                                        newTrade.totQuantum = newTrade.totProfit / newTrade.totDH / 10.0
                                        newTrade.bOpen = op(newTrade.bDayNo)
                                        newTrade.bHigh = hi(newTrade.bDayNo)
                                        newTrade.bLow = lo(newTrade.bDayNo)
                                        newTrade.bClose = cl(newTrade.bDayNo)
                                        If RMain.SS.Main.Trades > 1 Then
                                            lastTrade.Profit0 = RMainTS(RMain.SS.Main.Trades - 1).Profit0
                                        End If
                                        newTrade.sOpen = op(newTrade.sDayNo)
                                        newTrade.sHigh = hi(newTrade.sDayNo)
                                        newTrade.sLow = lo(newTrade.sDayNo)
                                        newTrade.sClose = cl(newTrade.sDayNo)
                                        RMainTS(RMain.SS.Main.Trades) = newTrade
                                        If newTrade.sTP.qHit Then
                                            newTrade.sTP.qHiti = 1
                                        Else
                                            newTrade.sTP.qHiti = 0
                                        End If
                                        '                            RMain.TS1(RMain.SS.Main.Trades) = newTrade1
                                        'End If
                                    End If
                                Else
                                    newTrade.bEntry.qHit = False
                                    RMain.SS.Main.bEntryStats.Misses = RMain.SS.Main.bEntryStats.Misses + 1
                                    '            RMain.ss.Main.qbuyonlastday(Counters.this_Day) = False
                                End If
                            Else
                                RMain.SS.Main.bEntryStats.Misses = RMain.SS.Main.bEntryStats.Misses + 1
                            End If
                        Else
                            RMain.SS.Main.bTrigger.Misses = RMain.SS.Main.bTrigger.Misses + 1
                            '              Form1.TxtBuyTrigger.BackColor = SystemColors.ControlLight
                        End If
                    End If
                End If
            End If
        Next Counters.this_Day
    End Sub
    Public Function QExeBDOW(ByRef thisDay As Integer) As Boolean
        Static q As Boolean
        With RMain.SP.bDOW
            Select Case .Idx
                Case 0
                    q = True
                Case 1
                    q = (dowNo(thisDay) = 1)
                Case 2
                    q = (dowNo(thisDay) <> 1)
                Case 3
                    q = (dowNo(thisDay) = 2)
                Case 4
                    q = (dowNo(thisDay) <> 2)
                Case 5
                    q = (dowNo(thisDay) = 3)
                Case 6
                    q = (dowNo(thisDay) <> 3)
                Case 7
                    q = (dowNo(thisDay) = 4)
                Case 8
                    q = (dowNo(thisDay) <> 4)
                Case 9
                    q = (dowNo(thisDay) = 5)
                Case 10
                    q = (dowNo(thisDay) <> 5)
                Case Else
                    q = False
                    Stop
            End Select
        End With
        QExeBDOW = q
    End Function
    Public Function QExeMonth(ByRef thisDay As Integer) As Boolean
        Static q As Boolean
        With RMain.SP.bMonth
            Select Case .Idx
                Case 0
                    q = True
                Case 1
                    q = (monthNo(thisDay) = 0)
                Case 2
                    q = (monthNo(thisDay) = 1)
                Case 3
                    q = (monthNo(thisDay) = 2)
                Case 4
                    q = (monthNo(thisDay) = 3)
                Case 5
                    q = (monthNo(thisDay) = 4)
                Case 6
                    q = (monthNo(thisDay) = 5)
                Case 7
                    q = (monthNo(thisDay) = 6)
                Case 8
                    q = (monthNo(thisDay) = 7)
                Case 9
                    q = (monthNo(thisDay) = 8)
                Case 10
                    q = (monthNo(thisDay) = 9)
                Case 11
                    q = (monthNo(thisDay) = 10)
                Case 12
                    q = (monthNo(thisDay) = 11)
                Case 13
                    q = (monthNo(thisDay) <> 0)
                Case 14
                    q = (monthNo(thisDay) <> 1)
                Case 15
                    q = (monthNo(thisDay) <> 2)
                Case 16
                    q = (monthNo(thisDay) <> 3)
                Case 17
                    q = (monthNo(thisDay) <> 4)
                Case 18
                    q = (monthNo(thisDay) <> 5)
                Case 19
                    q = (monthNo(thisDay) <> 6)
                Case 20
                    q = (monthNo(thisDay) <> 7)
                Case 21
                    q = (monthNo(thisDay) <> 8)
                Case 22
                    q = (monthNo(thisDay) <> 9)
                Case 23
                    q = (monthNo(thisDay) <> 10)
                Case 24
                    q = (monthNo(thisDay) <> 11)
                Case Else
                    q = False
                    Stop
            End Select
        End With
        QExeMonth = q
    End Function
    Public Function StrSign(x As Single) As String
        If x >= 0.0 Then
            StrSign = "+" & Format$(x, "000.000")
        Else
            StrSign = "~" & Format$(-x, "000.000")
        End If
    End Function
    Public Function qHitPatternZ(ByRef R As Results, ByRef TS() As Trades, ByRef trd As Integer, ByRef offset As Integer) As Boolean
        Static qz As Boolean
        If trd > 1 Then
            Select Case RMain.SP.bZScoreMode.Idx
                Case 0 'zA---
                    qz = True
                Case 1 'zBp--
                    qz = (RMainTS(trd - offset).Profit0 > 0.01)
                Case 2 'zCn--
                    qz = (RMainTS(trd - offset).Profit0 <= 0.01)
                Case 3 'zDpp-
                    qz = (RMainTS(trd - offset).Profit0 > 0.01) And (RMainTS(trd - offset - 1).Profit0 > 0.01)
                Case 4 'zEnn-
                    qz = (RMainTS(trd - offset).Profit0 <= 0.01) And (RMainTS(trd - offset - 1).Profit0 <= 0.01)
                Case 5 'zFppp
                    qz = (RMainTS(trd - offset).Profit0 > 0.01) And (RMainTS(trd - offset - 1).Profit0 > 0.01) And (RMainTS(trd - offset - 2).Profit0 > 0.01)
                Case 6 'zGnnn
                    qz = (RMainTS(trd - offset).Profit0 <= 0.01) And (RMainTS(trd - offset - 1).Profit0 <= 0.01) And (RMainTS(trd - offset - 2).Profit0 <= 0.01)
                Case 7 'zHpn-
                    qz = (RZA__TS(trd - offset).Profit0 > 0.01) And (RZA__TS(trd - offset - 1).Profit0 <= 0.01)
                Case 8 'zInp-
                    qz = (RZA__TS(trd - offset).Profit0 <= 0.01) And (RZA__TS(trd - offset - 1).Profit0 > 0.01)
                Case 9 'zJpnn
                    qz = (RZA__TS(trd - offset).Profit0 > 0.01) And (RZA__TS(trd - offset - 1).Profit0 <= 0.01) And (RZA__TS(trd - offset - 2).Profit0 <= 0.01)
                Case 10 'zKpnp
                    qz = (RZA__TS(trd - offset).Profit0 > 0.01) And (RZA__TS(trd - offset - 1).Profit0 <= 0.01) And (RZA__TS(trd - offset - 2).Profit0 > 0.01)
                Case 11 'zLnpp
                    qz = (RZA__TS(trd - offset).Profit0 <= 0.01) And (RZA__TS(trd - offset - 1).Profit0 > 0.01) And (RZA__TS(trd - offset - 2).Profit0 > 0.01)
                Case 12 'zMnpn
                    qz = (RZA__TS(trd - offset).Profit0 <= 0.01) And (RZA__TS(trd - offset - 1).Profit0 > 0.01) And (RZA__TS(trd - offset - 2).Profit0 <= 0.01)
                Case 13 'zNppn
                    qz = (RZA__TS(trd - offset).Profit0 > 0.01) And (RZA__TS(trd - offset - 1).Profit0 > 0.01) And (RZA__TS(trd - offset - 2).Profit0 <= 0.01)
                Case 14 'zOnnp
                    qz = (RZA__TS(trd - offset).Profit0 <= 0.01) And (RZA__TS(trd - offset - 1).Profit0 <= 0.01) And (RZA__TS(trd - offset - 2).Profit0 > 0.01)
                Case Else
                    qz = False
                    Stop
            End Select
        Else
            qz = False
            Stop
        End If
        qHitPatternZ = qz
    End Function
    REM
    Public Function QBExeEntry(ByRef thisDay As Integer) As Boolean
        QBExeEntry = False
        '        newTrade.bActualPrice = newTrade.bSignal.actualPrice
        With newTrade
            .bAmt = 0.0
            newTrade.Shares = 0.0
            '            .bentry.Signal_ = R.SP.bentry.Text & "td" & dtStr2(R.thisDay)
            Select Case RMain.SP.bEntry.Idx
                Case 0
                    .bDayNo = thisDay
                    .bDate = dtStr2(.bDayNo)
                    newTrade.bPrice = cl(.bDayNo) '.bSignal.trigPrice
                    .tradeDayNo = .bDayNo
                    .bEntry.executePrice = .bPrice
                    .bTrigger.triggerDay = .bDayNo
                    .bTrigger.Date_ = dtStr2(.bTrigger.triggerDay)
                    .bTrigger.executePrice = .bSignal.triggerPrice
                    .bSignal.executeDay = .bDayNo
                    .bSignal.executeDate = dtStr2(.bSignal.executeDay)
                    .bSignal.executePrice = .bSignal.triggerPrice
                    .bEntry.executeDay = .bDayNo
                    .bEntry.executeDate = dtStr2(.bDayNo)
                    .BuyOnOpen.qHit = False
                    .BuyOnClose.qHit = True
                    QBExeEntry = True
                    .bAmt = 1000.0
                    newTrade.Shares = .bAmt / .bPrice
                Case 1
                    newTrade.bDayNo = thisDay
                    newTrade.bDate = dtStr2(newTrade.bDayNo)
                    newTrade.bPrice = op(newTrade.bDayNo)
                    .tradeDayNo = .bDayNo
                    newTrade.bEntry.executePrice = .bPrice
                    .bTrigger.triggerDay = .bDayNo
                    .bTrigger.executeDate = dtStr2(.bTrigger.triggerDay)
                    .bTrigger.executePrice = .bSignal.triggerPrice
                    .bSignal.executeDay = .bDayNo
                    .bSignal.executeDate = dtStr2(.bSignal.executeDay)
                    .bSignal.executePrice = op(.bDayNo)
                    .bEntry.executeDay = .bDayNo
                    .bEntry.executeDate = dtStr2(.bDayNo)
                    .BuyOnOpen.qHit = True
                    QBExeEntry = True
                    '                   If .bPrice = 0.0 Then Stop
                    .bAmt = 1000.0
                    newTrade.Shares = .bAmt / .bPrice
                    If newTrade.bPrice = 0 Then
                        QBExeEntry = False
                        ' Stop
                    End If
                Case 2
                    .bDayNo = thisDay
                    .bDate = dtStr2(.bDayNo)
                    .bPrice = cl(.bDayNo)
                    .tradeDayNo = .bDayNo
                    .bEntry.executePrice = .bPrice
                    .bTrigger.triggerDay = .bDayNo
                    .bTrigger.executeDate = dtStr2(.bTrigger.triggerDay)
                    .bTrigger.executePrice = .bSignal.triggerPrice
                    .bSignal.executeDay = .bDayNo
                    .bSignal.executeDate = dtStr2(.bSignal.executeDay)
                    .bSignal.executePrice = op(.bDayNo)
                    .bEntry.executeDay = .bDayNo
                    .bEntry.executeDate = dtStr2(.bDayNo)
                    .BuyOnNextClose.qHit = True
                    QBExeEntry = True
                    'If .bPrice = 0.0 Then Stop
                    .bAmt = 1000.0
                    newTrade.Shares = .bAmt / .bPrice
                Case 3
                    If thisDay >= Counters.end_day Then
                        .bDayNo = thisDay
                    Else
                        .bDayNo = thisDay + 1
                    End If
                    .bPrice = op(newTrade.bDayNo)
                    .tradeDayNo = thisDay
                    .bEntry.executePrice = op(.bDayNo)
                    .bPrice = .bEntry.executePrice
                    .bTrigger.triggerDay = .bDayNo
                    .bTrigger.executeDate = dtStr2(.bTrigger.triggerDay)
                    .bTrigger.executePrice = .bSignal.triggerPrice
                    .bSignal.executeDay = .bDayNo
                    .bSignal.executeDate = dtStr2(.bSignal.executeDay)
                    .bSignal.executePrice = op(.bDayNo)
                    .bEntry.executeDay = .bDayNo
                    .bEntry.executeDate = dtStr2(.bDayNo)
                        .BuyOnNextOpen.qHit = True
                        QBExeEntry = True
                        .bAmt = 1000.0
                        newTrade.Shares = .bAmt / .bPrice
                    If .bPrice = 0.0 Then
                        Stop
                        QBExeEntry = False
                    End If
                Case 4
                    If thisDay >= Counters.end_day Then
                        newTrade.bDayNo = thisDay
                    Else
                        .bDayNo = thisDay + 1
                    End If
                    .bPrice = cl(.bDayNo)
                    .tradeDayNo = thisDay
                    newTrade.bEntry.executePrice = cl(newTrade.bDayNo)
                    newTrade.bPrice = .bEntry.executePrice
                    .bTrigger.triggerDay = .bDayNo
                    .bTrigger.executeDate = dtStr2(.bTrigger.triggerDay)
                    .bTrigger.executePrice = .bSignal.triggerPrice
                    .bSignal.executeDay = .bDayNo
                    .bSignal.executeDate = dtStr2(.bSignal.executeDay)
                    .bSignal.executePrice = op(.bDayNo)
                    .bEntry.executeDay = .bDayNo
                    .bEntry.executeDate = dtStr2(.bDayNo)
                    .BuyOnNextClose.qHit = True
                    QBExeEntry = True
                    .bAmt = 1000.0
                    newTrade.Shares = .bAmt / .bPrice
                    If .bPrice = 0.0 Then
                        Stop
                        QBExeEntry = False
                    End If
                Case 5
                    If thisDay <= Counters.end_day Then
                        .bDayNo = thisDay + 1
                        .tradeDayNo = thisDay
                        .bEntry.executePrice = op(.bDayNo)
                        .bPrice = .bEntry.executePrice
                        .bTrigger.triggerDay = .bDayNo - 1
                        .bTrigger.triggerPrice = op(.bTrigger.triggerDay)
                        .bTrigger.triggerDate = dtStr2(.bTrigger.triggerDay)
                        If .bEntry.executePrice <= .bTrigger.triggerPrice Then
                            QBExeEntry = True
                            .BuyOnOpen.qHit = True
                            .BuyOnClose.qHit = False
                            .BuyOnNextOpen.qHit = False
                            .BuyOnNextClose.qHit = False
                            .bTrigger.executePrice = .bSignal.triggerPrice
                            .bSignal.executeDay = .bDayNo
                            .bSignal.executeDate = dtStr2(.bSignal.executeDay)
                            .bSignal.executePrice = op(.bDayNo)
                            .bEntry.executeDay = .bDayNo
                            .bEntry.executeDate = dtStr2(.bDayNo)
                            .BuyOnNextOpen.qHit = True
                            QBExeEntry = True
                            .bAmt = 1000.0
                            newTrade.Shares = .bAmt / .bPrice
                        Else
                            QBExeEntry = False
                            Stop
                        End If
                    End If
                Case 6
                    If thisDay <= Counters.end_day Then
                        .bDayNo = thisDay + 1
                        .bPrice = cl(.bDayNo)
                        .tradeDayNo = thisDay
                        .bEntry.executePrice = cl(.bDayNo)
                        .bPrice = .bEntry.executePrice
                        .bTrigger.triggerDay = .bDayNo
                        .bTrigger.executeDate = dtStr2(.bTrigger.triggerDay)
                        .bTrigger.executePrice = .bSignal.triggerPrice
                        .bSignal.executeDay = .bDayNo
                        .bSignal.executeDate = dtStr2(.bSignal.executeDay)
                        .bSignal.executePrice = op(.bDayNo)
                        .bEntry.executeDay = .bDayNo
                        .bEntry.executeDate = dtStr2(.bDayNo)
                        .BuyOnClose.qHit = True
                        QBExeEntry = True
                        .bAmt = 1000.0
                        newTrade.Shares = .bAmt / .bPrice
                    Else
                        Stop
                    End If
                Case Else
                    Stop
            End Select
            'If newTrade.bPrice = 0.0 Then Stop
        End With
    End Function
    Private Sub SellIt(ByRef R As Results, ByRef R_TS() As Trades)
        Static qSell As Boolean
        With newTrade
            newTrade.qPastEndDay = False
            If .Shares = 0.0 Then Stop
            qSell = False
            .SellOnExpired.qHit = False
            Select Case R.SP.sEntry.Idx
                Case 1
                    .sDayNo = newTrade.bDayNo
                Case 2
                    .sDayNo = .bDayNo
                Case 3
                    .sDayNo = .bDayNo + 1
                Case 4
                    .sDayNo = .bDayNo + 1
                Case 5
                    .sDayNo = .bDayNo + 1
                Case 6
                    .sDayNo = .bDayNo + 1
                Case 7
                    .sDayNo = .bDayNo + 2
                Case 8
                    .sDayNo = .bDayNo + 2
                Case Else
                    Stop
            End Select
            GoTo exit8
expired:
            .Expired.qHit = True
            GoTo exitt
exit8:
            .Expired.qHit = True
            newTrade.sDayNo = .sDayNo + newTrade.maxDH
exitt:
            If newTrade.sDayNo > Counters.end_day Then
                newTrade.sDayNo = Counters.end_day
                newTrade.qPastEndDay = True
                GoTo exitt
            End If
            'If newTrade.sExecute.executePrice = 0.0 Then Stop
            .sDate = dtStr2(.sDayNo)
            .sDOW.Text = Strings.Left(DOWStr(newTrade.sDayNo), 1)
            .bDOW.Text = Strings.Left(DOWStr(newTrade.bDayNo), 1)
            .bDate = dtStr2(.bDayNo)
            .sEntry.executeDay = .sDayNo
            .sEntry.executeDate = .sDate
            newTrade.bEntry.executeDate = .bDate
            Call TradeSellEntry()
            newTrade.bOpen = op(newTrade.bDayNo)
            newTrade.bHigh = hi(newTrade.bDayNo)
            newTrade.bLow = lo(newTrade.bDayNo)
            newTrade.bClose = cl(newTrade.bDayNo)
            newTrade.sOpen = op(newTrade.sDayNo)
            newTrade.sHigh = hi(newTrade.sDayNo)
            newTrade.sLow = lo(newTrade.sDayNo)
            newTrade.sClose = cl(newTrade.sDayNo)
            R_TS(RMain.SS.Main.Trades) = newTrade
            Select Case R.SS.Main.Trades
                Case >= 6
                    newTrade.Profit1 = RMainTS(RMain.SS.Main.Trades - 1).Profit0
                    newTrade.Profit2 = RMainTS(RMain.SS.Main.Trades - 2).Profit0
                    newTrade.Profit3 = RMainTS(RMain.SS.Main.Trades - 3).Profit0
                    newTrade.Profit4 = RMainTS(RMain.SS.Main.Trades - 4).Profit0
                Case Else
                    newTrade.Profit1 = 0.0
                    newTrade.Profit2 = 0.0
                    newTrade.Profit3 = 0.0
                    newTrade.Profit4 = 0.0
            End Select
            newTrade.TradeNo = RMain.SS.Main.Trades
        End With
exit1:
    End Sub
    Private Sub SellItWithTargetProfit(ByRef R As Results, ByRef R_TS() As Trades)
        Static qSell As Boolean
        qSell = False
        With newTrade
            If newTrade.bPrice = 0.0 Then Stop
            If newTrade.Shares = 0.0 Then Stop
            newTrade.sTP.V = RMain.SP.TP.V
            newTrade.sTP.actualDay = newTrade.bDayNo
            newTrade.sTP.actualPrice = newTrade.bPrice
            newTrade.sTP.triggerDay = newTrade.bDayNo
            newTrade.sTP.triggerPrice = newTrade.bPrice * ((100 + newTrade.sTP.V) / 100)
            newTrade.sTP.qHit = False
            newTrade.qPastEndDay = False
            newTrade.SellOnExpired.qHit = False
            newTrade.qSellOnTP = False
LoopDH:
            If newTrade.sTP.actualDay >= Counters.end_day Then GoTo expired
            newTrade.sTP.executeDay = newTrade.sTP.actualDay
            newTrade.sDayNo = newTrade.sTP.executeDay
            newTrade.sTP.actualPrice = hi(newTrade.sTP.actualDay)
            newTrade.sTP.qHit = (newTrade.sTP.actualPrice >= newTrade.sTP.triggerPrice)
            If newTrade.sTP.qHit Then
                newTrade.qSellOnTP = True
                newTrade.sTP.executePrice = newTrade.sTP.triggerPrice
                newTrade.TPHits = newTrade.TPHits + 1
                R.SS.Main.sTP.Hits = R.SS.Main.sTP.Hits + 1
                newTrade.DH = newTrade.sDayNo - newTrade.bDayNo + 1
                newTrade.sPrice = newTrade.sTP.triggerPrice
                newTrade.Profit0 = newTrade.sPrice - newTrade.bPrice
                GoTo hitTP
            End If
            .dh1 = .sTP.actualDay - .bDayNo
            If .dh1 > .maxDH Then GoTo exit8
            newTrade.sTP.actualDay = newTrade.sTP.actualDay + 1
            If newTrade.sTP.actualDay >= Counters.end_day Then GoTo expired
            GoTo LoopDH
expired:
            .Expired.qHit = True
            newTrade.TPMisses = newTrade.TPMisses + 1
            If R.SP.sEntry.Idx > 2 Then
                newTrade.sDayNo = newTrade.sDayNo + 1
            End If
            GoTo exitt
exit8:
            .Expired.qHit = True
exitt:
            If newTrade.sTP.actualDay >= Counters.end_day Then
                'Stop
                newTrade.qPastEndDay = True
                GoTo exit1
            End If
            'If newTrade.sExecute.executePrice = 0.0 Then Stop
            newTrade.sDate = dtStr2(newTrade.sDayNo)
            .sDOW.Text = Strings.Left(DOWStr(newTrade.sDayNo), 1)
            .bDOW.Text = Strings.Left(DOWStr(newTrade.bDayNo), 1)
            .bDate = dtStr2(newTrade.bDayNo)
            .sEntry.executeDay = newTrade.sDayNo
            .sEntry.executeDate = .sDate
            newTrade.bEntry.executeDate = .bDate
hitTP:
            If newTrade.sDayNo <= 1 Then Stop
            If Not newTrade.qSellOnTP Then
                Call TradeSellEntry()
            Else
                ' newTrade.sPrice = newTrade.TP
                .SellOnClose.qHit = False
                newTrade.sDayNo = .sDayNo
                If newTrade.DH < 0 Then Stop
                If newTrade.DH = 0 Then newTrade.DH = 1
                .dh1 = .DH + 1
                newTrade.sDate = dtStr2(newTrade.sDayNo)
                newTrade.sAmt = newTrade.Shares * newTrade.sPrice
                newTrade.Profit0 = newTrade.sAmt - newTrade.bAmt
                newTrade.sEntry.Text = RMain.SP.sEntry.Text & ":" & Format$(newTrade.sPrice, "000.000")
                If newTrade.Profit0 > 80 Then
                    newTrade.Profit0 = 80
                End If
                If newTrade.Profit0 < -80 Then
                    newTrade.Profit0 = -80
                End If
                Call CalculateTradeStats()
                '    Call TradeSellEntry()
            End If
            '    If newTrade.Profit0 = 0.0 Then Stop
            R_TS(R.SS.Main.Trades) = newTrade
            If R.SS.Main.Trades > 3 Then
                newTrade.Profit1 = R_TS(R.SS.Main.Trades - 1).Profit0
                .Profit2 = R_TS(R.SS.Main.Trades - 2).Profit0
                .Profit3 = R_TS(R.SS.Main.Trades - 3).Profit0
            Else
                .Profit1 = 0.0
                .Profit2 = 0.0
                '                newTrade.Profit4 = R.TS(R.SS.Trades - 3).Profit0
            End If
            newTrade.TradeNo = R.SS.Main.Trades
        End With
exit1:
    End Sub
    Public Sub CountTD(ByRef dDay As Integer)
        With newTrade.bTD1
            .Idx = RMain.SP.bTD1.Idx
            .dbIdx = RMain.SP.bTD1.dbIdx
            .db = RMain.SP.bTD1.dbText
            RMain.SS.Main.bTD1.tot = RMain.SS.Main.bTD1.tot + 1
            If .Idx > 0 Then
                '             If .actualPrice = 0 Then Stop
                .qHit = QTD_(newTrade.bTD1, dDay, newTrade.bTD1.db)
                .Signal = RMain.SP.bTD1.Text1 & "//" & "on" & .actualDate & "=" & Format$(.actualPrice, "000.000") & "::" &
            .triggerDate & "=" & Format$(.triggerPrice, "000.000") &
              "EXE@" & .executeDate & "=" & Format$(.executePrice, "000.000")
            Else
                .qHit = True
                .Signal = RMain.SP.bTD1.Text1 & "//" & "on" &
            "--------" & "=" & Format$(.actualPrice, "000.000") & "::" &
            "--------" & "=" & Format$(.triggerPrice, "000.000") &
              "EXE@" & "--------" & "=" & Format$(.executePrice, "000.000")
            End If
        End With
        With newTrade.bTD2
            .Idx = RMain.SP.bTD2.Idx
            .dbIdx = RMain.SP.bTD2.dbIdx
            .db = RMain.SP.bTD2.dbText
            RMain.SS.Main.bTD2.tot = RMain.SS.Main.bTD2.tot + 1
            If .Idx > 0 Then
                .qHit = QTD_(newTrade.bTD2, dDay, newTrade.bTD2.db)
                .Signal = RMain.SP.bTD2.Text & "//" & "on" & .actualDate & "=" & Format$(.actualPrice, "000.000") & "::" &
            .triggerDate & "=" & Format$(.triggerPrice, "000.000") &
              "EXE@" & .executeDate & "=" & Format$(.executePrice, "000.000")
            Else
                .qHit = True
                .Signal = RMain.SP.bTD2.Text1 & "//" & "on" &
            "--------" & "=" & Format$(.actualPrice, "000.000") & "::" &
            "--------" & "=" & Format$(.triggerPrice, "000.000") &
              "EXE@" & "--------" & "=" & Format$(.executePrice, "000.000")
            End If
        End With
    End Sub
    Private Sub TradeSellEntry()
        Static lPr As Single
        With newTrade
            .SellOnClose.qHit = False
            .SellOnExpired.qHit = False
            '            .SExeTg.qHit = False
            '            .SExeOpen.qHit = False
            If newTrade.sDayNo <= 1 Then Stop
            Select Case RMain.SP.sEntry.Idx
                Case 0
                    .SellOnClose.qHit = True
                    newTrade.sPrice = op(newTrade.sDayNo)
                Case 1
                    .SellOnOpen.qHit = True
                    newTrade.sPrice = op(newTrade.sDayNo)
                Case 2
                    .SellOnClose.qHit = True
                    newTrade.sPrice = cl(newTrade.sDayNo)
                Case 3
                    .SellOnOpen.qHit = True
                    newTrade.sPrice = op(newTrade.sDayNo)
                Case 4
                    .SellOnClose.qHit = True
                    newTrade.sPrice = cl(newTrade.sDayNo)
                Case 5
                    .sDayNo = .sDayNo + 1
                    .SellOnOpen.qHit = True
                    newTrade.sDayNo = .sDayNo
                    If newTrade.sDayNo > Counters.end_day Then newTrade.sDayNo = Counters.end_day
                    newTrade.sPrice = op(newTrade.sDayNo)
                Case 6
                    .sDayNo = .sDayNo + 1
                    .SellOnClose.qHit = True
                    newTrade.sDayNo = .sDayNo + 1
                    If newTrade.sDayNo > Counters.end_day Then newTrade1.sDayNo = Counters.end_day
                    newTrade.sPrice = cl(newTrade.sDayNo)
                Case 7
                    .SellOnOpen.qHit = True
                    newTrade.sDayNo = .sDayNo
                    newTrade.sPrice = op(newTrade.sDayNo)
                Case 8
                    .SellOnClose.qHit = True
                    newTrade.sDayNo = .sDayNo + 1
                    newTrade.sPrice = cl(newTrade.sDayNo)
                Case Else
                    Stop
            End Select
            newTrade.DH = newTrade.sDayNo - newTrade.bDayNo
            If newTrade.DH < 0 Then Stop
            If newTrade.DH = 0 Then newTrade.DH = 1
            .dh1 = .DH + 1
            newTrade.sDate = dtStr2(newTrade.sDayNo)
            newTrade.sAmt = newTrade.Shares * newTrade.sPrice
            newTrade.Profit0 = newTrade.sAmt - newTrade.bAmt
            newTrade.sEntry.Text = RMain.SP.sEntry.Text & ":" & Format$(newTrade.sPrice, "000.000")
            If newTrade.Profit0 > 80 Then
                newTrade.Profit0 = 80
            End If
            If newTrade.Profit0 < -80 Then
                newTrade.Profit0 = -80
            End If
            Call CalculateTradeStats()
            '        If R.SS.Trades = 113 Then Stop
            Counters.totTradeNo = Counters.totTradeNo + 1
            .totTradeNum = Counters.totTradeNo
            .ldDayNo = Counters.end_day
            .ldDate = Counters.end_Date
            RMain.SS.Main.Trades = RMain.SS.Main.Trades + 1
            If RMain.SS.Main.Trades = 1 Then
                .lastProfit = 0.0
            Else
                newTrade.lastProfit = lPr
            End If
            lPr = newTrade.Profit0
            '            If .sTD1.qHit Then
            'R.SS.Main.sTD1.Hits = R.SS.Main.sTD1.Hits + 1
            'End If
            RMain.SS.Main.sEntryStats.Hits = RMain.SS.Main.sEntryStats.Hits + 1
            '            .sTD1.Signal_ = R.SP.sTD1.Text & "-" & Format$(R.SP.sTD1.dbV, "00") & "-" & dtStr2(.sDayNo) & "miss"
            '           .sSignal.Signal_ = R.SP.sSignal.Text & "exeon:" & _
            '              dtStr2(.sDayNo) & "@" & Format$(.sPrice, "000.000")
            '         .sExecute.Signal_ = R.SP.sExecute.Text & "exeon" & dtStr2(.sDayNo)
        End With
    End Sub
    Private Sub CalculateTradeStats()
        Static ldy As Integer
        With RMain.SS
            .Main.drawDown.max = 0.0
            For ldy = newTrade.bDayNo + 1 To newTrade.sDayNo
                If .Main.drawDown.max > (lo(ldy) - newTrade.bPrice) * newTrade.Shares Then
                    .Main.drawDown.max = (lo(ldy) - newTrade.bPrice) * newTrade.Shares
                End If
            Next ldy
            .Main.drawDown.tot = RMain.SS.Main.drawDown.tot + RMain.SS.Main.drawDown.max
            newTrade.totDrawDn = newTrade.totDrawDn + .Main.drawDown.max
            .Main.drawUp.max = 0.0
            For ldy = newTrade.bDayNo + 1 To newTrade.sDayNo
                If .Main.drawUp.max > (hi(ldy) - newTrade.bPrice) * newTrade.Shares Then
                    .Main.drawUp.max = (hi(ldy) - newTrade.bPrice) * newTrade.Shares
                End If
            Next ldy
            .Main.drawUp.tot = .Main.drawUp.tot + .Main.drawUp.max
            newTrade.bEntry.Text = RMain.SP.bEntry.Text
            newTrade.sEntry.Text = RMain.SP.sEntry.Text
        End With
    End Sub
    Public Function CalcZScore(ByRef R As Results,
                                ByRef TS() As Trades,
                                 ByRef startTrade As Integer,
                                  ByRef endTrade As Integer) As Single
        Dim tr As Integer, tmpProfit As Single, lastProfit As Single, noTrades As Integer
        With R.SS.Main
            noTrades = endTrade - startTrade + 1
            If noTrades > 12 Then
                R.SS.Main.W = 0
                R.SS.Main.L = 0
                .Runs = 0
                tmpProfit = 0.0
                lastProfit = TS(startTrade - 1).Profit0
                For tr = startTrade To endTrade
                    tmpProfit = RMainTS(tr).Profit0
                    '                If tmpProfit = 0.0 Then Stop
                    If tmpProfit <= 0.1 And lastProfit > 0.1 Or tmpProfit > 0.1 And lastProfit <= 0.1 Then
                        .Runs = .Runs + 1
                    End If
                    If tmpProfit > 0.1 Then
                        .W = .W + 1
                    Else
                        .L = .L + 1
                    End If
                Next
                If .W = 0 Then
                    .W = 1
                End If
                If .L = 0 Then
                    .L = 1
                End If
                .P = 2 * R.SS.Main.W * R.SS.Main.L
                If noTrades = 1 Then Stop
                R.SS.Main.zScore.V = (noTrades * (.Runs - 0.5) - .P) / ((.P * (.P - noTrades)) / (noTrades - 1)) ^ (1 / 2)
                R.SS.Main.expRuns = ((2 * .W * .L) / (.W + .L)) + 1
                lastProfit = tmpProfit
                CalcZScore = .zScore.V
            Else
                .zScore.V = 0.0
            End If
        End With
        CalcZScore = R.SS.Main.zScore.V
    End Function
    Private Sub CalcTradeLoop(ByRef R As Results, ByRef TS() As Trades, StartTrade As Integer, EndTrade As Integer)
        Dim trx As Integer, tmpPeak As Single
        With R.SS.Main
            .Q.avg = 0.0
            .Q.tot = 0.0
            .DH.tot = 0
            .DH.avg = 0.0
            .Profits.tot = 0.0
            .Profits.avg = 0.0
            .DH1.tot = 0
            .Trades = 0
            If EndTrade = 0 Then Stop
            For trx = StartTrade To EndTrade
                If TS(trx).DH = 0 Then TS(trx).DH = 1
                R.tmpProfit = TS(trx).Profit0
                .Trades = .Trades + 1
                .DH.tot = .DH.tot + TS(trx).DH
                .Profits.tot = .Profits.tot + R.tmpProfit
                .Profits1.tot = .Profits1.tot + R.tmpProfit1
                If R.tmpProfit > R.profitThreshhold Then
                    R.posOutliers = R.posOutliers + 1
                    TS(trx).Profit0 = 0.0 'R.profitThreshhold
                    R.tmpProfit = R.profitThreshhold
                End If
                If R.tmpProfit < -R.profitThreshhold Then
                    R.negOutliers = R.negOutliers + 1
                    TS(trx).Profit0 = 0.0 ' -R.profitThreshhold
                    R.tmpProfit = -R.profitThreshhold
                End If
                If trx > 1 Then
                    If TS(trx).Profit0 < 0.0 And TS(trx - 1).Profit0 >= 0.0 Or
                    TS(trx).Profit0 > 0.0 And TS(trx - 1).Profit0 <= 0.0 Then
                        R.SS.Main.Runs = R.SS.Main.Runs + 1
                    End If
                End If
                If R.tmpProfit > 0.0 Then
                    .W = .W + 1
                    .Winners.tot = .Winners.tot + R.tmpProfit
                    .WProfits.tot = .WProfits.tot + R.tmpProfit
                    .WDH.tot = .WDH.tot + TS(trx).DH
                Else
                    .L = .L + 1
                    .Losers.tot = .Losers.tot + R.tmpProfit
                    .LProfits.tot = .LProfits.tot + R.tmpProfit
                    .LDH.tot = .LDH.tot + TS(trx).DH
                End If
                If .Profits.tot < .Profits.min Then
                    .Profits.min = .Profits.tot
                End If
                If .Profits.tot > .Profits.max Then
                    .Profits.max = .Profits.tot
                End If
                TS(trx).totProfit = .Profits.tot
                TS(trx).avgProfit = .Profits.tot / .Trades
                TS(trx).totDH = .DH.tot
                TS(trx).avgDH = .DH.tot / .Trades
                TS(trx).Quantum = TS(trx).Profit0 / TS(trx).DH / 10
                TS(trx).totQuantum = TS(trx).totProfit / TS(trx).totDH / 10
                If Single.IsNaN(TS(trx).totQuantum) Then Stop
                If .Profits.tot > .Peak Then
                    .Peak = .Profits.tot
                    .trdPeak = trx
                End If
                If .Profits.tot < .Trough Then
                    .Trough = .Profits.tot
                    .trdTrough = trx
                End If
                tmpPeak = .Peak - .Profits.tot
                If tmpPeak > .drawDown.max Then
                    .drawDown.max = tmpPeak
                End If
                If .Q.avg < .Q.min Then
                    .Q.min = .Q.avg
                End If
                If .Q.avg > .Q.max Then
                    .Q.max = .Q.avg
                End If
skipp:
            Next trx
            R.SS.Main.Q.avg = R.SS.Main.Profits.tot / R.SS.Main.DH.tot / 10
        End With
    End Sub
    Public Function Calculate_BasicStatistics(ByRef R As Results, ByRef TS() As Trades, ByRef startTrade As Integer, endTrade As Integer) As Single
        Static A As Single, maxRisk As Single, capital As Single, t1 As Single, t2 As Single, Z As Single, Z1 As Single, Z2 As Single
        Static PP As Single, RR As Single, tmpWP As Single, A1 As Single, A2 As Single, PG As Single, trx As Integer, factor1 As Single, factor2 As Single
        Static za As Single, zd As Single, edge As Single, aw As Single, al As Single, e As Single, e2 As Single, probW As Single, probL As Single, P As Single
        Counters.Days = Counters.end_day - Counters.start_Day + 1
        With R.SS.Main
            .Days = Counters.end_day - Counters.start_Day + 1
            .W = 0
            .L = 0
            .Winners.tot = 0
            .Losers.tot = 0
            .DH.tot = 0
            .WDH.tot = 0.0
            .LDH.tot = 0.0
            .WProfits.tot = 0.0
            .LProfits.tot = 0.0
            .Profits.tot = 0.0
            .drawDown.max = -1000.0
            .drawDown.tot = 0.0
            .drawDown.min = 1000.0
            .drawUp.max = -1000.0
            .drawUp.min = 1000.0
            .drawUp.tot = 0.0
            .DH.tot = 0
            .Runs = 0
            R.tmpProfit = TS(1).Profit0
            TS(1).TradeNo = 1
            If TS(1).DH = 0 Then TS(1).DH = 1
            If R.tmpProfit > 0.0 Then
                .W = .W + 1
                .Winners.tot = R.tmpProfit
                .WDH.tot = .WDH.tot + TS(1).DH
            Else
                .L = .L + 1
                .LProfits.tot = R.tmpProfit
                .LDH.tot = .LDH.tot + TS(1).DH
            End If
            TS(1).totProfit = R.tmpProfit
            TS(1).avgProfit = R.tmpProfit
            TS(1).totDH = TS(1).DH
            TS(1).Quantum = R.tmpProfit / TS(1).DH / 10
            TS(1).totQuantum = TS(1).totProfit / TS(1).totDH / 10
            .DH.tot = TS(1).DH
            .Profits.tot = R.tmpProfit
            .Peak = .Profits.tot
            .Trough = .Profits.tot
            .drawDown.max = 0.0
            .drawUp.max = 0.0
            '            R.TS1(1).totProfit = R.tmpProfit1
            '           R.TS1(1).avgProfit = R.tmpProfit1
            '          R.TS1(1).totDH = R.TS1(1).DH
            '         R.TS1(1).Quantum = R.tmpProfit1 / R.TS1(1).DH / 10
            '        R.TS1(1).totQuantum = R.TS1(1).totProfit / R.TS1(1).totDH / 10
            R.negOutliers = 0
            R.posOutliers = 0
            .WProfits.avg = 0.0
            .LProfits.avg = 0.0
            .Q.min = 1000.0
            .Q.max = -1000.0
            .Profits.tot = 0.0
            .Profits.min = 1000.0
            .Profits.max = -1000.0
            R.profitThreshhold = 100.0
            If endTrade = 0 Then Stop
            Call CalcTradeLoop(R, TS, startTrade + 1, endTrade)
            R.SS.Main.avgDH = R.SS.Main.DH.tot / R.SS.Main.Trades
            R.SS.Main.Q.avg = R.SS.Main.Profits.tot / R.SS.Main.DH.tot / 10
            If Single.IsNaN(R.SS.Main.Profits.tot) Then Stop
            If R.SS.Main.Profits.tot = 1000.0 Then Stop
            'If R.SS.Main.Q.avg = -0 Then Stop
            .WDH.avg = .WDH.tot / .W
            .LDH.avg = .LDH.tot / .L
            If .W > 0 Then
                .WProfits.avg = .WProfits.tot / .W
                .Winners.avg = .Winners.tot / .W
            Else
                .WProfits.avg = 0.0
                .Winners.avg = 0
            End If
            If .L > 0 Then
                .LProfits.avg = .LProfits.tot / .L
                .Losers.avg = .Losers.tot / .L
            Else
                .LProfits.avg = 0.0
                .Losers.avg = 0.0
            End If
            If R.SS.Main.Trades > 0 Then
                R.SS.Main.wPcntg = R.SS.Main.W / R.SS.Main.Trades
                R.SS.Main.DH.avg = R.SS.Main.DH.tot / R.SS.Main.Trades
                R.SS.Main.drawDown.avg = R.SS.Main.drawDown.tot / R.SS.Main.Trades
                R.SS.Main.drawUp.avg = R.SS.Main.drawUp.tot / R.SS.Main.Trades
                R.SS.Main.Profits.avg = R.SS.Main.Profits.tot / R.SS.Main.Trades
            Else
                .wPcntg = 0.0
                .DH.avg = 0.0
                .drawDown.avg = 0.0
                .drawUp.avg = 0.0
                .Profits.avg = 0.0
            End If
            .sumSquaredDiff = 0.0
            For trx = startTrade To endTrade
                '                If Math.Abs(tmpR.TS(trx).Profit0) > 80.0 Then Stop
                .sumSquaredDiff = .sumSquaredDiff + (TS(trx).Profit0 - .Profits.avg) ^ 2
            Next trx
            .stdDeviation = (.sumSquaredDiff / .Trades) ^ 0.5
            probW = .wPcntg
            probL = 1 - probW
            aw = .WProfits.avg
            al = .LProfits.avg
            'If .LProfits.avg = 0.0 Then Stop
            A = - .WProfits.tot / .LProfits.tot
            e = .Profits.avg
            e2 = probW * (aw ^ 2) - (probL) * (al ^ 2)
            P = 0.5 + (e / (2 * (e2 ^ 0.5)))
            RR = ((1 - P) / P)
            If .LProfits.tot > 0.0 Then
                RR = ((1 - 1))
                '  risk_of_ruin = ((1 - edge) / (1 + edge)) ^ Capital_Units
                .ROR1 = 0.0
                If .wPcntg > 0.5 Then
                    ' A = 1.0 - .wPcntg
                    .ROR1 = ((1.0 - A) / (1.0 + A)) ^ 10
                Else
                    .ROR1 = 100.0
                End If
                If A <> 0.0 Then
                    .ROR1 = - .WProfits.tot / .LProfits.tot
                Else
                    .ROR1 = 0.0
                End If
            Else
                .ROR1 = 1.0
            End If
            .ROR2 = 0.0
            maxRisk = 0.05
            capital = 1000.0
            '  (1-(W-L))/(1+(W-L))^ U 
            If .wPcntg > 0.0 Then
                If .wPcntg > 0.5 Then
                    .ROR2 = ((1 - (.wPcntg - (1 - .wPcntg))) / (1 + (.wPcntg - (1 - .wPcntg)))) ^ 10
                Else
                    Z = 1 - .wPcntg
                    .ROR2 = ((1 - (Z - (1 - Z))) / (1 + (Z - (1 - Z)))) ^ 10
                End If
            Else
                .ROR2 = 0.0
            End If
            '            .ROR2_ = (((A + 1) * R.SS.wPcntg) - 1) / A
            'R =e^(-(2*0.06)/0.13*(ln(1-0.5)/(ln(1-0.13)))
            Z = 0.25
            za = Math.Abs(.Q.avg)
            zd = .stdDeviation
            '        RR = (1 - Z) ^ (-2 * za / (zd * Math.Log(1 - zd)))
            RR = 2.71 ^ (-(2 * za / zd * Math.Log(1 - zd)))
            .ROR3 = 0.0
            If .wPcntg > 0.5 Then
                edge = .wPcntg * 100 - 50
                .ROR3 = ((100 - edge) / (100 + edge)) ^ 10
            Else
                edge = (1 - .wPcntg) * 100 - 50
                .ROR3 = ((100 - edge) / (100 + edge)) ^ 10
            End If
            .ROR3 = 0.0
            If .Q.avg > 0.01 Then
                A1 = .wPcntg * ((.WProfits.avg / capital) ^ 2)
                A2 = (1.0 - .wPcntg) * ((.LProfits.avg / capital) ^ 2)
                A = (A1 + A2) ^ 0.5
                Z1 = .wPcntg * (.WProfits.avg) / capital
                Z2 = Math.Abs((1.0 - .wPcntg) * (.LProfits.avg / capital))
                Z = Z1 - Z2
                If Z > 0 Then
                    PP = 0.5 * (1.0 + (Z / A))
                    RR = ((1.0 - PP / PP) ^ (maxRisk / A))
                Else
                    RR = 1.0
                End If
            Else
                tmpWP = (1.0 - .wPcntg)
                A1 = tmpWP * ((- .LProfits.avg / capital) ^ 2)
                A2 = (1.0 - tmpWP) * ((- .WProfits.avg / capital) ^ 2)
                A = (A1 + A2) ^ 0.5
                Z1 = tmpWP * (- .LProfits.avg / capital)
                Z2 = Math.Abs(1.0 - tmpWP) * (.WProfits.avg / capital)
                Z = Z1 - Z2
                If Z > 0.0 Then
                    PP = 0.5 * (1.0 + (Z / A))
                    RR = ((1.0 - PP / PP) ^ (maxRisk / A))
                Else
                    RR = 1.0
                End If
            End If
            .ROR3 = RR
            If R.SS.Main.Profits.avg > 0.1 Then
                tmpWP = .wPcntg
            Else
                tmpWP = (1 - .wPcntg)
            End If
            RR = ((1.0 - tmpWP) / (1.0 + tmpWP)) ^ 20
            PG = 0.1
            If tmpWP < 1.0 And tmpWP > 0.0 Then
                RR = ((((1.0 + tmpWP) / (1.0 - tmpWP)) ^ PG) - 1.0) / (((1.0 - tmpWP)) ^ (1.0 + PG) - 1.0)
            Else
                RR = 0.0
            End If
            If R.SS.Main.Winners.tot > 0 And R.SS.Main.Losers.tot > 0 Then
                R.SS.Main.Kelly = R.SS.Main.wPcntg - ((1 - R.SS.Main.wPcntg) / (R.SS.Main.Winners.tot / (-R.SS.Main.Losers.tot)))
            Else
                R.SS.Main.Kelly = 0.0
            End If
            ' R.SS.Kelly_ = (R.SS.wPcntg / R.SS.LProfit.Average) - ((1 - R.SS.wPcntg) / R.SS.WProfit.Average)
            .ROR4 = .Kelly
            .ROR4 = 0.0
            '            Kelly % = W – [(1 – W) / R]
            .COV.V = 0.0
            If QFE1.QCalcCov.Checked Then
                For trx = 2 To .Trades
                    t1 = TS(trx).Profit0
                    t2 = TS(trx - 1).Profit0
                    factor1 = (t1 - R.SS.Main.Profits.avg)
                    factor2 = (t2 - R.SS.Main.Profits.avg)
                    .COV.V = .COV.V + factor1 * factor2
                Next
                .COV.V = .COV.V / (.Trades - 2)
                .Correlation.V = .COV.V / (.stdDeviation * R.SS.Main.stdDeviation)
            Else
                .COV.V = 0
                .Correlation.V = 0
            End If
            '            .zScore.V = CalcZScore(R, TS, R.SS.Main.Trades)
            .absQuantum = Math.Abs(.Q.avg)
            '            tmpR.SS.Main = R.SS.Main
            '            R.SS.Main = R.SS.Main
        End With
        Calculate_BasicStatistics = R.SS.Main.Q.avg
    End Function
    Public Sub Write_Parameters0(ByRef R As Results, ByRef TS() As Trades, ByRef thisDay As Integer)
        Static thisDate As String, thisDate1 As String
        '      Form1.Iters.Text = Format$(Counters.Iteration, "000000")
        thisDate = dtStr2(thisDay)
        thisDate1 = dtStr1(thisDay) ' & DateAndTime.TimeOfDay()
        If (QFE1.qOnLastDay.Checked And QFE1.qSaveOnSignals.Checked) Or (QFE1.qSaveOffSignals2.Checked) And R.SS.Main.Trades > 5 Then
            '                    If R.SS.absQuantum < 0.4 Then Stop
            Counters.threshHold = QFE1.quantumThreshHold.Value
            QFE1.threshholdTxt.Text = Format$(Counters.threshHold, "0.000")
            Counters.THHits = Counters.THHits + 1
            QFE1.threshholdHitstxt_.Text = Format$(Counters.THHits, "000000")
            Call Put_Signal_(R, 1, lastDayTrade)
            Call DoSave(R, TS)
        Else
            Counters.THMisses = Counters.THMisses + 1
            QFE1.threshholdMissestxt_.Text = Format$(Counters.THMisses, "000000")
            Stop
        End If
    End Sub
    Public Sub DoSave(ByRef R As Results, ByRef TS() As Trades)
        Static xxdd As Integer
        Dim qSave As Boolean
        qSave = QFE1.qSaveOnSignals.Checked And QFE1.qSaveOffSignals2.Checked
        If QFE1.qSaveOnSignals.Checked And Not QFE1.qSaveOffSignals2.Checked Then
            qSave = lastDayTrade.bEntry.qHit
        End If
        If Not QFE1.qSaveOnSignals.Checked And QFE1.qSaveOffSignals2.Checked Then
            qSave = Not lastDayTrade.bEntry.qHit
        End If
        With R
            .TFSignal = ""
            If qSave Then
                Counters.SystemNumber = Counters.SystemNumber + 1
                Call calcPeriodProfits(R, TS)
                If qCriteria(R, TS) Then
                    Counters.signalstradesth = Counters.signalstradesth + 1
                    QFE1.signalsTradesTH.Text = Format$(Counters.signalstradesth, "0000")
                    R.TFSignal = dtStr2(Counters.end_day) & "::"
                    If QFE1.qCalcSigString.Checked Then
                        Select Case .SP.bEntry.Idx
                            Case 0
                                .TFSignal = .TFSignal & "3"
                            Case 1, 2
                                .TFSignal = .TFSignal & "2"
                            Case 3, 4
                                .TFSignal = .TFSignal & "1"
                            Case 5, 6
                                .TFSignal = .TFSignal & "0"
                        End Select
                        For xxdd = Counters.end_day To Counters.end_day - 99 Step -1
                            If qbuyonlastday(xxdd) Then
                                .TFSignal = .TFSignal & "T"
                            Else
                                .TFSignal = .TFSignal & "F"
                            End If
                        Next xxdd
                    End If
                    If R.SS.Main.Trades = 0 Then Stop
                    Call Put_Signals(R, TS, 2)
                    Call Saves(R, TS)
                End If
            End If
        End With
    End Sub

    Private Function qCriteria(ByRef R As Results, ByRef TS() As Trades) As Boolean
        Static q As Boolean
        q = False
        qCriteria = True
        If QFE1.qCriteria_.GetItemChecked(1) Then
            q = R.SS.Main.profits15 > 0.0 And R.SS.Main.profits14 > 0.0 And R.SS.Main.profits13 > 0.0 _
            And (R.SS.Main.profits12 > 0.0) And (R.SS.Main.profits11 > 0.0) And (R.SS.Main.profits10 > 0.0) And
             (R.SS.Main.profits09 > 0.0) And (R.SS.Main.profits08 > 0.0) And (R.SS.Main.profits07 > 0) And (R.SS.Main.profits06 > 0.0)
            qCriteria = q
        End If
    End Function
    Public Sub Saves(ByRef R As Results, ByRef ts_() As Trades)
        With QFE1
            If .qSaveParams0.Checked Then Call Put_Parameters0(R)
            If .qSaveParams1.Checked Then Call Put_Parameters1(R)
            If .qSaveParams2.Checked Then Call Put_Parameters2(R)
            If .qSaveStats0.Checked Then Call Put_Statistics00(R)
            '            If .qSaveStats1.Checked Then Call .Put_Statistics01(R)
            '           If .qSaveStats2.Checked Then Call .Put_Statistics02(R)
            '          If .qSaveDist.Checked Then Call .put_distribution(R)
            '         If .qSaveTrades0.Checked Then Call .Put_Trades0(R)
            '    OleDBC.Connection = conn0
            'If .qSaveTrades1.Checked Then Call Put_Trades1(R, ts_)
            'If .qSaveTrades1a.Checked Then Call Form1.Put_Trades1a(R)
            '      If .qSaveTrades3.Checked Then Call .Put_Trades3(R)
            '     If .qSaveTrades2.Checked Then Call .Put_Trades2(R)
            '    If .qSaveTrades5.Checked Then Call .Put_ldTrades(R)
        End With
    End Sub
    Public Function SetLastDayTrade(ByRef R As Results, ByRef TS() As Trades, ByRef trd As Trades, ByRef thisDay As Integer) As Boolean
        With trd
            .bDOW.qHit = QExeBDOW(thisDay)
            R.SP.bDOW.Text1 = R.SP.bDOW.Text & ":" & Format$(op(thisDay), "000.000") & "!" & Format$(cl(thisDay), "000.000")
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
                .bTD1.qHit = QTD_Text1(.bTD1, thisDay)
                R.SP.bTD1.Text1 = Strings.Left(R.SP.bTD1.Date_, 12) & ":$" &
                 Format$(.bTD1.actualPrice, "000.000") &
                  "on" & .bTD1.Date_ & "&$" & Format$(.bTD1.triggerPrice, "000.000") & "on" & .bTD1.Date_
                .bTD1.Signal = R.SP.bTD1.Text
            Else
                .bTD1.qHit = True
                R.SP.bTD1.Text1 = Strings.Left(R.SP.bTD1.Date_, 12) & ":$" &
                 Format$(.bTD1.actualPrice, "000.000") &
                  "on" & .bTD1.Date_ & "&$" & Format$(.bTD1.triggerPrice, "000.000") & "on" & .bTD1.Date_
                .bTD1.Signal = R.SP.bTD1.Text
            End If
            If .bTD1.qHit Then
                .bTD1.qHiti = 255
            Else
                .bTD1.qHiti = 0
            End If
            ''            .bTD1.Text = R.SP.bTD1.Text1
            .bTD1.Signal = .bTD1.Text & ":"
            .bTD2.Idx = R.SP.bTD2.Idx
            .bTD2.dbIdx = R.SP.bTD2.dbIdx
            .bTD2.db = R.SP.bTD2.dbText
            If .bTD2.Idx > 0 Then
                .bTD2.qHit = QTD_Text1(.bTD2, thisDay)
                R.SP.bTD2.Text1 = Strings.Left(R.SP.bTD2.Text, 12) & ":" & Format$(.bTD2.executePrice, "000.000") & "on" & .bTD2.Text & "&" &
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
            If R.SP.bTrigger.Idx = 0 Then
                '                .qrig.qHit = True
                .bTrigger.Text = R.SP.bTrigger.Text
            Else
                .bTrigger.qHit = QBuyTrigger(trd, thisDay)
                '                .bTrig.Signal_ = R.SP.bTrigger.Text
            End If
            If .bTrigger.qHit Then
                .bTrigger.qHiti = 255
            Else
                .bTrigger.qHiti = 0
            End If
            If R.SP.bEntry.Idx = 0 Then
                .bEntry.qHit = True
                .bEntry.Text = R.SP.bEntry.Text & ":" & .bEntry.Text & "@" & Format$(0.0, "000.000")
            Else
                .bEntry.qHit = QBExeEntry(Counters.this_Day)
                .bEntry.Text = R.SP.bEntry.Text & ":" & .bEntry.Text & "@" & Format$(.bEntry.executePrice, "000.000")
            End If
            If .bEntry.qHit Then
                .bEntry.qHiti = 255
            Else
                .bEntry.qHiti = 0
            End If
            .bEntry.qHit = .bTrigger.qHit And .bEntry.qHit And .bDOW.qHit And .bTD1.qHit And .bTD2.qHit And .bZscoreMode.qHit
            If .bEntry.qHit Then
                .bEntry.qHiti = 255
            Else
                .bEntry.qHiti = 0
            End If
            .bEntry.Text = .bEntry.Text
            '           If .bentry.qHit Then Stop
            SetLastDayTrade = .bEntry.qHit
        End With
    End Function
    Public Function SetLastDayTradenozzz(ByRef R As Results, ByRef TS() As Trades, ByRef thisDay As Integer) As Boolean
        With lastDayTrade
            .bDOW.Idx = R.SP.bDOW.Idx
            R.SP.bDOW.Text1 = RMain.SP.bDOW.Text & ":" & Format$(op(thisDay), "000.000") & "!" &
              Format$(cl(thisDay), "000.000")
            lastDayTrade.bDOW.Text = R.SP.bDOW.Text1 & "::" & dtStr1(thisDay) & ":" & Format(thisDay, "00000")
            lastDayTrade.bDOW.qHit = QExeBDOW(thisDay)
            If .bDOW.qHit Then
                lastDayTrade.bDOW.qHiti = 255
                QFE1.qbDOW.Text = "qbDOW=T"
            Else
                .bDOW.qHiti = 0
                QFE1.qbDOW.Text = "qbDOW=F"
            End If
            .bMonth.Idx = R.SP.bMonth.Idx
            lastDayTrade.bMonth.Text = R.SP.bMonth.Text1 & "::" & dtStr1(thisDay) & ":" & Format(thisDay, "00000")
            lastDayTrade.bMonth.qHit = QExeMonth(thisDay)
            If lastDayTrade.bMonth.qHit Then
                lastDayTrade.bMonth.qHiti = 255
                QFE1.qBMonth.Text = "qbMon=T"
            Else
                .bMonth.qHiti = 0
                QFE1.qBMonth.Text = "qbMon=F"
            End If
            'If Not lastDayTrade.bTD1.qHit Then Stop
            lastDayTrade.bTD1.Idx = R.SP.bTD1.Idx
            lastDayTrade.bTD1.qHit = QTD_(lastDayTrade.bTD1, thisDay, lastDayTrade.bTD1.db)
            lastDayTrade.bTD1.Text = QTD_Text1(lastDayTrade.bTD1, thisDay)
            lastDayTrade.bTD2.Idx = R.SP.bTD2.Idx
            lastDayTrade.bTD2.qHit = QTD_(lastDayTrade.bTD2, thisDay, lastDayTrade.bTD2.db)
            lastDayTrade.bTD2.Text = QTD_Text1(lastDayTrade.bTD2, thisDay)
            R.SS.Main.bSignal.Pcntg = R.SS.Main.bSignal.Hits / R.SS.Main.Days
            .bTD1.Idx = R.SP.bTD1.Idx
            .bTD1.dbIdx = R.SP.bTD1.dbIdx
            .bTD1.db = R.SP.bTD1.dbText
            .bTD1.qHiti = 255
            If .bTD1.Idx > 0 Then
                .bTD1.qHit = QTD_(.bTD1, thisDay, .bTD1.db)
                R.SP.bTD1.Text1 = QTD_Text1(.bTD1, thisDay)
                If .bTD1.qHit Then
                    QFE1.qbTD1.Text = "qbTD1=T"
                Else
                    .bTD1.qHiti = 0
                    QFE1.qbTD1.Text = "qbTD1=F"
                End If
                '                   Strings.Left(R.SP.bTD2.Text, 12) & ":" &     Format$(.bTD2.actualPrice, "000.000") & "on" & .bTD2.Text & "&" & Format$(.bTD2.actualPrice, "000.000") & "on" & .bTD2.Text
            Else
                R.SP.bTD1.Text1 = "None"
                .bTD1.qHit = True
            End If
            .bTD2.Idx = R.SP.bTD2.Idx
            .bTD2.dbIdx = R.SP.bTD2.dbIdx
            .bTD2.db = R.SP.bTD2.dbText
            .bTD2.qHiti = 255
            If .bTD2.Idx > 0 Then
                .bTD2.qHit = QTD_(.bTD2, thisDay, .bTD2.db)
                R.SP.bTD2.Text1 = QTD_Text1(.bTD2, thisDay)
                If .bTD2.qHit Then
                    QFE1.qbTD2.Text = "qbTD2=T"
                Else
                    .bTD2.qHiti = 0
                    QFE1.qbTD2.Text = "qbTD2=F"
                End If
                '                   Strings.Left(R.SP.bTD2.Text, 12) & ":" &     Format$(.bTD2.actualPrice, "000.000") & "on" & .bTD2.Text & "&" & Format$(.bTD2.actualPrice, "000.000") & "on" & .bTD2.Text
            Else
                R.SP.bTD2.Text1 = "None"
                .bTD2.qHit = True
            End If
            R.SP.bTD0.Text = R.SP.bTD1.Text & R.SP.bTD2.Text
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
            lastDayTrade.bInDay.qHit = QbuyInDay(thisDay)
            lastDayTrade.bInDay.Text = R.SP.bInday.Text
            lastDayTrade.bInDay.Signal = R.SP.bInday.Text & QbuyInDaySignal(thisDay)
            If lastDayTrade.bInDay.qHit Then
                lastDayTrade.bInDay.qHiti = 255
                QFE1.qbInsideDay.Text = "qbInsideDy=T"
            Else
                lastDayTrade.bInDay.qHiti = 0
                QFE1.qbInsideDay.Text = "qbInsideDay=F"
            End If
            '            .bZscore.qHit = QZ_(R, R.SS.Trades, 0)
            .bZscoreMode.qHit = qHitPatternZ(R, RMainTS, R.SS.Main.Trades, 0)
            If .bZscoreMode.qHit Then
                .bZscoreMode.qHiti = 255
            Else
                .bZscoreMode.qHiti = 0
                QFE1.qbZScrMode.Text = "qbZScrM=F"
            End If
            .bZScoreValue.qHit = qHitValueZ(R, lastDayTrade)
            If .bZScoreValue.qHit Then
                .bZScoreValue.qHiti = 255
                QFE1.qbZScrValue.Text = "qbZScrV=T"
            Else
                .bZScoreValue.qHiti = 0
                QFE1.qbZScrValue.Text = "qbZScrV=F"
            End If
            'lastDayTrade.z.Text = R.SP.bZScore.Text & "ltpr=" & StrSign(lastTrade.Profit0) & "!" & _
            '               strSign(lastTrade.Profit1) & "!" & strSign(lastTrade.Profit2) & "!" & _
            '              strSign(lastTrade.Profit3) & "!"
            Select Case R.SP.bEntry.Idx
                Case 0
                    Stop
                Case 1
                    R.SP.bEntry.Text1 = R.SP.bEntry.Text & Format(op(thisDay), "000.000")
                Case 2
                    R.SP.bEntry.Text1 = R.SP.bEntry.Text & Format(cl(thisDay), "000.000")
                Case 3
                    R.SP.bEntry.Text1 = R.SP.bEntry.Text & Format(0.0, "000.000")
                Case 4
                    R.SP.bEntry.Text1 = R.SP.bEntry.Text & Format(0.0, "000.000")
                Case Else
                    Stop
            End Select
            lastDayTrade.bEntry.qHit = lastDayTrade.bDOW.qHit And lastDayTrade.bMonth.qHit And
                lastDayTrade.bInDay.qHit And
                 lastDayTrade.bTrigger.qHit And
                  lastDayTrade.bTD1.qHit And
                   lastDayTrade.bTD2.qHit And
                   lastDayTrade.bZscoreMode.qHit And lastDayTrade.bZScoreValue.qHit
            lastDayTrade.bEntry.Text = R.SP.bEntry.Text
            If lastDayTrade.bEntry.qHit Then
                lastDayTrade.bEntry.qHiti = 255
            Else
                lastDayTrade.bEntry.qHiti = 0
            End If
        End With
        SetLastDayTradenozzz = lastDayTrade.bEntry.qHit
    End Function
End Module

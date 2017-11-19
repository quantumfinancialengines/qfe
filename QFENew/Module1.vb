Module PutWrite

    Public Sub Put_Parameters0(ByRef R As Results)
        '       Dim OleDBC As New OleDbCommand
        '       Dim conn0 As New System.Data.OleDb.OleDbConnection
        With R
            .SS.Main.absQuantum = Math.Abs(.SS.Main.Q.V)
            .bsEntryExit =
                Format$(.SS.Main.Q.V, "0.000") & "bh" & Format$(.bhQ, "0.000") & "!" &
                Format$(.SS.Main.ROR1, "0.000") & "#" & Format$(.SS.Main.ROR2, "0.000") & "#" &
                Format$(.SS.Main.Correlation.V, "0.000") &
                "!" & .lSymbol & ":" &
                Strings.Left(.SP.bEntry.Text, 7) & ":" &
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
            '            OleDBC.Connection = conn0
            '           OleDBC.CommandText = "Insert Into Parameters00 VALUES ('" & Counters.totalIterations &
            '                "','" & .sSymbol &
            '               "','" & .lSymbol &
            '              "','" & .bsEntryExit &
            '            "','" & Now() & "')"
            '          OleDBC.ExecuteNonQuery()
        End With
    End Sub
    Public Sub Put_Parameters1(ByRef R As Results)
        '     Dim OleDBC As New OleDbCommand
        Dim conn0 As New System.Data.OleDb.OleDbConnection
        'With R
        '           OleDBC.Connection = conn0
        '          OleDBC.CommandText = "Insert Into Parameters01 VALUES ('" & Counters.totalIterations &
        '                "','" & .SP.bentry.qP &
        '               "','" & .SP.bentry.Text &
        '              "','" & .SS.Main.bentry.Hits &
        '             "','" & .SP.bMA1.qP &
        '            "','" & .SP.bMA1.Text &
        '           "','" & .SS.Main.bMA1.Hits &
        '          "','" & .SP.bMA2.qP &
        '         "','" & .SP.bMA2.Text &
        '        "','" & .SS.Main.bMA2.Hits &
        '       "','" & .SP.bMA3.qP &
        '      "','" & .SP.bMA3.Text &
        '     "','" & .SS.Main.bMA3.Hits &
        '    "','" & .SP.bMinDH.V &
        '   "','" & .SP.bMinDH.Text &
        '  "','" & .SP.sMinDH.V &
        ' "','" & "00" &
        '"','" & .SP.sMaxDH.V &
        '                "','" & .SP.sMaxDH.Text &
        '               "','" & .bsEntryExit &
        '              "','" & Now() & "')"
        '         OleDBC.ExecuteNonQuery()
        '    End With
    End Sub
    Public Sub Put_Parameters2(ByRef R As Results)
        '   Dim OleDBC As New OleDbCommand
        Dim conn0 As New System.Data.OleDb.OleDbConnection
        Static td1str As String, td2str As String
        'static td3str As String, td4str As String, td5str As String, tdsstr As String
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
            ' tdsstr = Strings.Left(.SP.sTD1.Text, 4) &
            'Format$(.SP.sTD1.dbV, "00") & "]"
            .SP.bTD1.qP = .SP.bTD1.Idx > 0
            '      OleDBC.Connection = conn0
            '     OleDBC.CommandText = "Insert Into Parameters02 VALUES ('" & Counters.totalIterations &
            '        "','" & .SS.Main.Days &
            '       "','" & .SS.Main.Trades &
            '      "','" & .SP.bTD1.qP &
            '     "','" & lastTrade.bTD1.qHiti &
            '    "','" & td1str &
            '   "','" & lastTrade.bTD1.Text & " " &
            '  "','" & .SS.Main.bTD1.Hits &
            ' "','" & .SS.Main.bTD1.Misses &
            '"','" & .SP.bTD2.qP &
            '               "','" & lastTrade.bTD2.qHiti &
            '              "','" & td2str &
            '             "','" & lastTrade.bTD2.Text & " " &
            '            "','" & .SS.Main.bTD2.Hits &
            '           "','" & .SP.bTD3.qP &
            '          "','" & lastTrade.bTD3.qHiti &
            '         "','" & td3str &
            '        "','" & lastTrade.bTD3.Text & " " &
            '       "','" & .SS.Main.bTD3.Hits &
            '      "','" & .SP.bTD4.qP &
            '     "','" & lastTrade.bTD4.qHiti &
            '    "','" & td4str &
            '   "','" & lastTrade.bTD4.Text & " " &
            '  "','" & .SS.Main.bTD4.Hits &
            ' "','" & .SP.bTD5.qP &
            '"','" & lastTrade.bTD5.qHiti &
            '                "','" & td5str &
            '               "','" & lastTrade.bTD5.Text & " " &
            '              "','" & .SS.Main.bTD5.Hits &
            '             "','" & .SP.sTD1.qP &
            '            "','" & .SP.sTD1.qHiti &
            '           "','" & tdsstr &
            '          "','" & .SP.sTD1.Text &
            '         "','" & .SS.Main.sTD1.Hits &
            '        "','" & Now() & "')"
            '   OleDBC.ExecuteNonQuery()
        End With
    End Sub
    Public Sub Put_Statistics00(ByRef R As Results)
        Static q As Boolean, bsignal As Single
        q = QBuySignal(Counters.end_day)
        ' With R
        '.sSymbol = Strings.Left(.sSymbol & "____", 4) & Strings.Left(R.SP.bDOW.Text, 2)
        'If .SS.Main.Losers.V = 0 Then
        '.SS.Main.Losers.avg = 0.0
        '.SS.Main.Losers.tot = 0.0
        'End If
        bsignal = 0.0
        '.SS.Main.Profits.Pcntg = .SS.Main.Profits.avg / 10
        '     OleDBC.Connection = conn0
        '    OleDBC.CommandText = "Insert Into Statistics00 VALUES ('" & Counters.totalIterations &
        '   "','" & Counters.SystemNumber &
        '  "','" & lastTrade.bentry.qHiti &
        ' "','" & lastTrade.bDOW.qHiti &
        '          "','" & lastTrade.bZscore.qHiti &
        '         "','" & lastTrade.bActual.qHiti &
        '        "','" & .sSymbol &
        '       "','" & .lSymbol &
        '      "','" & Counters.this_Date & " " &
        '     "','" & Strings.Left(.SP.bDOW.Text, 3) &
        '    "','" & Strings.Left(.SP.bentry.Text, 5) &
        '   "','" & Strings.Left(.SP.bSignal.Text, 9) & ":" & Format$(newTrade.bTrigger.triggerPrice, "000.000") &
        '  "','" & Strings.Left(.SP.sSignal.Text, 5) &
        ' "','" & .SS.Main.Q.avg &
        '"','" & .SS.Main.Q.avg &
        '           "','" & Math.Abs(.SS.Main.Q.avg) &
        '          "','" & .bhQ &
        '         "','" & .bhAveProfit &
        '        "','" & .SS.Main.Profits.tot &
        '       "','" & .SS.Main.Profits.avg &
        '      "','" & .Days &
        '     "','" & .SS.Main.bentry.Hits &
        '    "','" & .SS.Main.bSignal.Hits &
        '   "','" & .SS.Main.bDOW.Hits &
        '  "','" & .SS.Main.zScore.Hits &
        ' "','" & .posOutliers &
        ' "','" & .negOutliers &
        '        "','" & .SS.Main.Trades &
        '       "','" & .SS.Main.Winners.tot &
        '      "','" & .SS.Main.Losers.tot &
        '     "','" & .SS.Main.wPcntg &
        '    "','" & .SS.Main.Winners.tot &
        '   "','" & .SS.Main.Losers.tot &
        '  "','" & .SS.Main.Winners.avg &
        ' "','" & .SS.Main.Losers.avg &
        '"','" & .SS.Main.Profits.tot &
        '            "','" & .SS.Main.Profits.avg &
        '           "','" & .SS.Main.Profits.Pcntg &
        '          "','" & .SP.sMaxDH.V &
        '         "','" & .SS.Main.maxDH &
        '        "','" & .SS.Main.maxDH &
        '       "','" & .SS.Main.drawDown.tot &
        '      "','" & .SS.Main.drawDown.avg &
        '     "','" & .SS.Main.drawUp.tot &
        '    "','" & .SS.Main.drawUp.avg &
        '   "','" & .SS.Main.Profits.Peak &
        '  "','" & .SS.Main.Profits.Trough &
        ' "','" & .SS.Main.Correlation.V &
        '"','" & .SS.Main.COV.V &
        '           "','" & .SS.Main.ROR1 &
        '          "','" & .SS.Main.ROR2 &
        '         "','" & .SS.Main.ROR3 &
        '        "','" & Now() & "')"
        '       OleDBC.ExecuteNonQuery()
        '  End With
    End Sub
    Public Sub Put_Statistics01(ByRef R As Results, ByRef TS() As Trades)
        '        Dim OleDBC As New OleDbCommand
        '       Dim conn0 As New System.Data.OleDb.OleDbConnection
        'With R
        '            OleDBC.Connection = conn0
        '           OleDBC.CommandText = "Insert Into Statistics01 VALUES ('" & Counters.totalIterations &
        '          "','" & .SS.Main.Runs &
        '         "','" & .SS.Main.expRuns &
        '        "','" & .SS.Main.zScore.V &
        '       "','" & .SS.Main.ZScoreM_ &
        '      "','" & R.lastTrade.TradeNo &
        '     "','" & R.lastTrade.Profit0 &
        '    "','" & R.lastTrade.bDayNo &
        '   "','" & R.lastTrade.bDate &
        '  "','" & R.lastTrade.sDate &
        ' "','" & R.lastTrade.bentry.executePrice &
        '"','" & R.lastTrade.sExecute.executePrice &
        '            "','" & Now() & "')"
        '           OleDBC.ExecuteNonQuery()
        '      End With

    End Sub
    Public Sub Put_Trades0(ByRef R As Results, ByRef TS() As Trades)
        Static xx As Integer, plotQ As Single
        Static str1 As String
        '       Dim OleDBC As New OleDbCommand
        '       Dim conn0 As New System.Data.OleDb.OleDbConnection
        '      strConnString0 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Dad\Documents\dbactivity11.mdb"
        '                conn0.ConnectionString = strConnString0
        '      strConnString0 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Dad\Documents\dbactivity11.mdb"
        '        conn0.ConnectionString = strConnString0
        ''                conn0.Open()
        '        dbTextBox0.Text = strConnString0
        '  OleDBC.Connection = conn0
        lastTrade = New Trades
        For xx = 1 To R.SS.Main.Trades
            str1 = R.lSymbol & Format$(Counters.SystemNumber, "0000")
            With TS(xx)
                '               If Math.Abs(.Quantum) > 1.0 Then .Quantum = 0.0
                '                If Math.Abs(.totQuantum) > 1.0 Then .totQuantum = 0.0
                If xx > 15 Then
                    plotQ = .totQuantum
                Else
                    plotQ = .totQuantum * (xx / 15)
                End If
                .Descr = StrTr(R, TS(xx))
                '            OleDBC.CommandText = "Insert Into Trades00 VALUES ('" & Counters.totalIterations &
                '           "','" & Counters.SystemNumber &
                '          "','" & .Descr &
                '         "','" & Strings.Left(R.lSymbol, 170) &
                '        "','" & .bentry.Date_ &
                '       "','" & .TradeNo &
                '      "','" & .oldTradeNo &
                '     "','" & .totProfit &
                '    "','" & .totQuantum &
                '   "','" & Format$(plotQ, "0.000") &
                '  "','" & .DH &
                ' "','" & .totDH &
                '"','" & Strings.Left(str1, 170) &
                '              "','" & Now() & "')"
                '             OleDBC.ExecuteNonQuery()
            End With
            lastTrade = TS(xx)
        Next xx
    End Sub
    Public Sub Put_Trades1(ByRef R As Results, ByRef tsa() As Trades)
        Static totQ As Single, totPr As Single, totDH As Single, xxx As Integer
        Dim str1 As String, str2 As String, xx As Integer
        '        Dim OleDBC As New OleDbCommand
        '        Dim conn0 As New System.Data.OleDb.OleDbConnection
        '        OleDBC.Connection = conn0
        totQ = 0.0
        totPr = 0.0
        totDH = 0.0
        xx = Strings.InStr(R.sSymbol, " ")
        If xx <> 0 Then
            R.sSymbol = Strings.Left(R.sSymbol, xx) & "_____"
        Else
            R.sSymbol = R.sSymbol & "_____"
        End If
        str1 = R.sSymbol & ":" & Format$(Counters.SystemNumber, "00000")
        str1 = str1 & ":" & R.SP.bDOW.Text & ":" & R.SP.bDOW.Text & ":" & R.SP.bEntry.Text & ":" &
         Format(R.SP.bMaxDH.V, "00") & ":" &
          R.SP.sEntry.Text & ":"
        str2 = R.sSymbol & ":" & Format$(Counters.SystemNumber, "0000") & ":" &
                R.SP.bDOW.Text & R.SP.bTD1.Text & ":" & R.SP.bTD2.Text & ":" &
                 R.SP.sMaxDH.Text & "=" & Format$(R.SS.Main.Q.avg, "0.00")
        '        R.qPosLTr = TSa.(R.SS.Main.Trades). > 0.0
        R.sSymbol = Strings.Left(R.sSymbol, 5) & str1
        If R.qPosLTr Then
            xxx = 1
        Else
            xxx = 0
        End If
        If Strings.InStr(R.sSymbol, " ") = 0 Then
            R.sSymbol = Strings.Left(R.sSymbol & "_____", 5) & Format(Counters.SystemNumber, "0000")
        Else
            R.sSymbol = Strings.Left(Strings.Left(R.sSymbol, Strings.InStr(R.sSymbol, " ") - 1) & "_____", 5) & Format(Counters.SystemNumber, "0000")
        End If
        R.sSymbol = R.sSymbol & ":" & Strings.Left(R.SP.bDOW.Text, 4)
        For ix = 1 To R.SS.Main.Trades - 2
            Counters.Incr_ = Counters.Incr_ + 1
            Call insertTrade(R, tsa(ix))
        Next ix
        Application.DoEvents()
    End Sub
    Public Sub insertTrade(ByRef R As Results, ByRef TTrades As Trades)
        With TTrades
            Counters.Incr_ = Counters.Incr_ + 1
            Counters.totalTrades = Counters.totalTrades + 1
            If TTrades.sTP.qHit Then
                TTrades.sTP.qHiti = 1
            Else
                TTrades.sTP.qHiti = 0
            End If
            QFE1.oleDatabaseTrades.CommandText =
    "Insert Into Trades VALUES ('" & Counters.Incr_ &
     "','" & Counters.totalIterations &
     "','" & R.Symbol &
     "','" & R.sSymbol &
     "','" & R.lSymbol &
     "','" & R.signal &
     "','" & Counters.SystemNumber &
     "','" & R.SP.bDOW.Text &
      "','" & Counters.totalTrades &
      "','" & TTrades.TradeNo &
       "','" & TTrades.oldTradeNo &
        "','" & TTrades.Quantum &
         "','" & TTrades.totQuantum &
          "','" & TTrades.Profit0 &
           "','" & TTrades.totProfit &
            "','" & TTrades.bPrice &
             "','" & TTrades.sPrice &
              "','" & TTrades.bAmt &
               "','" & TTrades.sAmt &
               "','" & TTrades.Shares &
               "','" & TTrades.bDate &
               "','" & TTrades.sDate &
               "','" & TTrades.bDayNo &
               "','" & TTrades.sDayNo &
               "','" & Strings.Left(TTrades.bEntry.Text, 8) &
               "','" & Strings.Left(TTrades.sEntry.Text, 8) &
               "','" & TTrades.DH &
               "','" & TTrades.totDH &
               "','" & Format$(TTrades.maxDH, "00") &
               "','" & TTrades.bZscoreMode.Text &
               "','" & TTrades.bZScoreValue.Text &
               "','" & Format(TTrades.sTP.qHiti) &
               "','" & Format(R.SP.TP.V, "0.00") &
               "','" & Format(TTrades.bPrice, "0.000") &
               "','" & Format(TTrades.sPrice, "0.000") &
               "','" & Format(TTrades.sTP.actualPrice, "0.000") &
               "','" & Format(TTrades.sTP.triggerPrice, "0.000") &
               "','" & Format(.TPHits, "0000") &
               "','" & Format(.TPMisses, "0000") &
               "','" & TTrades.bInDay.Text &
               "','" & RMain.SP.bTD1.Text &
               "','" & TTrades.bTD1.Text &
               "','" & .sEntry.qHiti &
               "','" & TTrades.bOpen &
               "','" & TTrades.bHigh &
               "','" & TTrades.bLow &
               "','" & TTrades.bClose &
               "','" & TTrades.sOpen &
               "','" & TTrades.sHigh &
               "','" & TTrades.sLow &
               "','" & TTrades.sClose &
               "','" & TTrades.Profit0 &
               "','" & TTrades.Profit1 &
               "','" & TTrades.Profit2 &
               "','" & TTrades.Profit3 &
               "','" & TTrades.Profit4 &
    "','" & Now() & "')"
            QFE1.oleDatabaseTrades.ExecuteNonQuery()
        End With
    End Sub
    Public Sub Put_Trades1a(ByRef R As Results, ByRef TS() As Trades)
        Static totQ As Single, totPr As Single, totDH As Single, xx As Integer, xxx As Integer
        Static str1 As String, stra As String, strb As String, strc As String
        Static str As String
        Static totProfit As Single ', avgProfit As Single, avgDH As Single
        xx = 0
        '        Dim OleDBC As String 'New OleDbCommand
        '        Dim conn0 As New System.Data.OleDb.OleDbConnection
        '        OleDBC.Connection = conn0
        '        conn0.ConnectionString = strConnString0
        '       conn0.Open()
        str = R.lSymbol
        '        Me.lSymbol.Text = str
        str = "#2000/01/02#"
        TS(1).sliderDate = FormatDateTime(str, DateFormat.ShortDate)
        Counters.totalTrades = Counters.totalTrades + 1
        '       dbTextBox0.Text = strConnString1
        '      OleDBC.Connection = conn1
        totQ = 0.0
        totPr = 0.0
        totDH = 0.0
        str1 = R.lSymbol & ":" &
            Format$(Counters.SystemNumber, "00000") & ":" &
            Format$(Counters.totalIterations, "00000") & ":" &
            Format$(0, "0000")
        '       str2 = R.sSymbol & ":" & Format$(Counters.SystemNumber, "0000") & ":" & R.SP.bDOW.Text & _
        '           ":" & R.SP.bZScore.Text & ":" & _
        '           R.SP.bTD1.Text & ":" & R.SP.bTD2.Text & ":" & _
        '           R.SP.sMaxDH.Text & "=" & strSign(Format$(R.SS.Main.Q.avg, "0.00"))
        R.qPosLTr = TS(R.SS.Main.Trades).Profit0 > 0.0
        If R.qPosLTr Then
            xxx = 1
        Else
            xxx = 0
        End If
        If Counters.qBuyandHold.qHit Then
            Counters.qBuyandHold.qHiti = 255
        Else
            Counters.qBuyandHold.qHiti = 0
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
            stra = "#" & Strings.Left(.bDate, 4) & "/"
            strb = Strings.Mid(.bDate, 5, 2) & "/"
            strc = Strings.Mid(.bDate, 7, 2) & "#"
            TS(1).sliderDate = stra & strb & strc
            .timeDate = Strings.Mid(.bDate, 5, 2) & "-" & Strings.Mid(.bDate, 7, 2) & "-" & Strings.Left(.bDate, 4)
            '            OleDBC.CommandText = "Insert Into Trades00 VALUES ('" & Counters.totalTrades &
            '           "','" & Counters.totalIterations &
            '          "','" & Counters.SystemNumber &
            '         "','" & R.sSymbol &
            '        "','" & R.lSymbol &
            '       "','" & Counters.last_Date &
            '      "','" & .bDOW.Text & ":" &
            '     "','" & Counters.qBuyandHold.qHiti &
            '    "','" & Counters.qBuyandHold.Text &
            '   "','" & TS(1).bDate &
            ''  "','" & TS(1).sliderDate &
            '"','" & TS(1).sDate &
            '           "'','" & Counters.first_Date &
            '          "'','" & R.SP.bZScoreMode.Text &
            '         "','" & lastDayTrade.bentry.qHiti &
            '        "','" & 0 &
            '       "','" & 0 &
            '      "','" & 0 &
            '     "','" & 0 &
            '    "','" & 0.0 &
            '   "','" & 0.0 &
            '  "','" & 0.0 &
            ' "','" & 0.0 &
            '"','" & 0.0 &
            '            "','" & 0.0 &
            '           "','" & 0.0 &
            '          "','" & 0.0 &
            '         "','" & 0 &
            '        "','" & 0.0 &
            '       "','" & 0.0 &
            '      "','" & "000000" &
            '     "','" & "000000" &
            '    "','" & 0.0 &
            '   "','" & 0.0 &
            '  "','" & 0.0 &
            ' "','" & 0.0 &
            '"','" & Now() & "')"
            '            OleDBC.ExecuteNonQuery()
        End With
        totProfit = TS(1).Profit0 + TS(2).Profit0 + TS(3).Profit0
        totDH = TS(1).DH + TS(2).DH + TS(3).DH
        If totDH = 0 Then totDH = 3
        '        For xx = 4 To R.SS.Main.Trades
        '       With TS(xx)
        '      If xx > 10 Then
        '     .zScore00 = CalcZScore(R, TS, 1, xx)
        '    Else
        '   .zScore00 = 0.0
        '  End If
        ' If xx > 10 Then
        '.zScore10 = CalcZScore(R, TS, xx - 10, xx)
        '        Else
        '       .zScore10 = 0.0
        '      End If
        '     If xx > 20 Then
        '    .zScore20 = CalcZScore(R, TS, xx - 20, xx)
        '   Else
        '  .zScore20 = 0.0
        ' End If
        'If xx > 30 Then
        '       .zScore30 = CalcZScore(R, TS, xx - 30, xx)
        '      Else
        '     .zScore30 = 0.0
        '    End If
        '   TS(xx).absTotQuantum = Math.Abs(TS(xx).totQuantum)
        '  str1 = R.lSymbol & ":" &
        ' Format$(Counters.SystemNumber, "00000") & ":" &
        'Format$(Counters.totalIterations, "00000") & ":" &
        'Format$(.TradeNo, "0000")
        'Counters.Incr_ = Counters.Incr_ + 1
        'Counters.totalTrades = Counters.totalTrades + 1
        'stra = "#" & Strings.Left(TS(xx).bDate, 4) & "/"
        'strb = Strings.Mid(TS(xx).bDate, 5, 2) & "/"
        'strc = Strings.Right(TS(xx).bDate, 2) & "#"
        'TS(xx).sliderDate = stra & strb & strc
        'totProfit = totProfit + .Profit0
        '                TS(xx).totProfit = totProfit
        'avgProfit = totProfit / xx
        'If TS(xx).DH = 0 Then TS(xx).DH = 1
        '               totDH = totDH + .DH
        '                TS(xx).totDH = totDH
        'avgDH = totDH / xx
        '               .avgProfit = avgProfit
        '              .avgDH = avgDH
        '             .totQuantum = .avgProfit / .avgDH / 10
        'If .bDate <> "" Then
        '              If .sDate = "" Then .sDate = dtStr2(Counters.end_day)
        '              If .bEntry.Text = "" Then .bEntry.Text = "--"
        '              If .sEntry.Text = "" Then .sEntry.Text = "--"
        '                   If .TradeNo = 0 Then Stop
        '                .timeDate = Strings.Mid(.bDate, 5, 2) & "-" & Strings.Mid(.bDate, 7, 2) & "-" & Strings.Left(.bDate, 4)
        '                    OleDBC.CommandText = "Insert Into Trades00 VALUES ('" & Counters.totalTrades &
        '                   "','" & Counters.totalIterations &
        '                  "','" & Counters.SystemNumber &
        '                 "','" & R.sSymbol &
        '                "','" & R.lSymbol &
        '               "','" & Counters.last_Date &
        '              "','" & .bDOW.Text & ":" &
        '     "','" & Counters.qBuyandHold.qHiti &
        '    "','" & Counters.qBuyandHold.Text &
        '           "','" & .bDate &
        '          "','" & .sliderDate &
        '         "','" & .sDate &
        '        "','" & .sDOW.Text &
        '       "','" & R.SP.bZScoreMode.Text &
        '      "','" & lastDayTrade.bentry.qHiti &
        '     "','" & .TradeNo &
        '    "','" & .oldTradeNo &
        '   "','" & .Profit0 &
        '  "','" & .totProfit &
        ' "','" & .avgProfit &
        '"','" & .totQuantum &
        '                    "','" & .absTotQuantum &
        '                   "','" & .DH &
        '                  "','" & .totDH &
        '                 "','" & .avgDH &
        '                "','" & .maxDH &
        '               "','" & .bAmt &
        '              "','" & .sAmt &
        '             "','" & .bPrice &
        '            "','" & .sPrice &
        '           "','" & R.SP.bentry.Text & ":" & Format$(op(.bDayNo), "000.000") & "!" & Format$(cl(.bDayNo), "000.000") &
        '          "','" & R.SP.sExecute.Text & ":" & Format$(op(.sDayNo), "000.000") & "!" & Format$(cl(.sDayNo), "000.000") &
        '         "','" & .zScore00 &
        '        "','" & .zScore10 &
        '       "','" & .zScore20 &
        '      "','" & .zScore30 &
        '     "','" & Now() & "')"
        '    OleDBC.ExecuteNonQuery()
        'End If
        'End With
        'Next xx
        '        R.TS(R.SS.Main.Trades + 1) = R.TS(R.SS.Main.Trades)
        lastDayTrade.bDate = Counters.last_Date
        TS(R.SS.Main.Trades + 1) = TS(R.SS.Main.Trades) 'lastDayTrade
        With TS(R.SS.Main.Trades + 1)
            .bDate = "20200101"
            .TradeNo = R.SS.Main.Trades + 1
            Counters.Incr_ = Counters.Incr_ + 1
            Counters.totalTrades = Counters.totalTrades + 1
            stra = "#" & Strings.Left(.bDate, 4) & "/"
            strb = Strings.Mid(.bDate, 5, 2) & "/"
            strc = Strings.Right(.bDate, 2) & "#"
            .sliderDate = stra & strb & strc
            '            OleDBC.CommandText = "Insert Into Trades00 VALUES ('" & Counters.totalTrades &
            '           "','" & Counters.totalIterations &
            '          "','" & Counters.SystemNumber &
            '         "','" & R.sSymbol &
            '        "','" & R.lSymbol &
            '       "','" & Counters.last_Date &
            '      "','" & .bDOW.Text & ":" &
            '     "','" & Counters.qBuyandHold.qHiti &
            '    "','" & Counters.qBuyandHold.Text &
            '   "','" & .bDate &
            '  "','" & .sliderDate &
            ' "','" & .bDate &
            '"','" & .sDOW.Text &
            '            "','" & R.SP.bZScoreMode.Text &
            '           "','" & lastDayTrade.bentry.qHiti &
            '          "','" & .TradeNo &
            '         "','" & .oldTradeNo &
            '        "','" & .Profit0 &
            '       "','" & .totProfit &
            '      "','" & .avgProfit &
            '     "','" & .totQuantum &
            '    "','" & .absTotQuantum &
            '   "','" & .DH &
            '  "','" & .totDH &
            ' "','" & .avgDH &
            '"','" & .maxDH &
            '            "','" & .bAmt &
            '           "','" & .sAmt &
            '          "','" & .bPrice &
            '         "','" & .sPrice &
            '        "','" & R.SP.bentry.Text & ":" & Format$(op(.bDayNo), "000.000") & "!" & Format$(cl(.bDayNo), "000.000") &
            '       "','" & R.SP.sExecute.Text & ":" & Format$(op(.sDayNo), "000.000") & "!" & Format$(cl(.sDayNo), "000.000") &
            '      "','" & .zScore00 &
            '     "','" & .zScore10 &
            '    "','" & .zScore20 &
            '   "','" & .zScore30 &
            '  "','" & Now() & "')"
            ' '       OleDBC.ExecuteNonQuery()
        End With
        TS(R.SS.Main.Trades + 2) = TS(R.SS.Main.Trades)
        With TS(R.SS.Main.Trades + 2)
            .bDate = "20210101"
            .TradeNo = R.SS.Main.Trades + 2
            Counters.Incr_ = Counters.Incr_ + 1
            Counters.totalTrades = Counters.totalTrades + 1
            stra = "#" & Strings.Left(.bDate, 4) & "/"
            strb = Strings.Mid(.bDate, 5, 2) & "/"
            strc = Strings.Right(.bDate, 2) & "#"
            .sliderDate = stra & strb & strc
            'OleDBC.CommandText = "Insert Into Trades00 VALUES ('" & Counters.totalTrades &
            ' "','" & Counters.totalIterations &
            '  "','" & Counters.SystemNumber &
            '   "','" & R.sSymbol &
            '    "','" & R.lSymbol &
            '     "','" & Counters.last_Date &
            '      "','" & .bDOW.Text & ":" &
            '       "','" & Counters.qBuyandHold.qHiti &
            '        "','" & Counters.qBuyandHold.Text &
            '         "','" & .bDate &
            '          "','" & .sliderDate &
            '           "','" & .bDate &
            '            "','" & .sDOW.Text &
            '"','" & R.SP.bZScoreMode.Text &
            ' "','" & lastDayTrade.bentry.qHiti &
            '  "','" & .TradeNo &
            '   "','" & .oldTradeNo &
            '    "','" & .Profit0 &
            '     "','" & .totProfit &
            '      "','" & .avgProfit &
            '       "','" & .totQuantum &
            '        "','" & .absTotQuantum &
            '         "','" & .DH &
            '          "','" & .totDH &
            '           "','" & .avgDH &
            '            "','" & .maxDH &
            ' "','" & .bAmt &
            '  "','" & .sAmt &
            '   "','" & .bPrice &
            '    "','" & .sPrice &
            '     "','" & R.SP.bentry.Text & ":" & Format$(op(.bDayNo), "000.000") & "!" & Format$(cl(.bDayNo), "000.000") &
            '      "','" & R.SP.sExecute.Text & ":" & Format$(op(.sDayNo), "000.000") & "!" & Format$(cl(.sDayNo), "000.000") &
            '       "','" & .zScore00 &
            '        "','" & .zScore10 &
            '         "','" & .zScore20 &
            '          "','" & .zScore30 &
            '           "','" & Now() & "')"
            '            OleDBC.ExecuteNonQuery()
        End With
        Application.DoEvents()
    End Sub
    Public Sub Put_ldTrades(ByRef R As Results, ByRef TS() As Trades)
        Static totQ As Single, totPr As Single, totDH As Single
        '      Static xx As Integer
        '       Static str1 As String
        'Dim OleDBC As New OleDbCommand
        Dim conn0 As New System.Data.OleDb.OleDbConnection
        '        strConnString0 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Dad\Documents\dbactivity11.mdb"
        '        strConnString0 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=Z:\QFE_DB\QFE_StockData.mdb"
        'conn0.ConnectionString = strConnString0
        conn0.Open()
        '       dbTextBox0.Text = strConnString0
        '       OleDBC.Connection = conn0
        totQ = 0.0
        totPr = 0.0
        totDH = 0.0
        If R.SS.Main.Trades > 0 Then
            '        For xx = R.SS.Main.Trades To R.SS.Main.Trades
            '       str1 = R.lSymbol & Format$(Counters.SystemNumber, "0000")
            '      With TS(xx)
            '     '                totDH = totDH + .DH
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
            '                    OleDBC.CommandText = "Insert Into ldTrades VALUES ('" & Counters.totalIterations &
            '                   "','" & .TradeNo &
            '                  "','" & .pcntgProfit &
            '                 "','" & .Quantum &
            '                "','" & .totQuantum &
            '               "','" & .Profit0 &
            '              "','" & .totProfit &
            '             "','" & .bPrice &
            '            "','" & .sPrice &
            '           "','" & .bAmt &
            '          "','" & .sAmt &
            '         "','" & .Shares &
            '        "','" & .bDate &
            '       "','" & .sDate &
            '      "','" & .bDayNo &
            '     "','" & .sDayNo &
            '    "','" & Counters.end_day &
            '   "','" & Counters.end_Date &
            '  "','" & .DH &
            ' "','" & .totDH &
            '"','" & .maxDH &
            '                    "','" & Now() & "')"
            '                   OleDBC.ExecuteNonQuery()
            '              End With
            '         Next xx
        End If
    End Sub
    Public Sub Put_Trades2(ByRef R As Results, ByRef TS() As Trades)
        Static xx As Integer, xxx As Long
        '       Dim OleDBC As New OleDbCommand
        '       Dim conn0 As New System.Data.OleDb.OleDbConnection
        '        OleDBC.Connection = conn0
        '        For xx = 1 To R.SS.Main.Trades
        '       With TS(xx)
        '      If .bZscore.qHit Then
        '     .bZscore.qHiti = 255
        '    Else
        '   .bZscore.qHiti = 0
        '  End If
        ' If .bentry.qHit Then
        '.bentry.qHiti = 255
        'Else
        '.bentry.qHiti = 0
        'End If
        'If .BuyOnOpen.qHit Then
        '.BuyOnOpen.qHiti = 255
        'Else
        '.BuyOnOpen.qHiti = 0
        'End If
        'If .BuyOnClose.qHit Then
        '.BuyOnClose.qHiti = 255
        'Else
        '.BuyOnClose.qHiti = 0
        'End If
        'If .BuyOnNextOpen.qHit Then
        '.BuyOnNextOpen.qHiti = 255
        'Else
        '.BuyOnNextOpen.qHiti = 0
        'End If
        'If .SellOnOpen.qHit Then
        '.SellOnOpen.qHiti = 255
        'Else
        '.SellOnOpen.qHiti = 0
        'End If
        'If .SellOnClose.qHit Then
        '.SellOnClose.qHiti = 255
        'Else
        '.SellOnClose.qHiti = 0
        'End If
        If xx = -1 Then
            '                    OleDBC.CommandText = "Insert Into Trades02 VALUES ('" & Counters.totalIterations &
            '                   "','" & R.lSymbol &
            '                  "','" & .TradeNo &
            '                 "','" & op(.tradeDayNo) &
            '                "','" & hi(.tradeDayNo) &
            '               "','" & lo(.tradeDayNo) &
            '              "','" & cl(.tradeDayNo) &
            '             "','" & op(.tradeDayNo - 1) &
            '            "','" & hi(.tradeDayNo - 1) &
            '           "','" & lo(.tradeDayNo - 1) &
            '          "','" & cl(.tradeDayNo - 1) &
            '         "','" & op(.tradeDayNo - 2) &
            '        "','" & hi(.tradeDayNo - 2) &
            '       "','" & lo(.tradeDayNo - 2) &
            '      "','" & cl(.tradeDayNo - 2) &
            '     "','" & .bDate &
            '    "','" & .bentry.Text &
            '   "','" & .sExecute.Text &
            '  "','" & .BuyOnOpen.qHiti &
            ' "','" & .BuyOnClose.qHiti &
            '"','" & .BuyOnNextOpen.qHiti &
            '                    "','" & .BuyOnNextOpen.qHiti &
            '                   "','" & .BuyOnNextClose.qHiti &
            '                  "','" & .BuyOnNextClose.qHiti &
            '                 "','" & .SellOnOpen.qHiti &
            '                "','" & .SellOnClose.qHiti &
            '               "','" & .SellOnNextOpen.qHiti &
            '              "','" & .SellOnExpired.qHiti &
            '             "','" & .bentry.executePrice &
            '            "','" & .bSignal.signalPrice &
            '           "','" & .bSignal.actualPrice &
            '          "','" & .bSignal.executePrice &
            '         "','" & .bSignal.triggerPrice &
            '        "','" & .bSignal.signalPrice &
            '       "','" & .sPrice &
            '      "','" & 0.0 &
            '     "','" & 0.0 &
            '    "','" & .Peak &
            '   "','" & .Trough &
            '  "','" & .Highest &
            ' "','" & .Lowest &
            '"','" & .DrawDn &
            '                   "','" & .maxDrawDn &
            '                  "','" & .maxDrawDnQ &
            '                 "','" & .Runs &
            '                "','" & .ConsW &
            '               "','" & .ConsL &
            '              "','" & .avgConsW &
            '             "','" & .avgConsL &
            '            "','" & .maxConsW &
            '           "','" & .maxConsL &
            '          "','" & Now() & "')"
            '         OleDBC.ExecuteNonQuery()
            '        .totTradeNum = .totTradeNum + 1
            '            .TradeNo = .TradeNo + 1
        End If
        xxx = xxx + 1
        '                OleDBC.CommandText = "Insert Into Trades02 VALUES ('" & xxx & "','" & Counters.totalIterations &
        '               "','" & R.lSymbol &
        '              "','" & .TradeNo &
        '             "','" & .tradeDayNo &
        '            "','" & .bDOW.Text &
        '           "','" & .bentry.Text &
        '          "','" & .bentry.Date_ &
        '         "','" & .bentry.Date_ &
        '        "','" & op(.bDayNo) &
        '       "','" & hi(.bDayNo) &
        '      "','" & lo(.bDayNo) &
        '     "','" & cl(.bDayNo) &
        '    "','" & op(.bDayNo - 1) &
        '   "','" & hi(.bDayNo - 1) &
        '  "','" & lo(.bDayNo - 1) &
        ' "','" & cl(.bDayNo - 1) &
        '"','" & op(.bDayNo - 2) &
        '                "','" & hi(.bDayNo - 2) &
        '               "','" & lo(.bDayNo - 2) &
        '              "','" & cl(.bDayNo - 2) &
        '             "','" & .bSignal.Text &
        '            "','" & .sSignal.Text &
        '           "','" & .bZscore.qHiti &
        '          "','" & .bentry.qHiti &
        '         "','" & .BuyOnOpen.qHiti &
        '        "','" & .BuyOnClose.qHiti &
        '       "','" & .BuyOnNextOpen.qHiti &
        '      "','" & .BuyOnNextClose.qHiti &
        '     "','" & .BuyOnNextClose.qHiti &
        '    "','" & .BuyOnNextClose.qHiti &
        '   "','" & .SellOnOpen.qHiti &
        '  "','" & .SellOnClose.qHiti &
        ' "','" & .SellOnNextOpen.qHiti &
        '"','" & .SellOnExpired.qHiti &
        '               "','" & .bSignal.signalPrice &
        '              "','" & .bTrigger.triggerPrice &
        '             "','" & .bActual.actualPrice &
        '            "','" & .bentry.executePrice &
        '           "','" & .bTrigger.triggerPrice &
        '          "','" & .bActual.actualPrice &
        '         "','" & .sPrice &
        '        "','" & .Profit0 &
        '       "','" & .totProfit &
        '      "','" & .Peak &
        '     "','" & .Trough &
        '    "','" & .Highest &
        '   "','" & .Lowest &
        '  "','" & .DrawDn &
        ' "','" & .maxDrawDn &
        '"','" & .maxDrawDnQ &
        '               "','" & .Runs &
        '              "','" & .ConsW &
        '             "','" & .ConsL &
        '            "','" & .avgConsW &
        '           "','" & .avgConsL &
        '          "','" & .maxConsW &
        '         "','" & .maxConsL &
        '        "','" & Now() & "')"
        '       OleDBC.ExecuteNonQuery()
        '  End With
        'Next xx
        '       Application.DoEvents()
    End Sub
    Public Sub Put_Trades3(ByRef R As Results, ByRef TS() As Trades)
        Static xx As Integer
        '      Dim OleDBC As New OleDbCommand
        Dim conn0 As New System.Data.OleDb.OleDbConnection
        '      OleDBC.Connection = conn0
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
                If .bEntry.qHit Then
                    .bEntry.qHiti = 255
                Else
                    .bEntry.qHiti = 0
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
                If .sEntry.qHit Then
                    .sEntry.qHiti = 255
                Else
                    .sEntry.qHiti = 0
                End If
                '                OleDBC.CommandText = "Insert Into Trades03 VALUES ('" & Counters.totalIterations &
                '               "','" & .TradeNo &
                '              "','" & .bentry.qHiti &
                '             "','" & .bDOW.qHiti &
                '            "','" & .bTD1.qHiti &
                '           "','" & .bTD2.qHiti &
                '          "','" & .bSignal.qHiti &
                '         "','" & .bentry.qHiti &
                '        "','" & .MA1.qHiti &
                '       "','" & .MA2.qHiti &
                '      "','" & .MA3.qHiti &
                '     "','" & .sExecute.qHiti &
                '    "','" & .bTD1.qHiti &
                '   "','" & .bTD2.qHiti &
                '  "','" & .sSignal.qHiti &
                ' "','" & .sExecute.qHiti &
                '"','" & Now() & "')"
                '   OleDBC.ExecuteNonQuery()
            End With
        Next xx
    End Sub
    Public Sub ConnectDatabase()
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
        '        strConnString2 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Qfe\QFE_Signals.mdb"
        '       conn2.ConnectionString = strConnString2
        conn2.Open()
        '        dbTextBox2.Text = strConnString2
        '       strConnString0 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Dad\Desktop\QFE_StockData1.mdb"
        '      conn0.ConnectionString = strConnString0
        '     MsgBox("Opening file", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, strConnString0)
        '    '        conn0.Open()
        '   dbTextBox0.Text = strConnString0
        '  strConnString1 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Qfe\TradesSmall.mdb"
        '        strConnString1 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Dad\Desktop\TradesSmall1.mdb"
        ' strConnString1 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=C:\Users\Dad\Desktop\Documents\Database211.mdb"
        ' strConnString1 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=\\WDMYCLOUD\Public\QFE_DB\Database21.mdb"
        ' conn1.ConnectionString = strConnString1
        conn1.Open()
        '       dbTextBox1.Text = strConnString1
        '        strConnString2 = "PROVIDER=microsoft.Jet.OleDb.4.0;Data Source=\\WDMYCLOUD\Public\QFE_Signals.mdb"
    End Sub
    Public Function SetLastDayTrade(ByRef R As Results, ByRef TS() As Trades, ByRef trd As Trades, ByRef thisDay As Integer) As Boolean
        With trd
            .bMonth.qHit = QExeMonth(thisDay)
            '            R.SP.bMonth.Text1 = R.SP.bMonth.Text
            If .bMonth.qHit Then
                .bMonth.qHiti = 255
            Else
                .bMonth.qHiti = 0
            End If
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
                .bTD1.qHit = QTD_(.bTD1, thisDay, .bTD1.db)
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
            '            .bTD1.Text = R.SP.bTD1.Text1
            .bTD1.Signal = .bTD1.Text & ":"
            .bTD2.Idx = R.SP.bTD2.Idx
            .bTD2.dbIdx = R.SP.bTD2.dbIdx
            .bTD2.db = R.SP.bTD2.dbText
            If .bTD2.Idx > 0 Then
                .bTD2.qHit = QTD_(.bTD2, thisDay, .bTD2.db)
                R.SP.bTD2.Text1 =
                    Strings.Left(R.SP.bTD2.Text, 12) & ":" &
                    Format$(.bTD2.executePrice, "000.000") & "on" & .bTD2.Text & "&" &
                    Format$(.bTD2.executePrice, "000.000") & "on" & .bTD2.Text
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
            .bZscoreMode.qHit = qHitPatternZ(R, TS, R.SS.Main.Trades, 0)
            If .bZscoreMode.qHit Then
                .bZscoreMode.qHiti = 255
            Else
                .bZscoreMode.qHiti = 0
            End If
            '           .bZscore.Text = R.SP.bZScore.Text & "ltpr=" & strSign(lastTrade.Profit0) & "!" & _
            '               strSign(lastTrade.Profit1) & "!" & strSign(lastTrade.Profit2) & "!" & _
            '              strSign(lastTrade.Profit3) & "!"
            If R.SP.bTrigger.Idx = 0 Then
                '                .qrig.qHit = True
                .bTrigger.Text = Strings.Left(R.SP.bTrigger.Text, 4)
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
            .sEntry.Text = R.SP.sEntry.Text
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
            '           If .bentry.qHit Then Stop
            SetLastDayTrade = .bEntry.qHit
        End With
    End Function
    Public Function SetLastDayTradenozzzex(ByRef R As Results, ByRef TS() As Trades, ByRef thisDay As Integer) As Boolean
        Dim lastDayTrade As Trades
        lastDayTrade = New Trades
        With lastDayTrade
            .bMonth.Idx = R.SP.bMonth.Idx
            .bMonth.qHit = QExeMonth(thisDay)
            '            R.SP.bMonth.Text1 = R.SP.bMonth.Text
            If .bMonth.qHit Then
                .bMonth.qHiti = 255
            Else
                .bMonth.qHiti = 0
            End If
            lastDayTrade.bMonth.Text = R.SP.bMonth.Text
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
                .bTD1.qHit = QTD_(.bTD1, thisDay, .bTD1.db)
                R.SP.bTD1.Text1 = Strings.Left(R.SP.bTD1.Text, 12) & ":$" &
                 Format$(.bTD1.actualPrice, "000.000") &
                  "on" & .bTD1.Date_ & "&$" &
                   Format$(.bTD1.triggerPrice, "000.000") & "on" & .bTD1.Date_
                .bTD1.Signal = RMain.SP.bTD1.Text
            Else
                .bTD1.qHit = True
                R.SP.bTD1.Text1 = Strings.Left(R.SP.bTD1.Text, 12) & ":$" &
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
            '            .bTD1.Text = R.SP.bTD1.Text1
            .bTD1.Signal = .bTD1.Text & ":"
            .bTD2.Idx = R.SP.bTD2.Idx
            .bTD2.dbIdx = R.SP.bTD2.dbIdx
            .bTD2.db = R.SP.bTD2.dbText
            If .bTD2.Idx > 0 Then
                .bTD2.qHit = QTD_(.bTD2, thisDay, .bTD2.db)
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
            If R.SP.bEntry.Idx = 0 Then
                .bEntry.qHit = True
                .bEntry.Text = R.SP.bEntry.Text & ":" &
                 .bEntry.Text & "@" & Format$(0.0, "000.000")
            Else
                .bEntry.qHit = QBExeEntry(thisDay)
                .bEntry.Text = R.SP.bEntry.Text & ":" &
                 .bEntry.Text & "@" & Format$(.bEntry.actualPrice, "000.000")
            End If
            If .bEntry.qHit Then
                .bEntry.qHiti = 255
            Else
                .bEntry.qHiti = 0
            End If
            '            .bZscore.qHit = QZ_(R, R.SS.Trades, 0)
            .bZscoreMode.qHit = qHitPatternZ(R, TS, R.SS.Main.Trades, 0)
            If .bZscoreMode.qHit Then
                .bZscoreMode.qHiti = 255
            Else
                .bZscoreMode.qHiti = 0
            End If
            .bZScoreValue.qHit = True
            If .bZScoreValue.qHit Then
                .bZScoreValue.qHiti = 255
            Else
                .bZScoreValue.qHiti = 0
            End If

            .bInDay.qHit = QbuyInDay(thisDay)
            If .bInDay.qHit Then
                .bInDay.qHiti = 255
                '    Stop
            Else
                .bInDay.qHiti = 0
            End If
            lastDayTrade.bEntry.qHit = lastDayTrade.bDOW.qHit And
                .bInDay.qHit And
                .bZscoreMode.qHit And
                .bZScoreValue.qHit And
                 .bTrigger.qHit And
                  .bTD1.qHit And
                  .bTD2.qHit
            If .bEntry.qHit Then
                .bEntry.qHiti = 255
            Else
                .bEntry.qHiti = 0
            End If
            .bEntry.Text = .bEntry.Text
            SetLastDayTradenozzzex = lastDayTrade.bEntry.qHit
        End With
    End Function
    Public Function TradeString(ByRef R As Results, ByRef TS() As Trades, ns As String) As String
        Static MnLastTr As Integer, stats As String
        MnLastTr = R.SS.Main.Trades
        stats = " Tr:" & Format$(R.SS.Main.Trades, "0000") & "%" & Format(R.SS.Main.wPcntg, "0.00") &
                 " W:" & Format(R.SS.Main.W, "0000") & "/L:" & Format(R.SS.Main.L, "0000") & " "
        Select Case MnLastTr
            Case > 13
                TradeString = ns & "==" & stats & "01" &
         StrSign(TS(MnLastTr).Profit0) & "#" & TS(MnLastTr).bDate & "^02" &
         StrSign(TS(MnLastTr - 1).Profit0) & "#" & TS(MnLastTr - 1).bDate & "^03" &
         StrSign(TS(MnLastTr - 2).Profit0) & "#" & TS(MnLastTr - 2).bDate & "^04" &
         StrSign(TS(MnLastTr - 3).Profit0) & "#" & TS(MnLastTr - 3).bDate & "^05" &
         StrSign(TS(MnLastTr - 4).Profit0) & "#" & TS(MnLastTr - 4).bDate & "^06" &
         StrSign(TS(MnLastTr - 5).Profit0) & "#" & TS(MnLastTr - 5).bDate & "^07" &
         StrSign(TS(MnLastTr - 6).Profit0) & "#" & TS(MnLastTr - 6).bDate & "^08" &
         StrSign(TS(MnLastTr - 7).Profit0) & "#" & TS(MnLastTr - 7).bDate & "^09" &
         StrSign(TS(MnLastTr - 8).Profit0) & "#" & TS(MnLastTr - 8).bDate & "^10" &
         StrSign(TS(MnLastTr - 9).Profit0) & "#" & TS(MnLastTr - 9).bDate & "^11" &
         StrSign(TS(MnLastTr - 10).Profit0) & "#" & TS(MnLastTr - 10).bDate & "^12" &
         StrSign(TS(MnLastTr - 11).Profit0) & "#" & TS(MnLastTr - 11).bDate & "^13" &
         StrSign(TS(MnLastTr - 12).Profit0) & "#" & TS(MnLastTr - 12).bDate & "^14" &
         StrSign(TS(MnLastTr - 13).Profit0) & "#" & TS(MnLastTr - 13).bDate
            Case = 13
                TradeString = ns & "==" & stats & "01" &
         StrSign(TS(MnLastTr).Profit0) & "#" & TS(MnLastTr).bDate & "^02" &
         StrSign(TS(MnLastTr - 1).Profit0) & "#" & TS(MnLastTr - 1).bDate & "^03" &
         StrSign(TS(MnLastTr - 2).Profit0) & "#" & TS(MnLastTr - 2).bDate & "^04" &
         StrSign(TS(MnLastTr - 3).Profit0) & "#" & TS(MnLastTr - 3).bDate & "^05" &
         StrSign(TS(MnLastTr - 4).Profit0) & "#" & TS(MnLastTr - 4).bDate & "^06" &
         StrSign(TS(MnLastTr - 5).Profit0) & "#" & TS(MnLastTr - 5).bDate & "^07" &
         StrSign(TS(MnLastTr - 6).Profit0) & "#" & TS(MnLastTr - 6).bDate & "^08" &
         StrSign(TS(MnLastTr - 7).Profit0) & "#" & TS(MnLastTr - 7).bDate & "^09" &
         StrSign(TS(MnLastTr - 8).Profit0) & "#" & TS(MnLastTr - 8).bDate & "^10" &
         StrSign(TS(MnLastTr - 9).Profit0) & "#" & TS(MnLastTr - 9).bDate & "^11" &
         StrSign(TS(MnLastTr - 10).Profit0) & "#" & TS(MnLastTr - 10).bDate & "^12" &
         StrSign(TS(MnLastTr - 11).Profit0) & "#" & TS(MnLastTr - 11).bDate & "^13" &
         StrSign(TS(MnLastTr - 12).Profit0) & "#" & TS(MnLastTr - 12).bDate
            Case = 12
                TradeString = ns & "==" & stats & "01" &
         StrSign(TS(MnLastTr).Profit0) & "#" & TS(MnLastTr).bDate & "^02" &
         StrSign(TS(MnLastTr - 1).Profit0) & "#" & TS(MnLastTr - 1).bDate & "^03" &
         StrSign(TS(MnLastTr - 2).Profit0) & "#" & TS(MnLastTr - 2).bDate & "^04" &
         StrSign(TS(MnLastTr - 3).Profit0) & "#" & TS(MnLastTr - 3).bDate & "^05" &
         StrSign(TS(MnLastTr - 4).Profit0) & "#" & TS(MnLastTr - 4).bDate & "^06" &
         StrSign(TS(MnLastTr - 5).Profit0) & "#" & TS(MnLastTr - 5).bDate & "^07" &
         StrSign(TS(MnLastTr - 6).Profit0) & "#" & TS(MnLastTr - 6).bDate & "^08" &
         StrSign(TS(MnLastTr - 7).Profit0) & "#" & TS(MnLastTr - 7).bDate & "^09" &
         StrSign(TS(MnLastTr - 8).Profit0) & "#" & TS(MnLastTr - 8).bDate & "^10" &
         StrSign(TS(MnLastTr - 9).Profit0) & "#" & TS(MnLastTr - 9).bDate & "^11" &
         StrSign(TS(MnLastTr - 10).Profit0) & "#" & TS(MnLastTr - 10).bDate & "^12" &
         StrSign(TS(MnLastTr - 11).Profit0) & "#" & TS(MnLastTr - 11).bDate
            Case = 11
                TradeString = ns & "==" & stats & "01" &
         StrSign(TS(MnLastTr).Profit0) & "#" & TS(MnLastTr).bDate & "^02" &
         StrSign(TS(MnLastTr - 1).Profit0) & "#" & TS(MnLastTr - 1).bDate & "^03" &
         StrSign(TS(MnLastTr - 2).Profit0) & "#" & TS(MnLastTr - 2).bDate & "^04" &
         StrSign(TS(MnLastTr - 3).Profit0) & "#" & TS(MnLastTr - 3).bDate & "^05" &
         StrSign(TS(MnLastTr - 4).Profit0) & "#" & TS(MnLastTr - 4).bDate & "^06" &
         StrSign(TS(MnLastTr - 5).Profit0) & "#" & TS(MnLastTr - 5).bDate & "^07" &
         StrSign(TS(MnLastTr - 6).Profit0) & "#" & TS(MnLastTr - 6).bDate & "^08" &
         StrSign(TS(MnLastTr - 7).Profit0) & "#" & TS(MnLastTr - 7).bDate & "^09" &
         StrSign(TS(MnLastTr - 8).Profit0) & "#" & TS(MnLastTr - 8).bDate & "^10" &
         StrSign(TS(MnLastTr - 9).Profit0) & "#" & TS(MnLastTr - 9).bDate & "^11" &
         StrSign(TS(MnLastTr - 10).Profit0) & "#" & TS(MnLastTr - 10).bDate
            Case = 10
                TradeString = ns & "==" & stats & "01" &
         StrSign(TS(MnLastTr).Profit0) & "#" & TS(MnLastTr).bDate & "^02" &
         StrSign(TS(MnLastTr - 1).Profit0) & "#" & TS(MnLastTr - 1).bDate & "^03" &
         StrSign(TS(MnLastTr - 2).Profit0) & "#" & TS(MnLastTr - 2).bDate & "^04" &
         StrSign(TS(MnLastTr - 3).Profit0) & "#" & TS(MnLastTr - 3).bDate & "^05" &
         StrSign(TS(MnLastTr - 4).Profit0) & "#" & TS(MnLastTr - 4).bDate & "^06" &
         StrSign(TS(MnLastTr - 5).Profit0) & "#" & TS(MnLastTr - 5).bDate & "^07" &
         StrSign(TS(MnLastTr - 6).Profit0) & "#" & TS(MnLastTr - 6).bDate & "^08" &
         StrSign(TS(MnLastTr - 7).Profit0) & "#" & TS(MnLastTr - 7).bDate & "^09" &
         StrSign(TS(MnLastTr - 8).Profit0) & "#" & TS(MnLastTr - 8).bDate & "^10" &
         StrSign(TS(MnLastTr - 9).Profit0) & "#" & TS(MnLastTr - 9).bDate
            Case Else
                Stop
                TradeString = "none"
        End Select
    End Function
    Public Sub calcPeriodProfits(ByRef R As Results, ByRef TS() As Trades)
        Static tr04 As Integer, tr05 As Integer, tr06 As Integer, tr07 As Integer
        Static tr00 As Integer, tr01 As Integer, tr02 As Integer, tr03 As Integer, tr08 As Integer
        Static tr09 As Integer, tr10 As Integer, tr11 As Integer, tr12 As Integer, tr13 As Integer
        Static tr14 As Integer, tr15 As Integer
        Static div As Single, trds As Integer
        trds = R.SS.Main.Trades
        div = R.SS.Main.Trades / 16
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
        R.SS.Main.profits00 = TS(tr00).totProfit '/ TS(tr00).totDH / 10
        R.SS.Main.profits01 = TS(tr01).totProfit '/ TS(tr01).totDH / 10
        R.SS.Main.profits02 = TS(tr02).totProfit '/ TS(tr02).totDH / 10
        R.SS.Main.profits03 = TS(tr03).totProfit '/ TS(tr03).totDH / 10
        R.SS.Main.profits04 = TS(tr04).totProfit '/ TS(tr04).totDH / 10
        R.SS.Main.profits05 = TS(tr05).totProfit '/ TS(tr05).totDH / 10
        R.SS.Main.profits06 = TS(tr06).totProfit '/ TS(tr06).totDH / 10
        R.SS.Main.profits07 = TS(tr07).totProfit '/ TS(tr07).totDH / 10
        R.SS.Main.profits08 = TS(tr08).totProfit '/ TS(tr08).totDH / 10
        R.SS.Main.profits09 = TS(tr09).totProfit '/ TS(tr09).totDH / 10
        R.SS.Main.profits10 = TS(tr10).totProfit '/ TS(tr10).totDH / 10
        R.SS.Main.profits11 = TS(tr11).totProfit '/ TS(tr11).totDH / 10
        R.SS.Main.profits12 = TS(tr12).totProfit '/ TS(tr12).totDH / 10
        R.SS.Main.profits13 = TS(tr13).totProfit '/ TS(tr13).totDH / 10
        R.SS.Main.profits14 = TS(tr14).totProfit '/ TS(tr14).totDH / 10
        R.SS.Main.profits15 = TS(tr15).totProfit '/ TS(tr15).totDH / 10
        R.SS.Main.Profits.Pcntg = R.SS.Main.Profits.avg / 10
    End Sub
    Public Function QbuyInDaySignal(ByRef tday As Integer) As String
        Static str1 As String, str2 As String, str3 As String, str4 As String, q As Boolean ', str5 As String, str6 As String
        q = False
        Select Case RMain.SP.bInday.Idx
            Case 0
                q = True
                str1 = Format(0, "000.000")
                str2 = Format(0, "000.000")
                str3 = Format(0, "000.000")
                str4 = Format(0, "000.000")
            Case 1
                str1 = "y" & lastDayTrade.bInDay.qHit & "-" & Format$(hi(tday), "000.000") & "<"
                str2 = Format(hi(tday - 1), "000.000") & "::"
                str3 = Format(lo(tday), "000.000") & ">"
                str4 = Format(lo(tday - 1), "000.000")
                q = (hi(tday) < hi(tday - 1) And lo(tday) > lo(tday - 1))
            Case 2
                str1 = "n" & lastDayTrade.bInDay.qHit & "-" & Format(hi(tday), "000.000") & "<"
                str2 = Format(hi(tday - 1), "000.000") & "::"
                str3 = Format(lo(tday), "000.000") & ">"
                str4 = Format(lo(tday - 1), "000.000")
                q = Not (hi(tday) < hi(tday - 1) And lo(tday) > lo(tday - 1))
            Case 3
                str1 = "y" & lastDayTrade.bInDay.qHit & "-" & Format(hi(tday), "000.000") & "<"
                str2 = Format(hi(tday - 2), "000.000") & "::"
                str3 = Format(lo(tday), "000.000") & ">"
                str4 = Format(lo(tday - 2), "000.000")
                q = (hi(tday) < hi(tday - 2) And lo(tday) > lo(tday - 2))
            Case 4
                str1 = "n" & lastDayTrade.bInDay.qHit & "-" & Format(hi(tday), "000.000") & "<"
                str2 = Format(hi(tday - 2), "000.000") & "::"
                str3 = Format(lo(tday), "000.000") & ">"
                str4 = Format(lo(tday - 2), "000.000")
                q = Not (hi(tday) < hi(tday - 2) And lo(tday) > lo(tday - 2))
            Case 5
                str1 = "y" & lastDayTrade.bInDay.qHit & "-" & Format(hi(tday), "000.000") & "<"
                str2 = Format(hi(tday - 3), "000.000") & "::"
                str3 = Format(lo(tday), "000.000") & ">"
                str4 = Format(lo(tday - 3), "000.000")
                q = (hi(tday) < hi(tday - 3) And lo(tday) > lo(tday - 3)) 'And cl(tday) > cl(tday - 1)
            Case 6
                str1 = "n" & lastDayTrade.bInDay.qHit & "-" & Format(hi(tday), "000.000") & "<"
                str2 = Format(hi(tday - 3), "000.000") & "::"
                str3 = Format(lo(tday), "000.000") & ">"
                str4 = Format(lo(tday - 3), "000.000")
                q = (hi(tday) < hi(tday - 3) And lo(tday) > lo(tday - 3)) 'And cl(tday) < cl(tday - 1)
            Case 7
                str1 = "y" & lastDayTrade.bInDay.qHit & "-" & Format(hi(tday), "000.000") & "<"
                str2 = Format(hi(tday - 4), "000.000") & "::"
                str3 = Format(lo(tday), "000.000") & ">"
                str4 = Format(lo(tday - 4), "000.000")
                q = (hi(tday) < hi(tday - 4) And lo(tday) > lo(tday - 4)) 'And cl(tday) > cl(tday - 1)
            Case 8
                str1 = "n" & lastDayTrade.bInDay.qHit & "-" & Format(hi(tday), "000.000") & "<"
                str2 = Format(hi(tday - 4), "000.000") & "::"
                str3 = Format(lo(tday), "000.000") & ">"
                str4 = Format(lo(tday - 4), "000.000")
                q = (hi(tday) < hi(tday - 4) And lo(tday) > lo(tday - 4)) 'And cl(tday) < cl(tday - 1)
            Case Else
                Stop
        End Select
        QbuyInDaySignal = str1 & str2 & str3 & str4
    End Function
    Public Sub Put_Signals(ByRef R As Results, ByRef TS() As Trades, conNo As Integer)
        Dim Z_Trades As String, mnTrades As String, onOffStr As String, xtmp As Integer
        Dim statStr As String
        R.SS.Main.Q.avg = Calculate_BasicStatistics(R, TS, 1, R.SS.Main.Trades)
        Counters.iterationsWritten = Counters.iterationsWritten + 1
        If Counters.qBuyandHold.qHit Then
            Counters.qBuyandHold.qHiti = 255
        Else
            Counters.qBuyandHold.qHiti = 0
        End If
        'R.SS.Main.zScore.V = CalcZScore(R, TS, 12, R.SS.Main.Trades)
        mnTrades = TradeString(RMain, RMainTS, "mainTrades")
        Z_Trades = TradeString(R, TS, "ZTrades" & R.SP.bZScoreMode.Text)
        If R.SS.Main.Q.avg > 0.0 Then
            QFE1.lSymbol.ForeColor = Color.Crimson
        Else
            QFE1.lSymbol.ForeColor = Color.Gainsboro
        End If
        'QFE1.lSymbol.Text = R.lSymbol
        'rows = Form1.DGViewSignals.Rows.Count
        '      QFE1.DGViewSignals.Font.Size = 14
        ' Form1.DGViewSignals.Rows.Add(Now, R.sSymbol, .SS.Main.Q.avg, .SS.Main.absQuantum, R.lSymbol)
        QFE1.DGViewSignals.AllowUserToAddRows = True
        '            Form1.DGViewSignals.Rows.Item(rows - 1).Cells(2).Value = R.lSymbol
        '            DataGridViewSecurities.Rows.Add(foundFile)
        '            Form1.DataGridViewSecurities.Rows.Item(rowss).Cells(7).Value = Format(Counters.totalIterations, "00000")
        statStr = "QQ==" & StrSign(R.SS.Main.Q.avg) &
             "|avgDH=" & Format(R.SS.Main.avgDH, "0.00") &
              "|DH=" & Format(R.SS.Main.DH.tot, "0.0") &
               "|avgPr=" & StrSign(R.SS.Main.Profits.avg) &
                "|Pr=" & StrSign(R.SS.Main.Profits.tot) &
                 "|Tr=" & Format(R.SS.Main.Trades, "00000") &
                  "|W=" & StrSign(R.SS.Main.W) &
                   "|L=" & StrSign(R.SS.Main.L) &
                    "%=" & Format(R.SS.Main.wPcntg, "0.00") &
                     "|ZScrM=" & R.SP.bZScoreMode.Text &
                     "|ZScrV=" & StrSign(R.SS.Main.zScore.V) &
                      "|Runs=" & Format(R.SS.Main.Runs, "0000") &
                       "|expR=" & Format(R.SS.Main.expRuns, "0000") &
                        "--Correl=" & StrSign(R.SS.Main.Correlation.V) &
                         "--COV" & StrSign(R.SS.Main.COV.V)
        '           '           OleDBC.CommandText = "Insert Into Base VALUES ('" & Counters.Iteration &
        If lastDayTrade.bEntry.qHit Then
            onOffStr = "T"
        Else
            onOffStr = "F"
        End If
        With R
            QFE1.Q.Text = R.SS.Main.Q.avg
            If R.SS.Main.sTP.Hits = 0 Then
                R.SS.Main.sTP.Pcntg = 0.0
            Else
                R.SS.Main.sTP.Pcntg = R.SS.Main.sTP.Hits / R.SS.Main.Trades
            End If
            QFE1.Days_.Text = R.SS.Main.Days
        R.SP.bInday.Text1 = R.SP.bInday.Text1 & ": " & RHitMisses.SS.Main.bInDay.Hits
            R.SS.Main.Correlation.V = 0.0
            Dim column As DataGridViewColumn = QFE1.SignalsGrid.Columns(Counters.SystemNumber)
            column.Width = 450
            Dim row As DataGridViewRow = QFE1.SignalsGrid.Rows(Counters.SecNo)
            row.Height = 120
            R.SS.Main.COV.V = 0
            xtmp = InStr(Counters.currentSecurity, " ") - 1
            If xtmp > 0 Then
                Counters.currentSecurity = Strings.Left(Counters.currentSecurity, xtmp)
            End If
            R.Symbol = Strings.Left(RTrim(Counters.currentSecurity) & "_____", 5)
            R.sSymbol = R.Symbol & "#" & StrSign(R.SS.Main.Q.avg)
            R.lSymbol = R.sSymbol & "|" & R.SP.bDOW.Text & "|" & R.SP.bEntry.Text & ":" &
             Format(R.SP.sMaxDH.V, "00") & ":" &
              R.SP.sEntry.Text & "|" &
               R.SP.TP.Text
            R.signal = StrLD(R, TS)
            RMain.tradesString = mnTrades ' mainTradesStr(RMainTS)
            R.tradesString = Z_Trades ' zScoreTradesStr(R, TS)
            If R.SS.Main.Profits.tot = Single.NaN Then Stop
            '            QFE1.SignalsGrid.Rows(Counters.SecNo).Cells(Counters.SystemNumber).t
            '            QFE1.SignalsGrid.Rows(Counters.SecNo).Cells(Counters.SystemNumber).fo
            QFE1.lSymbol.Text = R.signal
            QFE1.SignalsGrid.Rows(Counters.SecNo).Cells(Counters.SystemNumber).Value = R.lSymbol & ":" & StrSign(R.SS.Main.Q.avg) & ":" & R.signal
            '            R.signal = R.sSymbol & onOffStr & "|" & Strings.Left(R.SP.bDOW.Text, 3) & "|tg" &
            '           Strings.Left(R.SP.bMonth.Text, 5) & R.SP.bEntry.Text1 & ":" &
            '          R.SP.bTrigger.Text & "!mx:" & Format(R.SP.sMaxDH.V, "00.0") & "|tp=" &
            '         R.SP.TP.Text1
            QFE1.oleDatabaseSignals.CommandText = "Insert Into Signals VALUES ('" & Counters.Iteration &
  "','" & onOffStr &
  "','" & Counters.SystemNumber &
  "','" & Counters.ZIterBase &
  "','" & Counters.ZIterNo &
   "','" & R.Symbol &
   "','" & R.sSymbol &
   "','" & R.lSymbol &
   "','" & R.signal &
    "','" & Counters.last_Date &
     "','" & Left(lastDayTrade.bEntry.Text, 6) &
      "','" & R.SS.Main.Q.avg &
       "','" & R.SS.Main.absQuantum &
        "','" & R.SS.Main.Q.min &
         "','" & R.SS.Main.Q.max &
          "','" & R.SS.Main.Days &
           "','" & R.SS.Main.Days &
            "','" & Format(R.SS.Main.Trades / R.SS.Main.Days, "0.00") &
             "','" & R.SS.Main.Trades &
              "','" & R.SS.Main.W &
              "','" & R.SS.Main.L &
               "','" & R.SS.Main.wPcntg &
                "','" & R.SS.Main.DH.avg &
                 "','" & R.SP.sMaxDH.V &
                  "','" & R.SS.Main.DH.tot &
                   "','" & R.SP.TP.V &
                    "','" & R.SS.Main.sTP.Hits &
                     "','" & R.SS.Main.sTP.Misses &
                      "','" & R.SS.Main.COV.V &
                       "','" & R.SS.Main.Correlation.V &
                       "','" & R.SS.Main.Profits.tot &
                       "','" & R.SS.Main.Profits.avg &
                        "','" & lastDayTrade.bEntry.qHiti &
                        "','" & R.SS.Main.bEntryStats.Hits &
                        "','" & R.SS.Main.bEntryStats.Misses &
                        "','" & lastDayTrade.bEntry.Text &
                        "','" & lastDayTrade.sEntry.qHiti &
                        "','" & R.SS.Main.sEntryStats.Hits &
                        "','" & R.SS.Main.sEntryStats.Misses &
                            "','" & R.SP.sEntry.Text &
                                "','" & lastDayTrade.bMonth.qHiti &
                                 "','" & R.SS.Main.bMonth.Hits &
                                  "','" & R.SS.Main.bMonth.Misses &
                                   "','" & Strings.Left(R.SP.bMonth.Text, 4) &
               "','" & lastDayTrade.bDOW.qHiti &
              "','" & R.SS.Main.bDOW.Hits &
                   "','" & R.SS.Main.bDOW.Misses &
                   "','" & Strings.Left(R.SP.bDOW.Text, 4) &
                  "','" & lastDayTrade.bInDay.qHiti &
                 "','" & RHitMisses.SS.Main.bInDay.Hits &
                "','" & RHitMisses.SS.Main.bInDay.Misses &
               "','" & lastDayTrade.bInDay.Text &
               "','" & lastDayTrade.bInDay.Signal &
               "','" & lastDayTrade.bTD1.qHiti &
              "','" & RHitMisses.SS.Main.bTD1.Hits &
                     "','" & RHitMisses.SS.Main.bTD1.Misses &
                    "','" & R.SP.bTD1.Text1 &
                   "','" & lastDayTrade.bTD2.qHiti &
                   "','" & RHitMisses.SS.Main.bTD2.Hits &
                  "','" & RHitMisses.SS.Main.bTD2.Misses &
                 "','" & R.SP.bTD2.Text1 &
                "','" & R.SS.Main.zScore.V &
               "','" & RHitMisses.SS.Main.zScore.Misses &
                "','" & lastDayTrade.bZscoreMode.qHiti &
                 "','" & RHitMisses.SS.Main.zScore.Hits &
                  "','" & RHitMisses.SS.Main.zScore.Misses &
                   "','" & R.SP.bZScoreMode.Idx &
                    "','" & R.SP.bZScoreMode.Text & StrSign(lastTrade.Profit0) & ":" & StrSign(TS(R.SS.Main.Trades - 1).Profit0) & ":" & StrSign(TS(R.SS.Main.Trades - 1).Profit0) &
                     "','" & lastDayTrade.bZScore.qHiti &
                      "','" & RHitMisses.SS.Main.zScore.Hits &
                       "','" & RHitMisses.SS.Main.zScore.Misses &
                        "','" & R.SP.bZScoreValue.Text &
                   "','" & R.SS.Main.Runs &
                  "','" & R.SS.Main.expRuns &
                  "','" & R.SS.Main.zScore.min &
                 "','" & R.SS.Main.zScore.max &
                "','" & R.lSymbol &
               "','" & Counters.start_Date &
             "','" & Counters.end_Date &
               "','" & R.tradesString &
               "','" & RMain.tradesString &
               "','" & statStr &
  "','" & Now() & "')"
            QFE1.oleDatabaseSignals.ExecuteNonQuery()
            '                           "','" & R.SS.Main.profits00 &
            ''                          "','" & R.SS.Main.profits01 &
            '                         "','" & R.SS.Main.profits02 pro
            '                        "','" & R.SS.Main.profits03 &
            '                       "','" & R.SS.Main.profits04 &
            '                     "','" & R.SS.Main.profits05 &
            '                     "','" & R.SS.Main.profits06 &
            '                    "','" & R.SS.Main.profits07 &
            ''                  "','" & R.SS.Main.profits08 &
            '                  "','" & R.SS.Main.profits09 &
            '                "','" & R.SS.Main.profits10 &
            '                                           "','" & R.SS.Main.profits11 &
            '                                           "','" & R.SS.Main.profits12 &
            '                                           "','" & R.SS.Main.profits13 &
            '                                           "','" & R.SS.Main.profits14 &
            '                                           "','" & R.SS.Main.profits15 &

            ' "','" & R.SS.Main.Trades &
            '  "','" & R.sSymbol & "on" & Counters.last_Day & "&" & Format$(cl(Counters.last_Day), "000.000") & "=" &
            '      R.SP.bInday.Text & "_" & R.SP.bTD1.Text & R.SP.bTD2.Text & "_" & R.SP.bEntry.Text1 & "_" & R.SP.sMaxDH.V & "_" &
            '     R.SP.sEntry.Text1 &
            '       "','" & R.SP.bTrigger.Text & ":" &
            '   Format(lo(Counters.last_Day), "000.000") & "<=" &
            '  Format(TS(R.SS.Main.Trades).bTrigger.triggerPrice, "000.000") &
            '  "','" & Format(op(Counters.last_Day), "000.000") & "!" & Format(op(Counters.last_Day - 1), "000.000") &
            '  "','" & Format(hi(Counters.last_Day), "000.000") & "!" & Format(hi(Counters.last_Day - 1), "000.000") &
            '  "','" & Format(lo(Counters.last_Day), "000.000") & "!" & Format(lo(Counters.last_Day - 1), "000.000") &
            '  "','" & Format(cl(Counters.last_Day), "000.000") & "!" & Format(op(Counters.last_Day - 1), "000.000") &
            '                        "','" & Counters.end_Date &
            '                       "','" & R.SS.Main.profits14 &
            '                      "','" & R.SS.Main.profits15 &
            '                     "','" & R.TFSignal &
            '     "','" & Now() & "')"
            '                     " ','" & Counters.start_Day &
            '                     "','" & Counters.end_day &
            '                     "','" & Counters.last_Date &
            '     '      ' " ','" & R.SS.Main.Q1.avg &
            '"','" & R.SS.Main.Q1.min &
            '            "','" & R.SS.Main.Q1.max &
            '      "','" & lastDayTrade.bTrigger.triggerPrice &
            '     "','" & R.SP.bTrigger.Text1 &
            '    "','" & RHitMisses.SS.Main.bTrigger.Hits &
            '   "','" & R.SP.bTD1.Text1 &
            ' "','" & R.SP.bEntry.Text & "!!" & Format$("00", R.SP.sMaxDH.V) & ":" & Format$("0.0", R.SS.Main.avgDH) & "!!" & R.SP.sEntry.Text &
            '"','" & statStr &
            '            "','" & Z_Trades &
            '       "','" & R.SS.Main.zScore.V &
            '      "','" & TS(R.SS.Main.Trades).Profit0 &
            '     "','" & TS(R.SS.Main.Trades).bDate &
            '    "','" & Now() & "')"
            '   OleDBC.ExecuteNonQuery()
            '  '            "','" & " " & 'R.lSymbol &
            '           "','" & R.sSymbol & ":" & StrSign(R.SS.Main.Q.avg) & R.SP.bDOW.Text & ":" &
            '         "','" & R.sSymbol & ":" & .SP.bZScoreMode.Text & ":" & R.SP.bZScoreStd.Text &
            '        "','" & R.sSymbol & R.SP.bTrigger.Text &
            '       "','" & R.sSymbol & "!" & ' R.SP.bTD1.Text & "!" & R.SP.bTD2.Text & "!" &
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
            '         "','" & R.TS(R.SS.Main.Trades).Profit0 &
            '        "','" & R.TS(R.SS.Main.Trades - 1).Profit0 &
            '       "','" & R.TS(R.SS.Main.Trades - 2).Profit0 &
            '      "','" & R.TS(R.SS.Main.Trades - 3).Profit0 &
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
            '            "','" & Now() & "')"
            '           Form1.DGViewSignals.Rows.Add(olestr)
            '''            Form1.oleConn.Close()
        End With
    End Sub
    Public Sub Put_Totals(ByRef R As Results, ByRef TS() As Trades, conNo As Integer)
        QFE1.oleDatabaseSignals.CommandText = "Insert Into Signals VALUES ('" & Counters.Iteration &
                           "','" & R.SS.Main.profits00 &
                          "','" & R.SS.Main.profits01 &
                         "','" & R.SS.Main.profits02 &
                        "','" & R.SS.Main.profits03 &
                       "','" & R.SS.Main.profits04 &
                      "','" & R.SS.Main.profits05 &
                     "','" & R.SS.Main.profits06 &
                    "','" & R.SS.Main.profits07 &
                   "','" & R.SS.Main.profits08 &
                  "','" & R.SS.Main.profits09 &
                 "','" & R.SS.Main.profits10 &
                                             "','" & R.SS.Main.profits11 &
                                             "','" & R.SS.Main.profits12 &
                                             "','" & R.SS.Main.profits13 &
                                             "','" & R.SS.Main.profits14 &
                                             "','" & R.SS.Main.profits15 &
                    "','" & R.lSymbol &
                "','" & Now() & "')"
        QFE1.oleDatabaseSignals.ExecuteNonQuery()
    End Sub
    Private Sub Put_Base(ByRef R As Results, conNo As Integer)
        With R
            Select Case conNo
                Case 1
        '            OleDBC.Connection = conn1
                Case 2
                    '           OleDBC.Connection = conn2
                Case Else
                    Stop
            End Select
            '           OleDBC.CommandText = "Insert Into Base VALUES ('" & Counters.Iteration &
            '          "','" & Counters.SystemNumber &
            '         "','" & ";;" &
            '        "','" & R.sSymbol &
            '       "','" & Counters.this_Day &
            '      "','" & Counters.this_Date &
            '     "','" & Counters.first_Day &
            '    "','" & Counters.first_Date &
            '   "','" & Counters.start_Day &
            '  "','" & Counters.start_Date &
            ' "','" & Counters.last_Day &
            '"','" & Counters.last_Date &
            '       "','" & Counters.end_day &
            '      "','" & Counters.end_Date &
            '     "','" & Now() & "')"
            '       OleDBC.ExecuteNonQuery()
            'End If
            'End If
        End With
        '    Me.SignalsGrid.Refresh()
        Application.DoEvents()
    End Sub
    Public Sub Put_LastDay(ByRef R As Results, conNo As Integer, ByRef lt As Trades)
        Static bEntryStr As String
        Select Case conNo
            Case 1
         '       OleDBC.Connection = conn1
            Case 2
                '        OleDBC.Connection = conn2
            Case Else
                Stop
        End Select
        With R
            bEntryStr = ""
            Select Case R.SP.bSignal.Idx
                Case 0
                    lastTrade.bEntry.executePrice = 0.001
                    bEntryStr = "***.***"
                Case 1
                    lastTrade.bEntry.executePrice = op(Counters.end_day - 1)
                    bEntryStr = Format$(op(Counters.end_day - 1), "000.000")
                Case 2
                    lastTrade.bEntry.executePrice = cl(Counters.end_day - 1)
                    bEntryStr = Format$(cl(Counters.end_day - 1), "000.000")
                Case 3
                    lastTrade.bEntry.executePrice = op(Counters.end_day)
                    bEntryStr = Format$(op(Counters.end_day), "000.000")
                Case 4
                    lastTrade.bEntry.executePrice = cl(Counters.end_day)
                    bEntryStr = Format$(cl(Counters.end_day), "000.000")
                Case 5
                    lastTrade.bEntry.executePrice = op(Counters.end_day)
                    bEntryStr = Format$(op(Counters.end_day), "000.000")
                Case 6
                    lastTrade.bEntry.executePrice = cl(Counters.end_day)
                    bEntryStr = Format$(cl(Counters.end_day), "000.000")
                Case 7
                    lastTrade.bEntry.executePrice = op(Counters.end_day)
                    bEntryStr = Format$(lastTrade.bEntry.executePrice, "000.000")
                Case 8
                    lastTrade.bEntry.executePrice = cl(Counters.end_day)
                    bEntryStr = Format$(lastTrade.bEntry.executePrice, "000.000")
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
            '            OleDBC.CommandText = "Insert Into lastDay VALUES ('" & Counters.Iteration &
            '           "','" & Counters.SystemNumber &
            '          "','" & .SP.bDOW.Text &
            '         "','" & lt.bentry.qHiti &
            '        "','" & lt.bTrigger.qHiti &
            '       "','" & lt.bZscore.qHiti &
            '      "','" & lt.bDOW.qHiti &
            '     "','" & lt.bTD1.qHiti &
            '    "','" & lt.bTD2.qHiti &
            '   "','" & lt.bEntry.qHiti &
            '  "','" & lt.bTrigger.Text & "-" &
            ' "','" & 0.0 &
            '"','" & dtStr2(Counters.end_day) &
            '            "','" & op(Counters.end_day) &
            '           "','" & hi(Counters.end_day) &
            '          "','" & lo(Counters.end_day) &
            '         "','" & cl(Counters.end_day) &
            '        "','" & dtStr2(Counters.end_day - 1) &
            '       "','" & op(Counters.end_day - 1) &
            '      "','" & hi(Counters.end_day - 1) &
            '     "','" & lo(Counters.end_day - 1) &
            '    "','" & cl(Counters.end_day - 1) &
            '   "','" & lt.bentry.Text &
            '  "','" & lt.Profit0 &
            ' "','" & op(Counters.end_day) &
            '"','" & hi(Counters.end_day) &
            '            "','" & lo(Counters.end_day) &
            '           "','" & cl(Counters.end_day) &
            '          "','" & lt.Profit1 &
            '         "','" & lt.Profit2 &
            '        "','" & lt.Profit3 &
            '       "','" & Now() & "')"
            '      OleDBC.ExecuteNonQuery()
        End With
    End Sub
    Public Sub Put_Signal_(ByRef R As Results, conNo As Integer, ByRef lt As Trades)
        '  Static OleDBC As New OleDbCommand
        Call Put_LastDay(R, conNo, lt)
        With R
            Select Case conNo
                Case 1
        '            OleDBC.Connection = conn1
                Case 2
                    '             OleDBC.Connection = conn2
                Case Else
                    Stop
            End Select
            'End If
            '            tdStr = Strings.Left(R.SP.bTD1.Text, 5) & ":" & Format()
            If R.SS.Main.Trades = 0 Then Stop
            R.sSymbol = Counters.currentSecurity
            R.SP.Buy.Text = R.sSymbol & "-" & Counters.this_Date & "-" &
                "-tr=" & Strings.Left(.SP.bTrigger.Text, 13) & "@" & Format$(lt.bTrigger.triggerPrice, "000.000") &
                "#" & lt.bSignal.Text & ":" &
                "td:" & lt.bTD1.Text & "@" & lt.bTD2.Text & "@" &
                .SP.bDOW.Text & ":" _
                & Strings.Left(.SP.bEntry.Text, 4) & ":" & lt.bZscoreMode.Text & ":" &
                Format$(.SP.sMaxDH.V, "00") & "/" & Format$(.SS.Main.DH.avg, "00.00") & ":" &
                Strings.Left(.SP.sEntry.Text, 5) & "-on-" & dtStr1(Counters.end_day) &
                "avPr=" & StrSign(Format$(.SS.Main.Profits.avg, "00.000")) &
                "Q==" & StrSign(Format$(.SS.Main.Q.avg, "00.000")) &
                "DTWL%" & Format$(.SS.Main.Days, "0000") & "-" & Format$(.SS.Main.Trades, "0000") & "!" &
                Format$(.SS.Main.Winners.tot, "000") & "/" & Format$(.SS.Main.Losers.tot, "000") & "%" &
                Format$(.SS.Main.wPcntg, "0.00")
            'If R.SP.bZScore.Idx = 2 And R.TS(R.SS.Trades).bZscore.qHit And R.TS(R.SS.Trades - 1).Profit <= 0.0 Then Stop
            '            OleDBC.CommandText = "Insert Into Signal_ VALUES ('" & Counters.Iteration &
            '           "','" & Counters.SystemNumber &
            '          "','" & lt.bentry.qHiti &
            '         "','" & .SP.bDOW.Text &
            '        "','" & .SS.Main.Days &
            '       "','" & R.SS.Main.Trades &
            '      "','" & Format$(.SP.sMaxDH.V, "00") &
            '     "','" & Strings.Left(.SP.bEntry.Text, 4) & ":" & Strings.Left(.SP.sEntry.Text, 5) &
            '    "','" & Strings.Left(lt.bTD1.Signal, 11) & ":" & Strings.Left(lt.bTD2.Signal, 11) &
            '   "','" & .sSymbol & ":" & Format$(R.SS.Main.Trades, "0000") & ":" & .SP.bDOW.Text & ":" &
            '  Format$(.SS.Main.bDOW.Hits, "000") & ":" & .SP.bZScoreMode.Text &
            ' Strings.Left(lt.bTD1.Signal, 11) & ":" & Format$(R.SS.Main.bTD1.Hits, "000") & "-" &
            '        Strings.Left(lt.bTD2.Signal, 11) & ":" & Format$(R.SS.Main.bTD2.Hits, "000") & "-" &
            '       Strings.Left(.SP.bEntry.Text, 4) & ":" &
            '      R.SP.sMaxDH.Text & "/" & Format$(R.SS.Main.DH.avg, "0.0") & ":" &
            '     Strings.Left(.SP.sEntry.Text, 5) & "==" & StrSign(Format$(R.SS.Main.Profits.avg, "00.000")) &
            '    "','" & .SS.Main.Q.avg &
            '   "','" & R.SS.Main.absQuantum &
            '  "','" & .SP.Buy.Text &
            ' "','" & Now() & "')"
            'OleDBC.ExecuteNonQuery()
        End With
    End Sub
    Public Sub Put_Statistics02(ByRef R As Results)
        '        Dim OleDBC As New OleDbCommand
        '        Dim conn0 As New System.Data.OleDb.OleDbConnection
        With R
            '      OleDBC.Connection = conn0
            '     OleDBC.CommandText = "Insert Into Statistics02 VALUES ('" & Counters.totalIterations &
            '    "','" & Counters.end_day &
            '   "','" & lastTrade.bentry.qHiti &
            '  "','" & lastTrade.bSignal.qHiti &
            ' "','" & lastTrade.bDOW.qHiti &
            '           "','" & lastTrade.bZscore.qHiti &
            '          "','" & lastTrade.bTD1.qHiti &
            '         "','" & lastTrade.bTD2.qHiti &
            '        "','" & lastTrade.bTD3.qHiti &
            '       "','" & lastTrade.bTD4.qHiti &
            '      "','" & lastTrade.bTD5.qHiti &
            '     "','" & R.TFSignal &
            '    "','" & Now() & "')"
            '   OleDBC.ExecuteNonQuery()
        End With
    End Sub
    Public Sub Put_distribution(ByRef RR As Results, ByRef TS() As Trades)
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
            '            OleDBC.Connection = conn0
            '           OleDBC.CommandText = "Insert Into dist VALUES ('" & Counters.Iteration &
            '          "','" & .sSymbol &
            '         "','" & .lSymbol &
            '         "','" & .SP.bDOW.Text &
            '      "','" & RR.SS.Main.Q.avg &
            '      "','" & .dist(1) &
            '     "','" & .dist(2) &
            '    "','" & .dist(3) &
            '   "','" & .dist(4) &
            '  "','" & .dist(5) &
            ' "','" & .dist(6) &
            '"','" & .dist(7) &
            '           "','" & .dist(8) &
            '          "','" & .dist(9) &
            '         "','" & .dist(10) &
            '        "','" & .dist(11) &
            '       "','" & .dist(12) &
            '      "','" & .dist(13) &
            '     "','" & .dist(14) &
            '    "','" & .dist(15) &
            '   "','" & .dist(16) &
            '  "','" & .dist(17) &
            ' "','" & .dist(18) &
            '"','" & .dist(19) &
            '            "','" & .dist(20) &
            '           "','" & .dist(21) &
            '          "','" & .dist(22) &
            '         "','" & .dist(23) &
            '        "','" & .dist(24) &
            '       "','" & .dist(25) &
            '      "','" & .dist(26) &
            '     "','" & .dist(27) &
            '    "','" & .dist(28) &
            '   "','" & .dist(29) &
            '  "','" & .dist(30) &
            ' "','" & .dist(31) &
            '"','" & .dist(32) &
            '            "','" & .dist(33) &
            '           "','" & .dist(34) &
            '          "','" & .dist(35) &
            '         "','" & .dist(36) &
            '        "','" & .dist(37) &
            '       "','" & .dist(38) &
            '      "','" & .dist(39) &
            '     "','" & .dist(40) &
            '    "','" & .dist(41) &
            '   "','" & .dist(42) &
            '  "','" & .dist(43) &
            ' "','" & .dist(44) &
            '"','" & .dist(45) &
            '            "','" & .dist(46) &
            '           "','" & .dist(47) &
            '          "','" & .dist(48) &
            '         "','" & .dist(49) &
            '        "','" & .dist(50) &
            '       "','" & .dist(51) &
            '      "','" & .dist(52) &
            '     "','" & .dist(53) &
            '    "','" & .dist(54) &
            '   "','" & .dist(55) &
            '  "','" & .dist(56) &
            ' "','" & .dist(57) &
            '"','" & .dist(58) &
            '            "','" & .dist(59) &
            '           "','" & .dist(60) &
            '          "','" & .dist(61) &
            '         "','" & .SS.Main.Trades &
            ''        "','" & Now() & "')"
            '           OleDBC.ExecuteNonQuery()
        End With
    End Sub
    '
    '
    '
    'Call ConnectDatabase()
    '        For Each foundFile As String In My.Computer.FileSystem.GetFiles(
    '  "c:\temp\")
    '     xxxx = xxxx + 1
    '    securitiesListBox0.Items.Add(foundFile)
    ' Next
    '    xxxx = -1
    '        For Each foundFile As String In My.Computer.FileSystem.GetFiles(
    '"\\WDMYCLOUD\Public\Gregory's data\QFE_Data\SectorX")
    '       xxxx = xxxx + 1
    '      securitiesListBox1.Items.Add(foundFile)
    '     Next
    '       For Each foundFile As String In My.Computer.FileSystem.GetFiles("C:\Users\Dad\OneDrive\QFE_Prices\")
    '  For Each foundFile As String In My.Computer.FileSystem.GetFiles("c:\temp")
    '      xxxx = xxxx + 1
    '       XSectorSecurities.Items.Add(foundFile)
    '    Next
    '     Me.Message.Text = "reading Parameters"
    '      Call Rdosecurities()
    '       Call Rdoparams1()
    'Call Rdoparams2()
    ' Call Rdoparams3()
    '  Call Rdoparams4()
    '   Call Rdoparams5()
    ''        Call write_Parameters1()
    '  My.Application.DoEvents()
    '   Me.Message.Text = "deleting Tables"
    ''       Call delete_Tables()
    '     Me.Message.Text = "tables Deleted"
    '      start_Seconds = DateDiff(DateInterval.Second, Now.Date, Now)
    '       start_Time = Date.Now
    'My.Application.DoEvents()
    ' Call Setlistview1cols()
    '  FileSystem.FileClose(1)
    'endd:'
    'Me.quantumThreshHold.Value = 0.75
    ' Me.quantumThreshholdtxt1.Text = Me.quantumThreshHold.Value
    '  Counters.threshHold = Me.quantumThreshHold.Value
    '   My.Application.DoEvents()
    'End Sub
    Public Sub DoBHTrades_OpnCl(ByRef R As Results, ByRef TS() As Trades)
        Static xline As Integer
        Call InitR(R)
        With R
            R.SP.bDOW.Text = "b_h"
            .SP.BandH.Text = "bhOpnCl"
            .SP.bMaxDH.V = 1
            .SP.bEntry.Text = "bh__Op"
            .SP.bSignal.Text = "bh__Op"
            .SP.bEntry.Text = "bh__Op"
            .SP.sEntry.Text = "bh_nCl"
            .SP.sSignal.Text = "bh_nCl"
            .SP.sEntry.Text = "bh_nCl"
            .SP.bZScoreMode.Text = "0Z--"
            For xline = Counters.start_Day To Counters.end_day
                .SS.Main.Trades = .SS.Main.Trades + 1
                TS(.SS.Main.Trades).TradeNo = .SS.Main.Trades
                TS(.SS.Main.Trades).bDOW.qHit = True
                TS(.SS.Main.Trades).bSignal.qHit = True
                TS(.SS.Main.Trades).bEntry.qHit = True
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
                .SS.Main.bEntryStats.Hits = .SS.Main.bEntryStats.Hits + 1
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
    Private Sub DoBHTrades_ClnCl(ByRef R As Results, ByRef TS() As Trades)
        Static xline As Integer
        Call InitR(R)
        With R.SP
            .BandH.Text = "bhClnCl"
            .bEntry.Text = "bh__Cl"
            .bDOW.Idx = 0
            .bDOW.Text = "b_h"
            .bSignal.Idx = 1
            .bSignal.Text = "bh_Cl"
            .bEntry.Text = "bh_Cl"
            .sSignal.Idx = 1
            .sSignal.Text = "bh_nCl_"
            .sEntry.Text = "bh_nCl_"
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
                .bEntry.qHit = True
                .bEntry.qHit = True
                '                .BExeMiss.qHit = False
                .BuyOnLow.qHit = False
                .BuyOnNextClose.qHit = False
                '               .BExeTg.qHit = False
            End With
        Next xline
        lastTrade = TS(R.SS.Main.Trades)
        lastTrade.bEntry.qHit = True
        lastTrade.bDOW.qHit = True
        lastTrade.bSignal.qHit = True
    End Sub
    Public Sub DobhTrades(ByRef R As Results, ByRef TS() As Trades)
        If QFE1.qDoBHOptoCl.Checked Then
            QFE1.DoBHStatus.ForeColor = Color.Azure
            QFE1.DoBHStatus.Text = Format$(0.0, "000.000")
            Call DoBHTrades_Op_Cl(R, TS)
            Call Saves(R, TS)
            Call Calculate_BasicStatistics(R, TS, 5, R.SS.Main.Trades - 5)
            Call Put_Signals(R, TS, 2)
        End If
        '  If Me.qDoBHOptoNOp.Checked Then
        '           Call DoBHTrades_OpnOp(R)
        '           Call Calculate_BasicStatistics(R, 1, R.SS.Main.Trades)
        '           Call Put_Signals_bh(R, 2)
        '  Call Saves(R)
        'End If
        '       If Me.qdoBHOptoNCl.Checked Then
        '           Call DoBHTrades_OpnCl(R)
        '          Call Calculate_BasicStatistics(R, 1, R.SS.Main.Trades)
        '           Call Put_Signals_bh(R, 2)
        '     Call Saves(R)
        '      End If
        '     If Me.qDoBHCltoNOp.Checked Then
        '         Call DoBHTrades_ClnOp(R)
        '        Call Calculate_BasicStatistics(R, 1, R.SS.Main.Trades)
        '           Call Put_Signals_bh(R, 2)
        '     Call Saves(R)
        '    End If
        '       If Me.qDoBHCltoNCl.Checked Then
        '       Call DoBHTrades_ClnCl(R)
        '       Call Calculate_BasicStatistics(R, 1, 0)
        '       Call Put_Signals_bh(R, 2)
        '      Call Saves(R)
        '      End If
        '   If Me.qDoBandHold.Checked Then
        Counters.qBuyandHold.qHit = True
        'Call DoBuyandHold(R)
        '        End If
    End Sub
    Public Sub DoBHTrades_Op_Cl(ByRef R As Results, ByRef TS() As Trades)
        Dim xline As Integer, totProfit As Single, totDH As Integer
        Dim maxProfit As Single, tr As Integer
        ReDim TS(0)
        ReDim TS(xxx)
        Call InitR(R)
        maxProfit = Val(QFE1.maxProfit.Text) * 10.0
        With R.SP
            .BandH.Text = "bhOp_Cl"
            .bEntry.Text = "bh__Op"
            .bDOW.Idx = 0
            .bDOW.Text = "b_h"
            .bSignal.Idx = 1
            .bSignal.Text = "bh__Op"
            .bEntry.Text = "bh__Op"
            .sSignal.Idx = 1
            .sSignal.Text = "bh__Cl"
            .sEntry.Text = "bh__Cl"
            .sEntry.Text = .sSignal.Text
            .bZScoreMode.Text = "0Z--"
            .bMaxDH.Text = "01"
        End With
        totProfit = 0.0
        totDH = 0
        tr = 0
        For xline = Counters.start_Day To Counters.end_day - 5
            R.SS.Main.Trades = R.SS.Main.Trades + 1
            tr = tr + 1
            With TS(tr)
                .TradeNo = tr
                .bDOW.Text = R.SP.bDOW.Text
                TS(tr).bDayNo = xline
                TS(tr).sDayNo = xline
                TS(tr).bDate = dtStr2(TS(tr).bDayNo)
                TS(tr).sEntry.Text = dtStr2(TS(tr).sDayNo)
                TS(tr).bEntry.Text = dtStr2(TS(tr).bDayNo)
                TS(tr).sDate = dtStr2(TS(tr).sDayNo)
                .DH = .sDayNo - .bDayNo + 1
                totDH = totDH + 1
                .totDH = totDH
                .bAmt = 1000.0
                TS(R.SS.Main.Trades).bPrice = op(xline)
                If .bPrice = Single.NaN Then Stop
                If .bPrice < 0 Then Stop
                TS(tr).sPrice = cl(.sDayNo)
                If .sPrice = Single.NaN Then Stop
                If .sPrice < 0 Then Stop
                'If .sPrice > 2000 Then Stop
                ' If .sPrice < -2000 Then Stop
                .Shares = .bAmt / .bPrice
                .sAmt = .Shares * .sPrice
                TS(R.SS.Main.Trades).Profit0 = .sAmt - .bAmt
                If Math.Abs(TS(tr).Profit0) > Val(maxProfit) Then
                    .Profit0 = 0.0
                    '    Stop
                End If
                If Single.IsNaN(TS(tr).Profit0) Then Stop
                totProfit = totProfit + .Profit0
                .Quantum = .Profit0 / .DH
                .totProfit = totProfit
                .totQuantum = totProfit / .totDH / 10
                .maxDH = 1
                .BuyOnOpen.qHit = True
                .BuyOnClose.qHit = False
                .bEntry.qHit = True
                '                .BExeMiss.qHit = False
                .BuyOnLow.qHit = False
                .BuyOnNextClose.qHit = False
                '                .BExeTg.qHit = False
            End With
        Next xline
        R.SS.Main.Trades = tr
        lastTrade = TS(tr)
        lastTrade.bEntry.qHit = True
        lastTrade.bDOW.qHit = True
        lastTrade.bSignal.qHit = True
    End Sub
    Public Sub DoBHTrades_OpnOp(ByRef R As Results, ByRef TS() As Trades)
        Static xline As Integer
        Call InitR(R)
        With R.SP
            .BandH.Text = "bhOpnOp"
            .bEntry.Text = "bh__Op"
            .bDOW.Idx = 0
            .bDOW.Text = "b_h"
            .bSignal.Idx = 1
            .bSignal.Text = "bh___Op"
            .bEntry.Text = "bh___Op"
            .sSignal.Idx = 1
            .sSignal.Text = "bh_nOp"
            .sEntry.Text = "bh_nOp"
            .sEntry.Text = .bEntry.Text
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
                .bEntry.qHit = True
                .bEntry.qHit = True
                '                .BExeMiss.qHit = False
                .BuyOnLow.qHit = False
                .BuyOnNextClose.qHit = False
                .BuyOnNextOpen.qHit = False
                '                .BExeTg.qHit = False
            End With
        Next xline
        lastTrade = TS(R.SS.Main.Trades)
        lastTrade.bEntry.qHit = True
        lastTrade.bDOW.qHit = True
        lastTrade.bSignal.qHit = True
    End Sub
    Private Sub DoBuyandHold(ByRef R As Results, ByRef TS() As Trades)
        Static incr As Integer, ddayt As Integer, dDay As Integer, totPr As Single, totDH As Integer
        With R
            .SP.BandH.Text = Counters.qBuyandHold.Text
            .SP.bDOW.Text = Counters.qBuyandHold.Text
            .SP.sDOW.Text = Counters.qBuyandHold.Text
            .SP.bSignal.Text = Counters.qBuyandHold.Text
            .SP.bEntry.Text = Counters.qBuyandHold.Text
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
            '       .SP.bZScoreStd.Text = Counters.qBuyandHold.Text
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
            lastDayTrade = TS(dDay - 1)
            lastDayTrade.bTD0.Text = "bh"
            lastDayTrade.bTD1.Text = "bh"
            lastDayTrade.bTD2.Text = "bh"
            lastDayTrade.bDOW.Text = "bh"
            lastDayTrade.bEntry.Text = "bh"
            lastDayTrade.bEntry.Text = "bh"
            '           Counters.Iteration = Counters.Iteration + 1
            '           Call Calculate_BasicStatistics(R, 2, R.SS.Main.Trades - 1)
            Call Put_Signals(R, TS, 2)
            '            Call saves(R)
        End With
    End Sub
    Private Sub DoDaytrades(ByRef R As Results, ByRef TS() As Trades)
        R.SP.BandH.Text = "BH---Mon"
        R.SP.bDOW.Text = "1Mon"
        R.SP.bSignal.Text = "0Op"
        R.SP.bEntry.Text = "0Op"
        R.SP.sSignal.Text = "0CL"
        R.SP.sEntry.Text = "0CL"
        lastTrade.bDOW.qHit = (dowNo(Counters.end_day) = 1)
        lastTrade.bEntry.qHit = lastTrade.bDOW.qHit
        lastTrade.bSignal.qHit = lastTrade.bDOW.qHit
        '        ReDim R.qBuyOnLastDay(Counters.end_day)
        '       ReDim R.qBuyOnLastDayI(Counters.end_day)
        Call DoTradesMon(R, TS)
        '        Call Calculate_BasicStatistics(R, 1, 0)
        '       Call Write_Parameters0(R, TS)
        R.SP.BandH.Text = "BH---Tue"
        R.SP.bDOW.Text = "2Tue"
        R.SP.bSignal.Text = "0Op"
        R.SP.bEntry.Text = "0Op"
        R.SP.sSignal.Text = "0CL"
        R.SP.sEntry.Text = "bhOp"
        lastTrade.bDOW.qHit = (dowNo(Counters.end_day) = 2)
        lastTrade.bEntry.qHit = lastTrade.bDOW.qHit
        lastTrade.bSignal.qHit = lastTrade.bDOW.qHit
        Call DoTradesTue(R, TS)
        '       Call Calculate_BasicStatistics(R, 1, 0)
        '        r.qBuyOnLastDay(Counters.end_day) = lastTrade.bDOW.qHit
        '      Call Write_Parameters0(R, TS)
        R.SP.BandH.Text = "BH---Wed"
        R.SP.bDOW.Text = "3Wed"
        R.SP.bSignal.Text = "0Op"
        R.SP.bEntry.Text = "0Op"
        R.SP.sSignal.Text = "0CL"
        R.SP.sEntry.Text = "0CL"
        lastTrade.bDOW.qHit = (dowNo(Counters.end_day) = 3)
        lastTrade.bEntry.qHit = lastTrade.bDOW.qHit
        lastTrade.bSignal.qHit = lastTrade.bDOW.qHit
        Call DoTradesWed(R, TS)
        '      Call Calculate_BasicStatistics(R, 1, 0)
        '        r.qBuyOnLastDay(Counters.end_day) = lastTrade.bDOW.qHit
        '       Call Write_Parameters0(R, TS)
        R.SP.BandH.Text = "BH---Thr"
        R.SP.bDOW.Text = "4Thr"
        R.SP.bSignal.Text = "0Op"
        '        r.SP.Text = "0Op"
        R.SP.sSignal.Text = "0CL"
        R.SP.sEntry.Text = "0CL"
        lastTrade.bDOW.qHit = (dowNo(Counters.end_day) = 4)
        lastTrade.bEntry.qHit = lastTrade.bDOW.qHit
        lastTrade.bSignal.qHit = lastTrade.bDOW.qHit
        Call DoTradesThr(R, TS)
        '     Call Calculate_BasicStatistics(R, 1, 0)
        '        r.qBuyOnLastDay(Counters.end_day) = lastTrade.bDOW.qHit
        '      Call Write_Parameters0(R, TS)
        R.SP.BandH.Text = "BH---Fri"
        R.SP.bDOW.Text = "5Fri"
        R.SP.bSignal.Text = "0Op"
        R.SP.bEntry.Text = "0Op"
        R.SP.sSignal.Text = "0CL"
        R.SP.sEntry.Text = "0CL"
        lastTrade.bDOW.qHit = (dowNo(Counters.end_day) = 5)
        lastTrade.bEntry.qHit = lastTrade.bDOW.qHit
        lastTrade.bSignal.qHit = lastTrade.bDOW.qHit
        Call DoTradesFri(R, TS)
        '    Call Calculate_BasicStatistics(R, 1, 0)
        '       r.qBuyOnLastDay(Counters.end_day) = lastTrade.bDOW.qHit
        '    Call Write_Parameters0(R, TS)
    End Sub
    Public Sub Rdoparams1()
        Static secno As Integer, xx As Integer
        QFE1.buyDowTxt.Text = "c\qfe:\parameterslist1.txt"
        FileSystem.FileOpen(1, "c:\qfe\parameterslist1.txt", OpenMode.Input, OpenAccess.Read)
        secno = -1
        Do While Not EOF(1) And secno < QFE1.buyDOW.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.buyDOW.SetItemChecked(secno, True)
            Else
                QFE1.buyDOW.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.buyTrigger.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.buyTrigger.SetItemChecked(secno, True)
            Else
                QFE1.buyTrigger.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.sellEntry.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.sellEntry.SetItemChecked(secno, True)
            Else
                QFE1.sellEntry.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.sellMaxDH.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.sellMaxDH.SetItemChecked(secno, True)
            Else
                QFE1.sellMaxDH.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.whichDates.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.whichDates.SetItemChecked(secno, True)
            Else
                QFE1.whichDates.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.buyMonth.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.buyMonth.SetItemChecked(secno, True)
            Else
                QFE1.buyMonth.SetItemChecked(secno, False)
            End If
        Loop
        FileSystem.Input(1, xx)
        FileSystem.FileClose(1)
        Application.DoEvents()
    End Sub
    Public Sub Rdoparams2()
        Static xx As Integer
        QFE1.dbTextBox4.Text = "c:\qfe\parameterslist2.txt"
        FileSystem.FileOpen(1, "c:\qfe\parameterslist2.txt", OpenMode.Input, OpenAccess.Read)
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qSaveTrades0.Checked = True
        Else
            QFE1.qSaveTrades0.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qSaveTrades1.Checked = True
        Else
            QFE1.qSaveTrades1.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qSaveTrades1a.Checked = True
        Else
            QFE1.qSaveTrades1a.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qSaveTrades2.Checked = True
        Else
            QFE1.qSaveTrades2.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qSaveTrades3.Checked = True
        Else
            QFE1.qSaveTrades3.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qSaveTrades4.Checked = True
        Else
            QFE1.qSaveTrades4.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qSaveTrades5.Checked = True
        Else
            QFE1.qSaveTrades5.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qSaveParams0.Checked = True
        Else
            QFE1.qSaveParams0.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qSaveParams1.Checked = True
        Else
            QFE1.qSaveParams1.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qSaveParams2.Checked = True
        Else
            QFE1.qSaveParams2.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qSaveParams3.Checked = True
        Else
            QFE1.qSaveParams3.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qSaveStats0.Checked = True
        Else
            QFE1.qSaveStats0.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qSaveDist.Checked = True
        Else
            QFE1.qSaveDist.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qSaveOnSignals.Checked = True
        Else
            QFE1.qSaveOnSignals.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qSaveOffSignals2.Checked = True
        Else
            QFE1.qSaveOffSignals2.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qOnLastDay.Checked = True
        Else
            QFE1.qOnLastDay.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qDoBHOptoCl.Checked = True
        Else
            QFE1.qDoBHOptoCl.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qDoBHOptoNOp.Checked = True
        Else
            QFE1.qDoBHOptoNOp.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qDoBHCltoNOp.Checked = True
        Else
            QFE1.qDoBHCltoNOp.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qDoBHCltoNCl.Checked = True
        Else
            QFE1.qDoBHCltoNCl.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qdoBHOptoNCl.Checked = True
        Else
            QFE1.qdoBHOptoNCl.Checked = False
        End If
        FileSystem.Input(1, xx)
        If xx = 1 Then
            QFE1.qDoBandHold.Checked = True
        Else
            QFE1.qDoBandHold.Checked = False
        End If
endd:
        FileSystem.FileClose(1)
        Application.DoEvents()
    End Sub
    Public Sub Rdoparams3()
        Static secno As Integer, xx As Integer
        QFE1.dbTextBox5.Text = "c:\qfe\parameterslist3.txt"
        FileSystem.FileOpen(1, "c:\qfe\parameterslist3.txt", OpenMode.Input, OpenAccess.Read)
        secno = -1
        Do While secno < QFE1.buyEntry.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.buyEntry.SetItemChecked(secno, True)
            Else
                QFE1.buyEntry.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.BuyInDay.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.BuyInDay.SetItemChecked(secno, True)
            Else
                QFE1.BuyInDay.SetItemChecked(secno, False)
            End If
        Loop
        FileSystem.FileClose(1)
        Application.DoEvents()
    End Sub
    Public Sub Rdoparams4()
        Static secno As Integer, xx As Integer
        'dbTextBox3.Text = "c\qfe:\parameterslist4.txt"
        FileSystem.FileOpen(1, "c:\qfe\parameterslist4.txt", OpenMode.Input, OpenAccess.Read)
        secno = -1
        Do While Not EOF(1) And secno < QFE1.buyZScoreMode.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.buyZScoreMode.SetItemChecked(secno, True)
            Else
                QFE1.buyZScoreMode.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.buyZScoreValue.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.buyZScoreValue.SetItemChecked(secno, True)
            Else
                QFE1.buyZScoreValue.SetItemChecked(secno, False)
            End If
        Loop
        FileSystem.FileClose(1)
        Application.DoEvents()
    End Sub
    Public Sub Rdoparams5()
        Static secno As Integer, xx As Integer
        QFE1.TxtParameters05.Text = "c\qfe:\parameterslist5.txt"
        FileSystem.FileOpen(1, "c:\qfe\parameterslist5.txt", OpenMode.Input, OpenAccess.Read)
        secno = -1
        Do While Not EOF(1) And secno < QFE1.buyTDSignal1.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.buyTDSignal1.SetItemChecked(secno, True)
            Else
                QFE1.buyTDSignal1.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.buyTDSignal2.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.buyTDSignal2.SetItemChecked(secno, True)
            Else
                QFE1.buyTDSignal2.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.buyTDSignal3.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.buyTDSignal3.SetItemChecked(secno, True)
            Else
                QFE1.buyTDSignal3.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.buyTDSignal4.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.buyTDSignal4.SetItemChecked(secno, True)
            Else
                QFE1.buyTDSignal4.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.buyTDSignal5.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.buyTDSignal5.SetItemChecked(secno, True)
            Else
                QFE1.buyTDSignal5.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.SellTDSignal1.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.SellTDSignal1.SetItemChecked(secno, True)
            Else
                QFE1.SellTDSignal1.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.buyTDDaysBack1.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.buyTDDaysBack1.SetItemChecked(secno, True)
            Else
                QFE1.buyTDDaysBack1.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.buyTDDaysBack2.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.buyTDDaysBack2.SetItemChecked(secno, True)
            Else
                QFE1.buyTDDaysBack2.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.buyTDDaysBack3.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.buyTDDaysBack3.SetItemChecked(secno, True)
            Else
                QFE1.buyTDDaysBack3.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.buyTDDaysBack4.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.buyTDDaysBack4.SetItemChecked(secno, True)
            Else
                QFE1.buyTDDaysBack4.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.buyTDDaysBack5.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.buyTDDaysBack5.SetItemChecked(secno, True)
            Else
                QFE1.buyTDDaysBack5.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.SellTDDaysBack1.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.SellTDDaysBack1.SetItemChecked(secno, True)
            Else
                QFE1.SellTDDaysBack1.SetItemChecked(secno, False)
            End If
        Loop
        secno = -1
        Do While Not EOF(1) And secno < QFE1.targetProfit.Items.Count - 1
            secno = secno + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                QFE1.targetProfit.SetItemChecked(secno, True)
            Else
                QFE1.targetProfit.SetItemChecked(secno, False)
            End If
        Loop
        ' FileSystem.Input(1, xx)
        ' If xx = 1 Then
        ' QFE1.qCalcSigString.Checked = True
        ' Else
        ' QFE1.qCalcSigString.Checked = False
        ' End If
        'FileSystem.Input(1, xx)
        Application.DoEvents()
        FileSystem.FileClose(1)
    End Sub
    Public Sub Delete_Tables1()
        'Dim OleDBC As New OleDbCommand
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
        '           .CommandText = "DELETE FROM Parameters01p
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
        '      With OleDBC
        '    .Connection = conn1
        '         .CommandText = "DELETE FROM Trades00"
        '           .ExecuteNonQuery()
        '   End With
        '  With OleDBC
        '       .Connection = conn2
        '      .CommandText = "DELETE FROM Signal_"
        '     .ExecuteNonQuery()
        '    End With
        '        With OleDBC
        '       .Connection = conn2
        '      .CommandText = "DELETE FROM Signals"
        '     .ExecuteNonQuery()
        '    End With
        '   With OleDBC
        '  .Connection = conn2
        ' .CommandText = "DELETE FROM Base"
        '.ExecuteNonQuery()
        '        End With
        '       With OleDBC
        '      .Connection = conn2
        '     .CommandText = "DELETE FROM lastDay"
        '    .ExecuteNonQuery()
        '   End With
        '  '  End With
        MsgBox("Record Deleted!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "SUCCESS")
    End Sub
    Public Sub Write_AllParameters()
        Call Wdosecurities()
        Call Wdoparams01()
        Call Wdoparams02()
        Call Wdoparams03()
        Call Wdoparams04()
        Call Wdoparams05()
        My.Application.DoEvents()
    End Sub
    Private Sub Wdosecurities()
        Static secno As Integer
        FileSystem.FileClose(1)
        My.Computer.FileSystem.DeleteFile("C:\Qfe\securitieslist1.txt")
        FileSystem.FileOpen(1, "C:\Qfe\securitiesList1.txt", OpenMode.Append, OpenAccess.ReadWrite)
        If QFE1.qBXSector.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        For secno = 0 To QFE1.XSectorSecurities.Items.Count - 1
            If QFE1.XSectorSecurities.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        FileSystem.FileClose(1)
        '       My.Computer.FileSystem.DeleteFile("C:\Qfe\securitieslist1.txt")
        '       FileSystem.FileOpen(1, "C:\Users\Dad\Documents\securitiesList1.txt", OpenMode.Append, OpenAccess.ReadWrite)
        '       For secno = 0 To XSectorSecurities.Items.Count - 1
        '       If XSectorSecurities.GetItemChecked(secno) Then
        '       FileSystem.Write(1, 1)
        '       Else
        '      FileSystem.Write(1, 0)
        '      End If
        '     Next secno
        FileSystem.FileClose(1)
    End Sub
    Private Sub Wdoparams01()
        FileSystem.FileClose(1)
        '      Static secno As Integer
        My.Computer.FileSystem.DeleteFile("c:\qfe\parameterslist1.txt")
        FileSystem.FileOpen(1, "c:\qfe\parameterslist1.txt", OpenMode.Append, OpenAccess.ReadWrite)
        For secno = 0 To QFE1.buyDOW.Items.Count - 1
            If QFE1.buyDOW.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.buyTrigger.Items.Count - 1
            If QFE1.buyTrigger.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.sellEntry.Items.Count - 1
            If QFE1.sellEntry.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.sellMaxDH.Items.Count - 1
            If QFE1.sellMaxDH.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.whichDates.Items.Count - 1
            If QFE1.whichDates.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.buyMonth.Items.Count - 1
            If QFE1.buyMonth.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        FileSystem.Write(1, QFE1.quantumThreshHold.TabIndex)
        FileSystem.FileClose(1)
        Application.DoEvents()
    End Sub
    Private Sub Wdoparams02()
        My.Computer.FileSystem.DeleteFile("c:\qfe\parameterslist2.txt")
        FileSystem.FileOpen(1, "c:\qfe\parameterslist2.txt", OpenMode.Append, OpenAccess.ReadWrite)
        If QFE1.qSaveTrades0.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qSaveTrades1.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qSaveTrades1a.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qSaveTrades2.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qSaveTrades3.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qSaveTrades4.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qSaveTrades5.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qSaveParams0.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qSaveParams1.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qSaveParams2.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qSaveParams3.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qSaveStats0.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qSaveDist.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qSaveOnSignals.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qSaveOffSignals2.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qOnLastDay.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qDoBHOptoCl.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qDoBHOptoNOp.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qDoBHCltoNOp.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qDoBHCltoNCl.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qdoBHOptoNCl.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        If QFE1.qDoBandHold.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        FileSystem.FileClose(1)
        Application.DoEvents()
    End Sub
    Private Sub Wdoparams03()
        Static secno As Integer
        My.Computer.FileSystem.DeleteFile("c:\qfe\parameterslist3.txt")
        FileSystem.FileOpen(1, "c:\qfe\parameterslist3.txt", OpenMode.Append, OpenAccess.ReadWrite)
        For secno = 0 To QFE1.buyEntry.Items.Count - 1
            If QFE1.buyEntry.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.BuyInDay.Items.Count - 1
            If QFE1.BuyInDay.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        QFE1.TxtParameters03.BackColor = SystemColors.MenuHighlight
        FileSystem.FileClose(1)
        Application.DoEvents()
    End Sub
    Private Sub Wdoparams04()
        FileSystem.FileClose(1)
        '        Static secno As Integer
        My.Computer.FileSystem.DeleteFile("c:\qfe\parameterslist4.txt")
        FileSystem.FileOpen(1, "c:\qfe\parameterslist4.txt", OpenMode.Append, OpenAccess.ReadWrite)
        For secno = 0 To QFE1.buyZScoreMode.Items.Count - 1
            If QFE1.buyZScoreMode.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.buyZScoreValue.Items.Count - 1
            If QFE1.buyZScoreValue.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        '   Form1.TxtParameters04.BackColor = SystemColors.MenuHighlight
        FileSystem.FileClose(1)
        Application.DoEvents()
    End Sub
    Private Sub Wdoparams05()
        FileSystem.FileClose(1)
        Static secno As Integer
        My.Computer.FileSystem.DeleteFile("c:\qfe\parameterslist5.txt")
        FileSystem.FileOpen(1, "c:\qfe\parameterslist5.txt", OpenMode.Append, OpenAccess.ReadWrite)
        For secno = 0 To QFE1.buyTDSignal1.Items.Count - 1
            If QFE1.buyTDSignal1.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.buyTDSignal2.Items.Count - 1
            If QFE1.buyTDSignal2.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.buyTDSignal3.Items.Count - 1
            If QFE1.buyTDSignal3.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.buyTDSignal4.Items.Count - 1
            If QFE1.buyTDSignal4.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.buyTDSignal5.Items.Count - 1
            If QFE1.buyTDSignal5.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.SellTDSignal1.Items.Count - 1
            If QFE1.SellTDSignal1.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.buyTDDaysBack1.Items.Count - 1
            If QFE1.buyTDDaysBack1.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.buyTDDaysBack2.Items.Count - 1
            If QFE1.buyTDDaysBack2.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.buyTDDaysBack3.Items.Count - 1
            If QFE1.buyTDDaysBack3.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.buyTDDaysBack4.Items.Count - 1
            If QFE1.buyTDDaysBack4.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.buyTDDaysBack5.Items.Count - 1
            If QFE1.buyTDDaysBack5.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.SellTDDaysBack1.Items.Count - 1
            If QFE1.SellTDDaysBack1.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        For secno = 0 To QFE1.targetProfit.Items.Count - 1
            If QFE1.targetProfit.GetItemChecked(secno) Then
                FileSystem.Write(1, 1)
            Else
                FileSystem.Write(1, 0)
            End If
        Next secno
        If QFE1.qCalcSigString.Checked Then
            FileSystem.Write(1, 1)
        Else
            FileSystem.Write(1, 0)
        End If
        FileSystem.FileClose(1)
        QFE1.TxtParameters05.BackColor = SystemColors.MenuHighlight
        Application.DoEvents()
    End Sub
    Public Sub RunIt(ByRef R As Results)
        '        Me.ThreshHoldTrades_.Text = Me.ThreshHoldTrades_.SelectedItem
        '    threshHoldTrades = Val(Me.ThreshHoldTrades_.Text)
        '       Me.ThreshholdQuantum_.SelectedIndex = 0
        '      Me.ThreshholdQuantum_.Text = Me.ThreshholdQuantum_.SelectedItem
        '     threshHoldQuantum = Val(Me.ThreshholdQuantum_.Text)
        Counters.Iteration = 0
        Counters.SystemNumber = 0
        sysBuyOnLastDay0 = 0
        sysBuyOnLastDay1 = 0
        sysBuyOnLastDay2 = 0
        sysBuyOnLastDay3 = 0
        sysBuyOnLastDay4 = 0
        sysBuyOnLastDay5 = 0
        Counters.SystemNumber = 0
        Counters.signalsSaved = 0
        Counters.totTradeNo = 0
        Counters.Incr_ = 0
        Call DoAllSecuritiesLoop(R)
        '        Call doSecSectorX()
        conn1.Close()
    End Sub
    Private Sub DoAllSecuritiesLoop(ByRef R As Results)
        Static rowss As Integer, spacePlace As Integer
        Dim fn As String, xx As Integer, tmpLng As Integer
        QFE1.SignalsGrid.GridColor = Color.Bisque
        QFE1.msg.Text = "processing Files"
        QFE1.msg.BackColor = Color.Bisque
        rowss = 0
        QFE1.SignalsGrid.RowCount = 100
        QFE1.SignalsGrid.ColumnCount = 500
        QFE1.Security_.BackColor = Color.AliceBlue
        For Counters.SecNo = 0 To QFE1.XSectorSecurities.Items.Count - 1
            QFE1.XSectorSecurities.SelectedIndex = Counters.SecNo
            If QFE1.XSectorSecurities.GetItemChecked(Counters.SecNo) Then
                Counters.FileString = QFE1.XSectorSecurities.Items(Counters.SecNo).ToString
                tmpLng = Strings.Len(Counters.FileString)
                xx = 3
fnloop:
                If Counters.FileString(xx) = "\" Then
                    spacePlace = xx
                Else
                    xx = xx + 1
                    If xx >= tmpLng Then Stop
                    GoTo fnLoop
                End If
                xx = xx + 1
                fn = ""
fnloop1:
                If Counters.FileString(xx) = " " Then GoTo exitfnloop
                fn = fn & Counters.FileString(xx)
                xx = xx + 1
                If xx >= tmpLng Then Stop
                GoTo fnloop1
exitfnloop:
                ' Counters.currentSecurity = Strings.Left(fn, Strings.InStr(fn, " ") - 1)
                fn = Strings.Left(fn & "_____", 5)
                R.sSymbol = fn
                QFE1.SignalsGrid.Rows.Item(rowss).Cells(0).Value = R.sSymbol
                QFE1.stripSec.Text = R.sSymbol
                QFE1.Security_.Text = R.sSymbol
                QFE1.txtSymbol.Text = fn
                Call Process_Files(Counters.FileString, Counters.SecNo, 1, R)
                QFE1.DataGridViewSecurities.Rows.Item(rowss).Cells(0).Value = Counters.currentSecurity
                QFE1.DataGridViewSecurities.Rows.Item(rowss).Cells(1).Value = dtStr1(1)
                QFE1.DataGridViewSecurities.Rows.Item(rowss).Cells(2).Value = dtStr1(Counters.last_Day)
                QFE1.DataGridViewSecurities.Rows.Item(rowss).Cells(3).Value = Format(op(Counters.last_Day), "000.000")
                QFE1.DataGridViewSecurities.Rows.Item(rowss).Cells(4).Value = Format(hi(Counters.last_Day), "000.000")
                QFE1.DataGridViewSecurities.Rows.Item(rowss).Cells(5).Value = Format(lo(Counters.last_Day), "000.000")
                QFE1.DataGridViewSecurities.Rows.Item(rowss).Cells(6).Value = Format(cl(Counters.last_Day), "000.000")
                '                Me.endDate.Text = Counters.end_Date & DOWStr(Counters.end_day)
                '   Me.startDate.Text = Counters.start_Date
                '  Me.days.Text = Counters.Days
                ' Me.startDay.Text = Counters.start_Day
                'Me.endDay.Text = Counters.end_day
                Counters.this_Day = Counters.end_day
                Counters.this_Date = Counters.end_Date
                Counters.elapsedSeconds = current_Seconds - start_Seconds
                Counters.iterationsPerSecond = Counters.totalIterations / Counters.elapsedSeconds
                '            Me.iterpersecond.Text = Format$(Counters.iterationsPerSecond, "0.00")
                'Me.elapsedSeconds_.Text = Format$(Counters.elapsedSeconds, "0")
                Counters.end_Date = dtStr2(Counters.end_day)
                Counters.end_DOW = Left(DOWStr(Counters.end_day), 3)
                Counters.Month = Left(monthStr(Counters.end_day), 3)
                Counters.Days = Counters.end_day - Counters.start_Day + 1
                Counters.SystemNumber = 0
                Counters.xDayIdx = -1
                Call Put_Base(R, 2)
                '            For tmpDay = .end_day To .last_Day - (Me.whichDates.Items.Count - 1) Step -1
                Counters.xDayIdx = Counters.xDayIdx + 1
                '                Me.whichDates.Items(.xDayIdx) = Format$(.xDayIdx, "000") & " " &
                ' Format$(tmpDay, "0000") & " " & R.sSymbol
                '            Next
                Application.DoEvents()
                Counters.qBuyandHold.Text = "bhNon"
                Counters.qBuyandHold.Text1 = "bhNon"
                Counters.qBuyandHold.qHit = False
                QFE1.startDateTxt.Text = dtStr2(Counters.start_Day)
                QFE1.lastDateText.Text = dtStr2(Counters.end_day)
                If QFE1.qDoBandHold.Checked Then
                    R.SP.bMaxDH.V = 1
                    '   Call DobhTrades(R, RMainTS)
                    '                If Me.qSaveTrades1a.Checked Then
                    '                              Call Me.Put_Trades1a(R, RMainTS.tradeSeries)
                End If
                Call Do_WhichDayLoop(R)
                '       End If
                'If XSectorSecurities.GetItemChecked(.SecNo) Then
                '            End If
                '                QFE1.DataGridViewSecurities.
                '               Dim column As DataGridViewColumn = dataGridView.Columns(0)
                '              column.Width = 120
                QFE1.DataGridViewSecurities.Rows.Item(Counters.SecNo + 1).Cells(7).Value = Format(Counters.totalIterations, "00000")
                rowss = rowss + 1
                '   QFE1.SignalsGrid.RowCount = QFE1.SignalsGrid.RowCount + 1
            End If
        Next Counters.SecNo
    End Sub

    Private Sub Do_WhichDayLoop(ByRef R As Results)
        With Counters
            .xDayIdx = -1
            For .last_day1 = .last_Day To .last_Day - 4 Step -1
                .xDayIdx = .xDayIdx + 1
                Counters.end_Date = dtStr1(Counters.last_day1)
                Counters.end_day = Counters.last_day1
                Counters.last_date1 = dtStr1(.last_day1)
                QFE1.whichDates.Items(.xDayIdx) = Format$(.xDayIdx, "000") & " " &
                      Format$(.last_day1, "0000") & " " & .last_Date & " " & .last_date1 & " " &
                         " op:" & Format$(op(.last_day1), "000.000") & " hi:" & Format$(hi(.last_day1), "000.000") &
                        " lo:" & Format$(lo(.last_day1), "000.000") & " cl:" & Format$(cl(.last_day1), "000.000") & " " &
                   Format$(Counters.Days, "00000") & " " & R.sSymbol
                If QFE1.whichDates.GetItemChecked(.xDayIdx) Then
                    QFE1.whichDates.SelectedIndex = .xDayIdx
                    QFE1.txtNowDates.Text = Now
                    Counters.qBuyandHold.qHit = False
                    '                    Form1.Dates_.Text = Counters.last_Date
                    '                    Counters.last_Day = Counters.last_day1
                    '                    Counters.last_Date = Counters.last_date1
                    Call ByDOW(Counters.last_day1)
                End If
                QFE1.Dates_.Text = Counters.last_Date
                Application.DoEvents()
            Next
        End With
    End Sub

End Module

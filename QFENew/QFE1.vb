Public Class QFE1
    Public oleDatabaseSignals As OleDb.OleDbCommand, oleDatabaseTrades As OleDb.OleDbCommand, oleDatabaseTotals As OleDb.OleDbCommand
    Public oleConnSignals As System.Data.OleDb.OleDbConnection, oleConnTrades As System.Data.OleDb.OleDbConnection ', oleDatabaseTotals As OleDb.OleDbCommand
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'QFE_SignalsDataSet.Signals' table. You can move, or remove it, as needed.;       

        Static xNo As Integer
        oleDatabaseSignals = New OleDb.OleDbCommand
        oleConnSignals = New OleDb.OleDbConnection
        oleConnSignals.ConnectionString = ""
        oleConnSignals.ConnectionString = "Provider=Microsoft.JET.OLEDB.4.0;" &
                                                        "Data Source=C:\public\QFE_Signals.mdb;"
        oleDatabaseSignals.Connection = oleConnSignals
        Me.oleConnSignals.Open()
        '        oleDatabaseTotals = New OleDb.OleDbCommand
        '        oleConnTotals = New OleDb.OleDbConnection
        '        oleConnTotals.ConnectionString = ""
        '        oleConnTotals.ConnectionString = "Provider=Microsoft.JET.OLEDB.4.0;" &
        '                                                        "Data Source=C:\public\QFE_Signals.mdb;"
        '        oleDatabaseTotals.Connection = oleConnSignals
        '        Me.oleConnSignals.Open()
        Me.dbTextBox0.Text = Me.oleDatabaseSignals.Connection.ConnectionString
        'Me.SignalsTableAdapter.Fill(Me.QFE_SignalsDataSet.Signals)
        oleDatabaseTrades = New OleDb.OleDbCommand
        oleConnTrades = New OleDb.OleDbConnection
        oleConnTrades.ConnectionString = ""
        oleConnTrades.ConnectionString = "Provider=Microsoft.JET.OLEDB.4.0;" &
                                                        "Data Source=C:\Public\QFE_trades.mdb;"
        oleDatabaseTrades.Connection = oleConnTrades
        Me.oleConnTrades.Open()
        Me.dbTextBox1.Text = Me.oleDatabaseTrades.Connection.ConnectionString
        'Me.SignalsTableAdapter.Fill(Me.QFE_SignalsDataSet.Signals)
        buyDowTxt.Text = "txtParameters1.txt"
        Call Rdoparams1()
        Call Rdoparams2()
        Call Rdoparams3()
        Call Rdoparams4()
        Call Rdoparams5()
        xNo = 0
        DataGridViewSecurities.ColumnCount = 9
        DataGridViewSecurities.Columns(0).Name = "Name"
        DataGridViewSecurities.Columns(1).Name = "FirstDate"
        DataGridViewSecurities.Columns(2).Name = "LastDate"
        DataGridViewSecurities.Columns(3).Name = "OPen"
        DataGridViewSecurities.Columns(4).Name = "High"
        DataGridViewSecurities.Columns(5).Name = "Low"
        DataGridViewSecurities.Columns(6).Name = "Close"
        '        DataGridViewSecurities.Rows.Add()
        For Each foundFile As String In My.Computer.FileSystem.GetFiles("c:\temp")
            xNo = xNo + 1
            XSectorSecurities.Items.Add(foundFile)
            DataGridViewSecurities.Rows.Add(foundFile)
            DataGridViewSecurities.Rows.Item(xNo).Cells(1).Value = "f"
            DataGridViewSecurities.Rows.Item(xNo).Resizable = True
            DataGridViewSecurities.Rows.Item(xNo).Resizable = True
            '        = Format(xNo, "00000")
        Next
        Call Rdosecurities()
        oleConnSignals.Close()
    End Sub
    Public Sub Rdosecurities()
        Static x As Integer, xx As Integer
        FileSystem.FileOpen(1, "C:\Qfe\securitieslist1.txt", OpenMode.Input, OpenAccess.Read)
        FileSystem.Input(1, xx)
        If xx = 1 Then
            Me.qBXSector.Checked = True
        Else
            Me.qBXSector.Checked = False
        End If
        x = -1
        Do While Not EOF(1) And x < Me.XSectorSecurities.Items.Count - 1
            x = x + 1
            FileSystem.Input(1, xx)
            If xx = 1 Then
                Me.XSectorSecurities.SetItemChecked(x, True)
            Else
                Me.XSectorSecurities.SetItemChecked(x, False)
            End If
        Loop
        FileSystem.FileClose(1)
        '       '       FileSystem.FileOpen(1, "C:\Qfe\securitieslist1.txt", OpenMode.Input, OpenAccess.Read)
        '      '      secno = -1
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
    Private Sub Run_Click1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Run.Click
        Static toolStripStatusLabel1 As ToolStripStatusLabel, fStr As String
        Me.Text = "Pending . . .."
        Run.BackColor = Color.DarkTurquoise
        Run.Text = Now()
        MsgBox("Starting . . ..", MsgBoxStyle.Exclamation, "Initialization")
        MessageBox.Show("Starting1 . . ..")
        Counters.THHits = 0
        Counters.THMisses = 0
        Counters.signalsSaved = 0
        Counters.signalsUnSaved = 0
        Counters.totTradeNo = 0
        Counters.SystemNumber = 0
        Counters.Iteration = 0
        Counters.totalIterations = 0
        Counters.iterationsWritten = 0
        Counters.elapsedSeconds = 0
        Counters.THHits = 0
        Counters.THMisses = 0
        Counters.signalsSaved = 0
        Counters.signalstradesth = 0
        Counters.Onn = 0
        Counters.Off = 0
        Counters.threshHold = Me.quantumThreshHold.Value
        Counters.Incr_ = 0
        Me.quantumThreshholdtxt1.Text = Format$(Counters.threshHold)
        Me.startTime.Text = Counters.startSeconds
        Counters.startSeconds = DateDiff(DateInterval.Second, Now.Date, Now)
        Counters.currentSeconds = DateDiff(DateInterval.Second, Now.Date, Now)
        Counters.elapsedSeconds = Counters.currentSeconds - Counters.startSeconds
        Counters.iterationsPerSecond = Counters.iterationsPerSecond / Counters.totalIterations
        Me.iterpersecond.Text = Counters.iterationsPerSecond
        oleConnSignals.Close()
        oleDatabaseSignals.CommandText = "DELETE FROM Signals"
        oleConnSignals.Open()
        oleDatabaseSignals.ExecuteNonQuery()
        oleConnTrades.Close()
        oleDatabaseTrades.CommandText = "DELETE FROM Trades"
        oleConnTrades.Open()
        oleDatabaseTrades.ExecuteNonQuery()
        Me.DataGridViewSecurities.Update()
        start_Seconds = DateDiff(DateInterval.Second, Now.Date, Now)
        Call Write_AllParameters()
        Counters.THHits = 0
        '        OleDBC.Connection = conn2
        Application.EnableVisualStyles()
        '        ReDim RMainTS
        '       ReDim TS1(2500)
        '       ReDim RZA__TS(2500)
        '      ReDim RZA__TS1(2500)
        '   StatusStrip1.Items.Add(Now)
        Run.BackColor = Color.AntiqueWhite
        Call RunIt(RMain)
        fStr = "C:\Public\QFE_Signals1.mdb"
        My.Computer.FileSystem.CopyFile("C:\Public\QFE_Signals.mdb", fStr, True)
        Me.dbTextBox1.Text = fStr
        Run.BackColor = Color.Red
        toolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        toolStripStatusLabel1.Text = "gg" ' DateAndTime
        '     StatusStrip1.Items.Insert(1, DateTime.Today)
    End Sub

    Private Sub initBuyInDay_Click(sender As Object, e As EventArgs) Handles initBuyInDay.Click
        Static x As Integer
        For x = 0 To BuyInDay.Items.Count - 1
            BuyInDay.SetItemChecked(x, True)
        Next x
    End Sub



    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
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

    Private Sub QuantumThreshHold_ValueChanged(sender As Object, e As EventArgs) Handles quantumThreshHold.ValueChanged
        Counters.threshHold = quantumThreshHold.Value
    End Sub

    Private Sub SetAll1_Clickr(sender As Object, e As EventArgs)
        Static x As Integer
        For x = 0 To XSectorSecurities.Items.Count - 1
            XSectorSecurities.SetItemChecked(x, True)
        Next x
    End Sub

    Private Sub SetAll1_Click(sender As Object, e As EventArgs)
        Static x As Integer
        For x = 0 To XSectorSecurities.Items.Count - 1
            XSectorSecurities.SetItemChecked(x, True)
        Next x
    End Sub

    Private Sub buyZScoreMode_SelectedIndexChanged(sender As Object, e As EventArgs) Handles buyZScoreMode.SelectedIndexChanged

    End Sub

    Private Sub clear1_Click(sender As Object, e As EventArgs) Handles clear1.Click
        Static x As Integer
        For x = 0 To XSectorSecurities.Items.Count - 1
            XSectorSecurities.SetItemChecked(x, False)
        Next x
    End Sub

    Private Sub dbTextBox0_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub qSaveOnSignals_CheckedChanged(sender As Object, e As EventArgs) Handles qSaveOnSignals.CheckedChanged

    End Sub

    Private Sub qSaveTrades0_CheckedChanged(sender As Object, e As EventArgs) Handles qSaveTrades0.CheckedChanged

    End Sub

    Private Sub qSaveTrades1_CheckedChanged(sender As Object, e As EventArgs) Handles qSaveTrades1.CheckedChanged

    End Sub

    Private Sub Message_TextChanged(sender As Object, e As EventArgs) Handles Message.TextChanged

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub StatusStrip3_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs)

    End Sub

    Private Sub StatusStrip2_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs)
    End Sub

    Private Sub BZScrVal__Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub buyDOW_SelectedIndexChanged(sender As Object, e As EventArgs) Handles buyDOW.SelectedIndexChanged

    End Sub

    Private Sub Dates_Click(sender As Object, e As EventArgs) Handles Dates.Click

    End Sub

    Private Sub whichDates_SelectedIndexChanged(sender As Object, e As EventArgs) Handles whichDates.SelectedIndexChanged

    End Sub

    Private Sub buyEntryTxt_TextChanged(sender As Object, e As EventArgs) Handles buyEntryTxt.TextChanged

    End Sub

    Private Sub On__TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub qsaveoffsignals1_CheckedChanged(sender As Object, e As EventArgs) Handles qsaveoffsignals1.CheckedChanged

    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click

    End Sub

    Private Sub endDateText_TextChanged(sender As Object, e As EventArgs) Handles endDateText.TextChanged

    End Sub

    Private Sub dbTextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub StatusStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles StatusStrip1.ItemClicked

    End Sub

    Private Sub buyTrigger_SelectedIndexChanged(sender As Object, e As EventArgs) Handles buyTrigger.SelectedIndexChanged

    End Sub

    Private Sub QFE_Click(sender As Object, e As EventArgs) Handles QFE.Click

    End Sub

    Private Sub TDeMark_Click(sender As Object, e As EventArgs) Handles TDeMark.Click

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub sellEntry_SelectedIndexChanged(sender As Object, e As EventArgs) Handles sellEntry.SelectedIndexChanged

    End Sub

    Private Sub clearBTrigger_Click(sender As Object, e As EventArgs) Handles clearBTrigger.Click
        Static x As Integer
        For x = 0 To buyTrigger.Items.Count - 1
            buyTrigger.SetItemChecked(x, False)
        Next x
        buyTrigger.SetItemChecked(0, True)
    End Sub

    Private Sub setAll1_Click_1(sender As Object, e As EventArgs) Handles setAll1.Click
        Static x As Integer
        For x = 0 To XSectorSecurities.Items.Count - 1
            XSectorSecurities.SetItemChecked(x, True)
        Next x
    End Sub
    Private Sub SetBTrigger_Click(sender As Object, e As EventArgs) Handles setBTrigger.Click
        Static x As Integer
        For x = 0 To buyTrigger.Items.Count - 1
            buyTrigger.SetItemChecked(x, True)
        Next x
    End Sub

End Class



' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
'Protected Overrides Sub Finalize()
'    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
'    Dispose(False)
'    MyBase.Finalize()
'End Sub

' This code added by Visual Basic to correctly implement the disposable pattern.




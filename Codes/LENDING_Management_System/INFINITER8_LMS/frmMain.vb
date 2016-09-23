Public Class frmMain

    Dim LL, II, PP As Integer
    Dim TXT As String

    Private Sub BorrowerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BorrowerToolStripMenuItem.Click
        frmBorrower.ShowDialog()
    End Sub

    Private Sub LoanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoanToolStripMenuItem.Click
        frmLoan.ShowDialog()
    End Sub

    Private Sub CollectablesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CollectablesToolStripMenuItem.Click
        FrmCollectorCollectables.ShowDialog()
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ' lblTime.Text = Format(Date.Now, "Long Time")
    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If MsgBox("Are you sure you want to close window?", MsgBoxStyle.YesNo, "Close Window") = MsgBoxResult.No Then
            e.Cancel = True
        Else
            End
        End If

    End Sub

    Private Sub frmMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        getData()


    End Sub

    Private Sub PostingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PostingToolStripMenuItem.Click
        ' frmPosting.ShowDialog()
    End Sub

    Private Sub ListOfRemittanceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListOfRemittanceToolStripMenuItem.Click
        frmCollectorRemittance.ShowDialog()
    End Sub

    Private Sub SummaryReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SummaryReportToolStripMenuItem.Click
        frmSummaryAll.ShowDialog()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        frmBorrower.ShowDialog()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        frmLoan.ShowDialog()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        FrmCollectorCollectables.ShowDialog()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        frmCollectorRemittance.ShowDialog()
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        frmSummaryAll.ShowDialog()
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        'frmPosting.ShowDialog()
    End Sub

    Private Sub LogoutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogoutToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub ConfigureToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConfigureToolStripMenuItem.Click
        frmConfiguration.ShowDialog()
    End Sub

    Private Sub UserToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserToolStripMenuItem.Click
        frmUser.ShowDialog()
    End Sub


    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        frmPastDue.ShowDialog()
    End Sub

    Private Sub ListOfPastDueToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListOfPastDueToolStripMenuItem.Click
    End Sub

    Private Sub WeeklyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frmWeeklyPastDue.ShowDialog()
    End Sub

    Private Sub LoanToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoanToolStripMenuItem1.Click
        frmPastDue.ShowDialog()
    End Sub

    Private Sub WeeklyToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frmWeeklyPastDue.ShowDialog()
    End Sub

    Private Sub DeleteDataToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)
        frmDeleteData.ShowDialog()
    End Sub

    Private Sub TestToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs)
        Dim test As String = DateAdd("m", 0, DateSerial(Year(Today), Month(Today), 1)).ToString()
        MsgBox(test)
    End Sub

End Class

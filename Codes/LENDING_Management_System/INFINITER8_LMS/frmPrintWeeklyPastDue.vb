Imports MySql.Data.MySqlClient
Public Class frmPrintWeeklyPastDue

    Dim totAmountDue As Double
    Dim totPaidAmount As Double
    Dim totBalance As Double
    Dim y As Integer

    Private Sub LoadRemittance()

        Try


            ConnDB()
            sqL = "SELECT DateRemitted2, CONCAT(Lastname, ', ', FirstName, ' ', MI) as BorName, Area, (WRemittance + PrevBalance) as AmountDue, TotalAmount, ((WRemittance + PrevBalance) - TotalAmount) as aBalance, TransactionNO FROM borrower as B, loan as L, remittance as R WHERE B.BorrowerNo = L.BorrowerNo AND L.LoanNo = R.LoanNo AND Collector LIKE '" & frmWeeklyPastDue.cmbCollector.Text & "%' AND DateRemitted2 LIKE '" & frmWeeklyPastDue.weeks & "%' AND Area LIKE '" & frmWeeklyPastDue.cmbArea.Text & "%' AND ((WRemittance + PrevBalance) - TotalAmount) > 0 ORDER BY DateRemitted2, Lastname"
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            totAmountDue = 0
            totPaidAmount = 0
            totBalance = 0
            Do While dr.Read = True
                y += 18
                dgw.Rows.Add(dr(0), dr(1), dr(2), Format(dr(3), "#,##0.00"), Format(dr(4), "#,##0.00"), Format(dr(5), "#,##0.00"), dr(6))
                totAmountDue += dr(3)
                totPaidAmount += dr(4)
                totBalance += dr(5)
            Loop
            lblAmountDue.Text = Format(totAmountDue, "#,##0.00")
            lblPaidAmount.Text = Format(totPaidAmount, "#,##0.00")
            lblBalance.Text = Format(totBalance, "#,##0.00")
            lbly.Text = y

            If Me.Height + y > 730 Then
                Me.Height = 730
            ElseIf Me.Height + y < 730 Then
                Me.Height = 730
            End If

            If Me.dgw.Height + y > 450 Then
                dgw.Height = 450


            ElseIf Me.dgw.Height + y < 450 Then
                dgw.Height = 450

            End If

            If Panel1.Height + y > 667 Then
                Panel1.Height = 667
            ElseIf Panel1.Height + y < 667 Then
                Panel1.Height = 667
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub


    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim bm As New Bitmap(Me.Panel1.Width, Me.Panel1.Height)

        Panel1.DrawToBitmap(bm, New Rectangle(0, 0, Me.Panel1.Width, Me.Panel1.Height))
        e.Graphics.DrawImage(bm, 25, 60)

        Dim aPS As New PageSetupDialog
        aPS.Document = PrintDocument1
    End Sub

    Private Sub PrintDocument1_QueryPageSettings(ByVal sender As Object, ByVal e As System.Drawing.Printing.QueryPageSettingsEventArgs) Handles PrintDocument1.QueryPageSettings

        e.PageSettings.Landscape = False

    End Sub


    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Me.Close()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        PrintDialog1.Document = Me.PrintDocument1

        Dim ButtonPressed As DialogResult = PrintDialog1.ShowDialog()
        If (ButtonPressed = DialogResult.OK) Then
            Me.Height = Me.Height + Val(lbly.Text)
            Me.dgw.Height = Me.dgw.Height + Val(lbly.Text)
            Panel1.Height = Panel1.Height + Val(lbly.Text)
            PrintDocument1.Print()
            Me.Close()
        End If
    End Sub

    Private Sub frmPrintRemittance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Location = New Point(270, 5)
        Me.Height = 378
        Me.dgw.Height = 50
        Me.Panel1.Height = 275
        y = 0
        lbldate.Text = Date.Now.ToString("MMMM dd, yyyy")
        lblWeekDate.Text = frmWeeklyPastDue.dtpWeek.Value.ToString("MMM dd, yyyy")
        If frmWeeklyPastDue.cmbCollector.Text = "" Then
            lblCollector.Text = "(ALL)"
        Else
            lblCollector.Text = frmWeeklyPastDue.cmbCollector.Text
        End If


        If frmWeeklyPastDue.cmbArea.Text = "" Then
            lblArea.Text = "(ALL)"
        Else
            lblArea.Text = frmWeeklyPastDue.cmbArea.Text
        End If
      

        LoadRemittance()
    End Sub
End Class
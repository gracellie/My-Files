
Imports MySql.Data.MySqlClient
Public Class frmPrintWeeklyRem

    Dim y As Integer
    Private Sub LoadBorrower()
        Try
            sqL = "SELECT CONCAT(Lastname, ', ', Firstname, ' ', MI) as BorName, Area, Collector FROM borrower WHERE BorrowerNo ='" & frmLoan.txtBorrowerNo.Text & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)

            If dr.Read = True Then
                lblBorrowerName.Text = dr(0)
                lblAddress.Text = dr(1)
                lblCollector.Text = dr(2)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub LoadWeeklyRemitted()
        Try
            sqL = "SELECT DateRemitted2, wRemittance, PrevBalance, (wRemittance + PrevBalance) as AmountDue, TotalAmount, R.Balance FROM loan as L, remittance as R WHERE L.LoanNo = R.LoanNo AND L.LoanNo = '" & lblLoanNo.Text & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            Do While dr.Read = True
                y += 18
                dgw.Rows.Add(dr(0), Format(dr(1), "#,##0.00"), Format(dr(2), "#,##0.00"), Format(dr(3), "#,##0.00"), Format(dr(4), "#,##0.00"), Format(dr(5), "#,##0.00"))
            Loop
            y = y - 15
            Me.Height = Me.Height + y
            Panel1.Height = Panel1.Height + y
            dgw.Height = dgw.Height + y

            If Me.Height + y < 730 Then
                Me.Height = 730
            End If

            lblline.Location = New Point(6, 283 + y)
            Panel2.Location = New Point(9, 298 + y)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub


    Private Sub LoadInfo()
        With frmLoan
            lblrate.Text = .cmbInterestRate.Text & "%"
            lblPrincipal.Text = .txtPricipalLoan.Text
            lblInterest.Text = .txtInterestAmount.Text
            lblMatValue.Text = .txtMatValue.Text
            lblBalance.Text = .txtBalance.Text
        End With
    End Sub

    Private Sub frmPrintWeeklyRem_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Location = New Point(270, 3)
        Me.Height = 381
        Me.dgw.Height = 41
        Me.Panel1.Height = 275
        y = 0
        lblLoanNo.Text = frmLoan.lblLoanNo.Text
        LoadBorrower()
        LoadWeeklyRemitted()
        LoadInfo()
    End Sub
    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim bm As New Bitmap(Me.Panel1.Width, Me.Panel1.Height)

        Panel1.DrawToBitmap(bm, New Rectangle(0, 0, Me.Panel1.Width, Me.Panel1.Height))
      

        e.Graphics.DrawImage(bm, 50, 60)
        Dim aPS As New PageSetupDialog
        aPS.Document = PrintDocument1
    End Sub

    Private Sub PrintDocument1_QueryPageSettings(ByVal sender As Object, ByVal e As System.Drawing.Printing.QueryPageSettingsEventArgs) Handles PrintDocument1.QueryPageSettings

        '   If m_nxtPage = 1 Then
        e.PageSettings.Landscape = False
        '     ElseIf m_nxtPage = 2 Then

        '   e.PageSettings.Landscape = False
        ' End If

    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Me.Close()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        PrintDialog1.Document = Me.PrintDocument1

        Dim ButtonPressed As DialogResult = PrintDialog1.ShowDialog()
        If (ButtonPressed = DialogResult.OK) Then
            PrintDocument1.DefaultPageSettings.Margins.Left = 15
            PrintDocument1.DefaultPageSettings.Margins.Top = 30
            PrintDocument1.Print()
            Me.Close()
        End If
    End Sub
End Class
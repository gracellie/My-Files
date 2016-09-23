Imports MySql.Data.MySqlClient
Public Class frmPrintPastDue
    Dim y As Integer

    Private Sub LoadPastdue()
        Dim totPastDue As Double
        Try
            sqL = "SELECT MatDate, Concat(Lastname, ', ', Firstname, ' ', MI) as BorName, Area, MatValue, WRemittance, L.Balance, COUNT(TransactionNo) as RemNum, Duration FROM loan as L, borrower as B, remittance as R WHERE L.BorrowerNo = B.BorrowerNo AND L.LoanNo = R.LoanNo AND L.Balance > 0 AND Collector LIKE '" & frmPastDue.cmbCollector.Text & "%' AND AREA LIKE '" & frmPastDue.cmbArea.Text & "%' GROUP BY L.LoanNo ORDER BY MatDate, Lastname"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)

            dgw.Rows.Clear()
            totPastDue = 0
            y = 0
            Do While dr.Read = True
                If dr(6) = dr(7) Or dr(6) = dr(7) Then
                    y += 18
                    dgw.Rows.Add(dr(0), dr(1), dr(2), Format(dr(3), "#,##0.00"), Format(dr(4), "#,##0.00"), Format(dr(5), "#,##0.00"))
                    totPastDue += dr(5)
                End If

            Loop

            lblTotalBalance.Text = Format(totPastDue, "#,##0.0")

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
        e.Graphics.DrawImage(bm, 50, 60)

        Dim aPS As New PageSetupDialog
        aPS.Document = PrintDocument1
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

    Private Sub PrintDocument1_QueryPageSettings(ByVal sender As Object, ByVal e As System.Drawing.Printing.QueryPageSettingsEventArgs) Handles PrintDocument1.QueryPageSettings

        e.PageSettings.Landscape = False

    End Sub


    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Me.Close()
    End Sub

    Private Sub frmPrintPastDue_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Location = New Point(270, 5)
        Me.Height = 378
        Me.dgw.Height = 50
        Me.Panel1.Height = 275
        y = 0

        If frmPastDue.cmbCollector.Text = "" Then
            lblCollector.Text = "(ALL)"
        Else
            lblCollector.Text = frmPastDue.cmbCollector.Text
        End If

        If frmPastDue.cmbArea.Text = "" Then
            lblArea.Text = "(ALL)"
        Else
            lblArea.Text = frmPastDue.cmbArea.Text
        End If
        lbldate.Text = Date.Now.ToString("MMMM dd, yyyy")
        LoadPastdue()
    End Sub
End Class
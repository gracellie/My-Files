
Imports MySql.Data.MySqlClient
Public Class frmRemitted


    Private Sub LoadWeeklyRemitted()
        Try
            sqL = "SELECT DateRemitted2, wRemittance, PrevBalance, (wRemittance + PrevBalance) as AmountDue, TotalAmount, R.Balance FROM loan as L, remittance as R WHERE L.LoanNo = R.LoanNo AND L.LoanNo = '" & frmLoan.lblLoanNo.Text & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            Do While dr.Read = True
                dgw.Rows.Add(dr(0), Format(dr(1), "#,##0.00"), Format(dr(2), "#,##0.00"), Format(dr(3), "#,##0.00"), Format(dr(4), "#,##0.00"), Format(dr(5), "#,##0.00"))
            Loop
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub


    Private Sub frmWeeklyRemitted_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadWeeklyRemitted()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        frmPrintWeeklyRem.ShowDialog()
    End Sub
End Class
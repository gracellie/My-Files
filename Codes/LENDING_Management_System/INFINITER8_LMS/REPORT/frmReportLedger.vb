Imports Microsoft.Reporting.WinForms
Imports MySql.Data.MySqlClient

Public Class frmReportLedger

    Private Sub frmReportLedger_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'infiniter8_dbDataSet.ledger' table. You can move, or remove it, as needed.

        LoadReport()
    End Sub


    Private Sub LoadReport()
        Try
            sqL = "SELECT LedgerNo, DueDate, ReturnAmount, led.Interest, led.TotalAmount, led.LoanNo, isRemitted, PaidDate, PaidAmount,  RemainingBalance, l.InterestRate, l.PrincipalAmount, l.InterestAmount, l.MatValue, l.EffectiveDate, l.MatDate, CONCAT(Lastname, ', ', Firstname, ' ', MI) Borrower, CONCAT(Address,', ', Area) as Address, Collector, l.balance"
            sqL = sqL & " FROM ledger led INNER JOIN loan l ON led.LoanNo = l.LoanNo INNER JOIN borrower b ON b.BorrowerNo = l.BorrowerNo"
            sqL = sqL & " WHERE l.LoanNo = '" & frmLoan.lblLoanNo.Text & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            da = New MySqlDataAdapter(cmd)

            Me.infiniter8_dbDataSet.ledger.Clear()
            da.Fill(Me.infiniter8_dbDataSet.ledger)

            Me.ReportViewer1.SetDisplayMode(DisplayMode.PrintLayout)
            Me.ReportViewer1.ZoomPercent = 90
            Me.ReportViewer1.ZoomMode = Microsoft.Reporting.WinForms.ZoomMode.Percent

            Me.ReportViewer1.RefreshReport()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
End Class
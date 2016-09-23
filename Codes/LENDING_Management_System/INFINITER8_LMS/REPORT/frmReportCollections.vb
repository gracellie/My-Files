Imports MySql.Data.MySqlClient
Imports Microsoft.Reporting.WinForms

Public Class frmReportCollections

    Private Sub frmReportCollections_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'infiniter8_dbDataSet.remittance' table. You can move, or remove it, as needed.
        ' Me.remittanceTableAdapter.Fill(Me.infiniter8_dbDataSet.remittance)

        ' Me.ReportViewer1.RefreshReport()
        LoadReport()
    End Sub

    Private Sub LoadReport()
        Dim pTerm As String
        Dim strCollector As String
        Dim strarea As String
        Dim strHeader As String

        If frmCollectorRemittance.CheckBox1.Checked = True Then
            strHeader = frmCollectorRemittance.dtpWeek.Value.ToString("MMMM dd, yyyy")
        Else
            strHeader = frmCollectorRemittance.cmbMonth.Text & " " & frmCollectorRemittance.cmbYear.Text
        End If

        If frmCollectorRemittance.cmbCollector.Text = "" Then
            strCollector = "All"
        Else
            strCollector = frmCollectorRemittance.cmbCollector.Text
        End If

        If frmCollectorRemittance.cmbArea.Text = "" Then
            strarea = "All"
        Else
            strarea = frmCollectorRemittance.cmbArea.Text
        End If

        If frmCollectorRemittance.cmbTerm.Text = "All" Then
            pTerm = ""
        Else
            pTerm = frmCollectorRemittance.cmbTerm.Text
        End If
        Try

            sqL = "SELECT DateRemitted2, CONCAT(Lastname, ', ', FirstName, ' ', MI) as BorName, Area, (WRemittance + PrevBalance) as AmountDue, AmountRemit, Interest, TotalAmount, Collector FROM borrower as B, loan as L, remittance as R WHERE B.BorrowerNo = L.BorrowerNo AND L.LoanNo = R.LoanNo AND DateRemitted2 LIKE '" & frmCollectorRemittance.noMonth & "%' AND DateRemitted2 LIKE '%" & frmCollectorRemittance.cmbYear.Text & "' AND Collector LIKE '" & frmCollectorRemittance.cmbCollector.Text & "%' AND DateRemitted2 LIKE '" & frmCollectorRemittance.weeks & "%' AND AREA LIKE '" & frmCollectorRemittance.cmbArea.Text & "%' AND PaymentTerm LIKE '" & pTerm & "%' ORDER BY DateRemitted2, Lastname"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            da = New MySqlDataAdapter(cmd)

            Me.infiniter8_dbDataSet.remittance.Clear()
            da.Fill(Me.infiniter8_dbDataSet.remittance)


            Dim header = New ReportParameter("Header", strHeader)
            Dim collector = New ReportParameter("Collector", strCollector)
            Dim area = New ReportParameter("Area", strarea)
            Dim paymentTerm = New ReportParameter("PaymentTerm", frmCollectorRemittance.cmbTerm.Text)
            Dim HeaderParams As ReportParameter() = {header, collector, area, paymentTerm}

            For Each param As ReportParameter In HeaderParams
                ReportViewer1.LocalReport.SetParameters(param)
            Next

            Me.ReportViewer1.SetDisplayMode(DisplayMode.PrintLayout)

            Me.ReportViewer1.ZoomPercent = 90
            Me.ReportViewer1.ZoomMode = Microsoft.Reporting.WinForms.ZoomMode.Percent

            Me.ReportViewer1.RefreshReport()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
End Class
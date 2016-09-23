Imports MySql.Data.MySqlClient
Imports Microsoft.Reporting.WinForms

Public Class frmReportCollectable

    Private Sub frmReportCollectable_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'infiniter8_dbDataSet.Collectable' table. You can move, or remove it, as needed.
        LoadReport()

    End Sub

    Private Sub LoadReport()
        Dim pTerm As String
        Dim strCollector As String
        Dim strarea As String
        Dim strHeader As String

        If FrmCollectorCollectables.CheckBox1.Checked = True Then
            strHeader = FrmCollectorCollectables.dtpWeek.Value.ToString("MMMM dd, yyyy")
        Else
            strHeader = FrmCollectorCollectables.cmbMonth.Text & " " & FrmCollectorCollectables.cmbYear.Text
        End If

        If FrmCollectorCollectables.cmbCollector.Text = "" Then
            strCollector = "All"
        Else
            strCollector = FrmCollectorCollectables.cmbCollector.Text
        End If

        If FrmCollectorCollectables.cmbArea.Text = "" Then
            strarea = "All"
        Else
            strarea = FrmCollectorCollectables.cmbArea.Text
        End If

        If FrmCollectorCollectables.cmbTerm.Text = "All" Then
            pTerm = ""
        Else
            pTerm = FrmCollectorCollectables.cmbTerm.Text
        End If
        Try

            sqL = "SELECT DueDate, CONCAT(Lastname, ', ', Firstname, ' ', MI) as Borrower, Area, L.MatValue, ReturnAmount,  Interest, TotalAmount, Collector FROM borrower as B, loan as L, ledger as Led WHERE B.BorrowerNo = L.BorrowerNo AND L.LoanNO = Led.LoanNo AND DueDate LIKE '" & FrmCollectorCollectables.noMonth & "%' AND DueDate LIKE '%" & FrmCollectorCollectables.cmbYear.Text & "' AND Collector LIKE '" & FrmCollectorCollectables.cmbCollector.Text & "%' AND DueDAte Like '" & FrmCollectorCollectables.weeks & "%' AND isRemitted = 'NO' AND Area LIKE '" & FrmCollectorCollectables.cmbArea.Text & "%' AND PaymentTerm LIKE '" & pTerm & "%' ORDER BY DueDate, Lastname"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            da = New MySqlDataAdapter(cmd)

            Me.infiniter8_dbDataSet.Collectable.Clear()
            da.Fill(Me.infiniter8_dbDataSet.Collectable)


            Dim header = New ReportParameter("Header", strHeader)
            Dim collector = New ReportParameter("Collector", strCollector)
            Dim area = New ReportParameter("Area", strarea)
            Dim paymentTerm = New ReportParameter("PaymentTerm", FrmCollectorCollectables.cmbTerm.Text)
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
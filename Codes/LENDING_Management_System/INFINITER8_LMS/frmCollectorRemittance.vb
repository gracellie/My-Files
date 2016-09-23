
Imports MySql.Data.MySqlClient
Public Class frmCollectorRemittance
    Public updating As Boolean

    Public weeks As String
    Public isweekly As Boolean
    Public monthly As String
    Public noMonth As String

    Private Sub LoadCollectors()
        Try
            sqL = "SELECT CollectorName FROM collector Order By CollectorName"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            cmbCollector.Items.Clear()
            Do While dr.Read = True
                cmbCollector.Items.Add(dr(0))
            Loop
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub
    Private Sub IntMonth()

        noMonth = ""
        If cmbMonth.Text = "January" Then
            noMonth = "01"
        ElseIf cmbMonth.Text = "February" Then
            noMonth = "02"
        ElseIf cmbMonth.Text = "March" Then
            noMonth = "03"
        ElseIf cmbMonth.Text = "April" Then
            noMonth = "04"
        ElseIf cmbMonth.Text = "May" Then
            noMonth = "05"
        ElseIf cmbMonth.Text = "June" Then
            noMonth = "06"
        ElseIf cmbMonth.Text = "July" Then
            noMonth = "07"
        ElseIf cmbMonth.Text = "August" Then
            noMonth = "08"
        ElseIf cmbMonth.Text = "September" Then
            noMonth = "09"
        ElseIf cmbMonth.Text = "October" Then
            noMonth = "10"
        ElseIf cmbMonth.Text = "November" Then
            noMonth = "11"
        ElseIf cmbMonth.Text = "December" Then
            noMonth = "12"
        End If
    End Sub

    Public Sub LoadRecord()
        IntMonth()
        Dim totalCol As Double
        Me.GroupBox1.Text = "For the month of " & cmbMonth.Text & " " & cmbYear.Text

        Try
            If cmbTerm.Text = "All" Then
                sqL = "SELECT DateRemitted2, CONCAT(Lastname, ', ', FirstName, ' ', MI) as BorName, Area, (WRemittance + PrevBalance) as AmountDue, AmountRemit, Interest, TotalAmount, TransactionNO, L.LoanNo FROM borrower as B, loan as L, remittance as R WHERE B.BorrowerNo = L.BorrowerNo AND L.LoanNo = R.LoanNo AND DateRemitted2 LIKE '" & noMonth & "%' AND DateRemitted2 LIKE '%" & cmbYear.Text & "' AND Collector LIKE '" & cmbCollector.Text & "%' AND DateRemitted2 LIKE '" & weeks & "%' AND Area LIKE '" & cmbArea.Text & "%' ORDER BY DateRemitted2, Lastname"
            Else
                sqL = "SELECT DateRemitted2, CONCAT(Lastname, ', ', FirstName, ' ', MI) as BorName, Area, (WRemittance + PrevBalance) as AmountDue, AmountRemit, Interest, TotalAmount, TransactionNO, L.LoanNo FROM borrower as B, loan as L, remittance as R WHERE B.BorrowerNo = L.BorrowerNo AND L.LoanNo = R.LoanNo AND DateRemitted2 LIKE '" & noMonth & "%' AND DateRemitted2 LIKE '%" & cmbYear.Text & "' AND Collector LIKE '" & cmbCollector.Text & "%' AND DateRemitted2 LIKE '" & weeks & "%' AND Area LIKE '" & cmbArea.Text & "%' AND PaymentTerm LIKE '" & cmbTerm.Text & "%' ORDER BY DateRemitted2, Lastname"
            End If

            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            totalCol = 0
            Do While dr.Read = True
                dgw.Rows.Add(dr(0), dr(1), dr(2), Format(dr(3), "#,##0.00"), Format(dr(4), "#,##0.00"), Format(dr(5), "#,##0.00"), Format(dr(6), "#,##0.00"), dr(7), dr(8))
                totalCol += dr(6)
            Loop
            lblCollectables.Text = Format(totalCol, "#,##0.00")
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub
    Private Sub LoadArea()
        Try
            sqL = "SELECT AREANAME FROM areas ORDER BY AREANAME"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            cmbArea.Items.Clear()
            Do While dr.Read = True
                cmbArea.Items.Add(dr(0))
            Loop
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub



    Private Sub cmbCollector_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        LoadRecord()
    End Sub

    Private Sub dtpWeek_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpWeek.ValueChanged
        weeks = dtpWeek.Value.ToString("MM/dd/yyyy")
        cmbMonth.Text = dtpWeek.Value.ToString("MMMM")
        LoadRecord()
    End Sub

    Private Sub dtpMonth_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        LoadRecord()
    End Sub



    Private Sub CheckBox1_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.CheckState = CheckState.Checked Then
            dtpWeek.Enabled = True
            weeks = dtpWeek.Value.ToString("MM/dd/yyyy")
            cmbMonth.Text = dtpWeek.Value.ToString("MMMM")
            LoadRecord()
        Else
            dtpWeek.Enabled = False
            weeks = ""
            LoadRecord()
        End If

    End Sub

    Private Sub cmbYear_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbYear.TextChanged
        LoadRecord()
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub frmCollectorRemittance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cmbYear.Text = Date.Now.ToString("yyyy")
        cmbMonth.Text = Date.Now.ToString("MMMM")
        cmbTerm.Text = "All"
        cmbCollector.Text = ""
        cmbArea.Text = ""
        LoadCollectors()
        LoadRecord()
        LoadArea()
    End Sub


    Private Sub dgw_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgw.CellDoubleClick
        ' updating = True
        ' frmPosting.adding = False
        '  frmRemit.ShowDialog()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        'frmPrintRemittance.ShowDialog()
        frmReportCollections.ShowDialog()
    End Sub

    Private Sub cmbYear_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbYear.SelectedIndexChanged
        LoadRecord()
    End Sub

    Private Sub cmbMonth_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbMonth.SelectedIndexChanged
        LoadRecord()
    End Sub

    Private Sub cmbTerm_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbTerm.SelectedIndexChanged
        LoadRecord()
    End Sub

#Region "DELETE"

    Private Sub DeleteRemittance()

        Try
            sqL = "DELETE FROM  remittance WHERE TransactionNo = '" & dgw.CurrentRow.Cells(7).Value & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            Dim i As Integer
            i = cmd.ExecuteNonQuery
            If i > 0 Then
                MsgBox("Deleted successfully", MsgBoxStyle.Information, "Remittance")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub UpdateLoanWhenDeleting()
        Dim aReturn As Double = Val(dgw.CurrentRow.Cells(4).Value.Replace(",", ""))
        Dim aInterest As Double = Val(dgw.CurrentRow.Cells(5).Value.Replace(",", ""))
        Dim totalRemit As Double = Val(dgw.CurrentRow.Cells(6).Value.Replace(",", ""))
        Try


            sqL = "UPDATE loan SET OutStandingCap = OutStandingCap + " & aReturn & ", UnearnedInterest = UnearnedInterest + " & aInterest & ", Balance = Balance + " & totalRemit & ", earnedInterest = earnedInterest - " & aInterest & " WHERE LoanNo = '" & dgw.CurrentRow.Cells(8).Value & "'"

            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

   

    Private Sub UpdateLedger()
        Try
            sqL = "UPDATE ledger SET isRemitted = 'NO', RemainingBalance = 0, PaidDate = '', PaidAmount = 0, TransactionNo = ''  WHERE PaidDate = '" & dgw.CurrentRow.Cells(0).Value & "' AND LoanNo = '" & dgw.CurrentRow.Cells(8).Value & "' AND TransactionNo = '" & dgw.CurrentRow.Cells(7).Value & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            Dim i As Integer = cmd.ExecuteNonQuery()
            If i > 0 Then

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub


    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        If MsgBox("Are you sure you want to delete?", vbYesNo, "DELETE COLLECTION") = vbYes Then
            UpdateLoanWhenDeleting()
            UpdateLedger()
            DeleteRemittance()
            LoadRecord()
        End If

    End Sub
#End Region

    Private Sub cmbCollector_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbCollector.SelectedIndexChanged
        LoadRecord()
    End Sub

    Private Sub cmbArea_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbArea.SelectedIndexChanged
        LoadRecord()
    End Sub
End Class
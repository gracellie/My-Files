
Imports MySql.Data.MySqlClient

Public Class frmSummaryAll
    Dim noMonth As String

    Private Sub LoadCollectors()
        Try
            sqL = "SELECT CollectorName FROM collector Order By CollectorName"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            cmbCollector.Items.Clear()
            cmbCollector.Items.Add("All")
            Do While dr.Read = True
                cmbCollector.Items.Add(dr(0))
            Loop
            cmbCollector.Text = "All"
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
            cmbArea.Items.Add("All")
            Do While dr.Read = True
                cmbArea.Items.Add(dr(0))
            Loop
            cmbArea.Text = "All"
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub
  
    Private Sub LoadRecord()

        Dim totPrincipal As Double
        Dim totOustCap As Double
        Dim totOustBal As Double
        Dim totAccInt As Double
        Dim totEarnInt As Double
        Dim totNoofLoans As Integer

        Dim strTerm As String
        Dim strIsPaid As String
        Dim strDatePaid As String

       


        Dim strMonth As String = dtpWeek.Value.ToString("MM")
        Dim strYear As String = dtpWeek.Value.ToString("yyyy")

        Dim strArea As String
        Dim strCollector As String


        Dim strDateFilter As String

        If chkByDate.Checked = True Then
            strDateFilter = " AND EffectiveDate = '" & dtpWeek.Value.ToString("MM/dd/yyyy") & "'"
        ElseIf chkByMonth.Checked = True Then
            strDateFilter = " AND EffectiveDate LIKE '" & strMonth & "%' AND EffectiveDate LIKE '%" & strYear & "'"
        Else
            strDateFilter = ""
        End If

        If cmbTerm.Text = "All" Then
            strTerm = ""
        Else
            strTerm = cmbTerm.Text
        End If

        If cmbArea.Text = "" Or cmbArea.Text = "All" Then
            strArea = ""
        Else
            strArea = cmbArea.Text
        End If

        If cmbCollector.Text = "" Or cmbCollector.Text = "All" Then
            strCollector = ""
        Else
            strCollector = cmbCollector.Text
        End If

        Try

            If chkFullyPaid.Checked = True Then
                sqL = "SELECT CONCAT(Lastname, ', ', FIrstname, ' ', MI) as BorName, PrincipalAmount, OutstandingCap, L.Balance, UnearnedInterest, EarnedInterest, EffectiveDate, Dp.DateRemitted FROM borrower as B INNER JOIN loan as L ON B.BorrowerNo = L.BorrowerNo LEFT JOIN (select R.LoanNo, L.Balance,MAX( STR_TO_DATE( R.DateRemitted2, '%m/%d/%Y')) as DateRemitted from remittance R inner join loan L ON R.LoanNo = L.LoanNo where L.balance <=0 group by LoanNo,  L.Balance) as Dp ON Dp.LoanNo = L.LoanNo WHERE Collector LIKE '" & strCollector & "%' AND LASTNAME LIKE '" & txtLastname.Text & "%' AND Area LIKE '" & strArea & "%' AND PaymentTerm LIKE '" & strTerm & "%' AND L.Balance <= 0 " & strDateFilter & "  GROUP BY L.LoanNo ORDER BY Lastname"
            ElseIf chkUnPaid.Checked = True Then
                sqL = "SELECT CONCAT(Lastname, ', ', FIrstname, ' ', MI) as BorName, PrincipalAmount, OutstandingCap, L.Balance, UnearnedInterest, EarnedInterest, EffectiveDate, Dp.DateRemitted FROM borrower as B INNER JOIN loan as L ON B.BorrowerNo = L.BorrowerNo LEFT JOIN (select R.LoanNo, L.Balance,MAX( STR_TO_DATE( R.DateRemitted2, '%m/%d/%Y')) as DateRemitted from remittance R inner join loan L ON R.LoanNo = L.LoanNo where L.balance <=0 group by LoanNo,  L.Balance) as Dp ON Dp.LoanNo = L.LoanNo WHERE Collector LIKE '" & strCollector & "%' AND LASTNAME LIKE '" & txtLastname.Text & "%' AND Area LIKE '" & strArea & "%' AND PaymentTerm LIKE '" & strTerm & "%' AND L.Balance > 0 " & strDateFilter & " GROUP BY L.LoanNo ORDER BY Lastname"
            Else
                sqL = "SELECT CONCAT(Lastname, ', ', FIrstname, ' ', MI) as BorName, PrincipalAmount, OutstandingCap, L.Balance, UnearnedInterest, EarnedInterest, EffectiveDate, Dp.DateRemitted FROM borrower as B INNER JOIN loan as L ON B.BorrowerNo = L.BorrowerNo LEFT JOIN (select R.LoanNo, L.Balance,MAX( STR_TO_DATE( R.DateRemitted2, '%m/%d/%Y')) as DateRemitted from remittance R inner join loan L ON R.LoanNo = L.LoanNo where L.balance <=0 group by LoanNo,  L.Balance) as Dp ON Dp.LoanNo = L.LoanNo WHERE Collector LIKE '" & strCollector & "%' AND LASTNAME LIKE '" & txtLastname.Text & "%' AND Area LIKE '" & strArea & "%' AND PaymentTerm LIKE '" & strTerm & "%' " & strDateFilter & " GROUP BY L.LoanNo ORDER BY Lastname"
            End If
          
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            totPrincipal = 0
            totOustCap = 0
            totOustBal = 0
            totAccInt = 0
            totEarnInt = 0
            totNoofLoans = 0

            Do While dr.Read = True
                If dr(3) > 0 Then
                    strIsPaid = ""
                    strDatePaid = ""
                Else
                    strIsPaid = "Fully Paid"
                   
                    Dim dt As Date = Date.Parse(dr(7))
                    strDatePaid = dt.ToString("MM/dd/yyyy")
                End If
                dgw.Rows.Add(dr(0), dr(6), Format(dr(1), "#,##0.00"), Format(dr(2), "#,##0.00"), Format(dr(3), "#,##0.00"), Format(dr(4), "#,##0.00"), Format(dr(5), "#,##0.00"), strIsPaid, strDatePaid)
                totPrincipal += dr(1)
                totOustCap += dr(2)
                totOustBal += dr(3)
                totAccInt += dr(4)
                totEarnInt += dr(5)
                totNoofLoans += 1
            Loop
            lblPrincipal.Text = Format(totPrincipal, "#,##0.00")
            lblOustandingCap.Text = Format(totOustCap, "#,##0.00")
            lblOustandingBal.Text = Format(totOustBal, "#,##0.00")
            lblAccruedInterest.Text = Format(totAccInt, "#,##0.00")
            lblInterestEarned.Text = Format(totEarnInt, "#,##0.00")
            lblTotalNoofLoans.Text = totNoofLoans
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub frmSummaryAll_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '   cmbYear.Text = Date.Now.ToString("yyyy")
        'uncheck checkboxes
        chkByDate.Checked = False
        chkByMonth.Checked = False
        chkFullyPaid.Checked = False
        chkUnPaid.Checked = False

        cmbTerm.Text = "All"
        cmbCollector.Text = ""
        cmbArea.Text = ""
        LoadCollectors()
        LoadRecord()
        LoadArea()
    End Sub


    Private Sub cmbYear_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        LoadRecord()
    End Sub

    Private Sub cmbMonth_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        LoadRecord()
    End Sub


    Private Sub cmbCollector_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbCollector.TextChanged
        LoadRecord()
    End Sub

    Private Sub txtLastname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLastname.TextChanged
        LoadRecord()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        'frmPrintSummary.ShowDialog()
        'frmReportSummary.ShowDialog()
    End Sub

    Private Sub cmbArea_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbArea.TextChanged
        LoadRecord()
    End Sub

    Private Sub cmbTerm_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbTerm.SelectedIndexChanged
        LoadRecord()
    End Sub

    Private Sub chkUnPaid_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkUnPaid.CheckedChanged
        If chkUnPaid.Checked = True Then
            chkFullyPaid.Checked = False

        End If
        LoadRecord()
    End Sub

    Private Sub chkFullyPaid_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkFullyPaid.CheckedChanged
        If chkFullyPaid.Checked = True Then
            chkUnPaid.Checked = False

        End If
        LoadRecord()
    End Sub

    Private Sub chkByDate_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkByDate.CheckedChanged
        If chkByDate.Checked = True Then
            chkByMonth.Checked = False

        End If
        LoadRecord()
    End Sub

    Private Sub chkByMonth_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkByMonth.CheckedChanged
        If chkByMonth.Checked = True Then
            chkByDate.Checked = False

        End If
        LoadRecord()
    End Sub

    Private Sub dtpWeek_ValueChanged(sender As System.Object, e As System.EventArgs) Handles dtpWeek.ValueChanged
        LoadRecord()
    End Sub
End Class
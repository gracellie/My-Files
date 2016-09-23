
Imports MySql.Data.MySqlClient
Public Class FrmCollectorCollectables
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

    Private Sub LoadCollectables()
        IntMonth()
        Dim totalCol As Double
        Me.GroupBox1.Text = "For the month of " & cmbMonth.Text & " " & cmbYear.Text
        Try
            If cmbTerm.Text = "All" Then
                sqL = "SELECT DueDate, CONCAT(Lastname, ', ', FIrstname, ' ', MI) as BorName, Area, L.Matvalue, WRemittance, COllector FROM borrower as B, loan as L, ledger as Led WHERE B.BorrowerNo = L.BorrowerNo AND L.LoanNO = Led.LoanNo AND DueDate LIKE '" & noMonth & "%' AND DueDate LIKE '%" & cmbYear.Text & "' AND Collector LIKE '" & cmbCollector.Text & "%' AND DueDAte Like '" & weeks & "%' AND isRemitted = 'NO' AND Area LIKE '" & cmbArea.Text & "%' AND L.Balance > 0  ORDER BY DueDate, Lastname"
            Else
                sqL = "SELECT DueDate, CONCAT(Lastname, ', ', FIrstname, ' ', MI) as BorName, Area, L.Matvalue, WRemittance, COllector FROM borrower as B, loan as L, ledger as Led WHERE B.BorrowerNo = L.BorrowerNo AND L.LoanNO = Led.LoanNo AND DueDate LIKE '" & noMonth & "%' AND DueDate LIKE '%" & cmbYear.Text & "' AND Collector LIKE '" & cmbCollector.Text & "%' AND DueDAte Like '" & weeks & "%' AND isRemitted = 'NO' AND Area LIKE '" & cmbArea.Text & "%' AND PaymentTerm LIKE '" & cmbTerm.Text & "%' AND L.Balance > 0 ORDER BY DueDate, Lastname"
            End If


            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            totalCol = 0
            Do While dr.Read = True
                dgw.Rows.Add(dr(0), dr(1), dr(2), Format(dr(3), "#,##0.00"), Format(dr(4), "#,##0.00"), dr(5))
                totalCol += dr(4)
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


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FrmCollectorHistory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cmbYear.Text = Date.Now.ToString("yyyy")
        cmbMonth.Text = Date.Now.ToString("MMMM")
        cmbTerm.Text = "All"
        cmbCollector.Text = ""
        cmbArea.Text = ""
        LoadCollectors()

        LoadArea()
        LoadCollectables()
    End Sub


    Private Sub cmbCollector_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbCollector.TextChanged
        LoadCollectables()
    End Sub

    Private Sub dtpWeek_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpWeek.ValueChanged
        weeks = dtpWeek.Value.ToString("MM/dd/yyyy")
        cmbMonth.Text = dtpWeek.Value.ToString("MMMM")
        LoadCollectables()
    End Sub

    Private Sub dtpMonth_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        LoadCollectables()
    End Sub



    Private Sub CheckBox1_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.CheckState = CheckState.Checked Then
            dtpWeek.Enabled = True
            weeks = dtpWeek.Value.ToString("MM/dd/yyyy")
            cmbMonth.Text = dtpWeek.Value.ToString("MMMM")
            LoadCollectables()
        Else
            dtpWeek.Enabled = False
            weeks = ""
            LoadCollectables()
        End If

    End Sub

    Private Sub cmbYear_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        LoadCollectables()
    End Sub


    Private Sub cmbMonth_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        LoadCollectables()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        'frmPrintCollectables.ShowDialog()
        frmReportCollectable.ShowDialog()
    End Sub


    
    Private Sub cmbTerm_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbTerm.SelectedIndexChanged
        LoadCollectables()
    End Sub

    Private Sub cmbYear_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbYear.SelectedIndexChanged
        LoadCollectables()
    End Sub

    Private Sub cmbMonth_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbMonth.SelectedIndexChanged
        LoadCollectables()
    End Sub
End Class
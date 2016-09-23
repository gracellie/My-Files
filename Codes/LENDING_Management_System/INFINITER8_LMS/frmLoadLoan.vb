
Imports MySql.Data.MySqlClient
Public Class frmLoadLoan

    Private Sub LoadLoan()
        Dim strStatus As String
        Try

            If CheckBox1.CheckState = CheckState.Checked Then
                sqL = "SELECT LoanNo, CONCAT(Lastname, ', ', Firstname, ' ', MI) as BorName, Collector, PrincipalAmount, WRemittance, MatDate, Balance FROM loan as L, borrower as B WHERE B.BorrowerNo = L.BorrowerNo ORDER BY Lastname"
            Else
                sqL = "SELECT LoanNo, CONCAT(Lastname, ', ', Firstname, ' ', MI) as BorName, Collector, PrincipalAmount, WRemittance, MatDate, Balance FROM loan as L, borrower as B WHERE B.BorrowerNo = L.BorrowerNo AND BAlance > 0 ORDER BY Lastname"
            End If

            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            Do While dr.Read = True
                If dr(6) <= 0 Then
                    strStatus = "Fully Paid"
                Else
                    strStatus = ""
                End If
                dgw.Rows.Add(dr(0), dr(1), dr(2), Format(dr(3), "#,##0.00"), Format(dr(4), "#,##0.00"), dr(5), strStatus)
            Loop
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub SearchLoan()
        Dim strStatus As String
        Try
            If CheckBox1.CheckState = CheckState.Checked Then
                sqL = "SELECT LoanNo, CONCAT(Lastname, ', ', Firstname, ' ', MI) as BorName, Collector, PrincipalAmount, WRemittance, MatDate, Balance FROM loan as L, borrower as B WHERE B.BorrowerNo = L.BorrowerNo AND Lastname LIKE '" & TextBox1.Text & "%' ORDER BY Lastname"
            Else
                sqL = "SELECT LoanNo, CONCAT(Lastname, ', ', Firstname, ' ', MI) as BorName, Collector, PrincipalAmount, WRemittance, MatDate, Balance FROM loan as L, borrower as B WHERE B.BorrowerNo = L.BorrowerNo AND BAlance > 0 AND Lastname LIKE '" & TextBox1.Text & "%' ORDER BY Lastname"
            End If

            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            Do While dr.Read = True
                If dr(6) <= 0 Then
                    strStatus = "Fully Paid"
                Else
                    strStatus = ""
                End If
                dgw.Rows.Add(dr(0), dr(1), dr(2), Format(dr(3), "#,##0.00"), Format(dr(4), "#,##0.00"), dr(5), strStatus)
            Loop
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub frmLoadLoan_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F1 Then
            If frmLoan.LSearch = True Then
                frmLoan.lblLoanNo.Text = dgw.CurrentRow.Cells(0).Value
                Me.Close()
            End If
        End If
    End Sub

    Private Sub frmLoadLoan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        CheckBox1.CheckState = CheckState.Unchecked
        LoadLoan()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        SearchLoan()
    End Sub


    Private Sub dgw_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgw.CellDoubleClick
        If frmLoan.LSearch = True Then
            frmLoan.lblLoanNo.Text = dgw.CurrentRow.Cells(0).Value
            Me.Close()
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If TextBox1.Text <> "" Then
            SearchLoan()
        Else
            LoadLoan()
        End If



      
    End Sub
End Class
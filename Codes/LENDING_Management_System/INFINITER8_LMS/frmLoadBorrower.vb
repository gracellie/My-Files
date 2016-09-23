
Imports MySql.Data.MySqlClient
Public Class frmLoadBorrower

    Private Sub LoadBorrower()
        Try
            sqL = "SELECT BorrowerNo, CONCAT(Lastname, ', ', Firstname, ' ', MI) as BorName, CONCAT(Address, ', ', Area ) as Address, Collector FROM borrower Order by Lastname, Firstname"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            Do While dr.Read = True
                dgw.Rows.Add(dr(0), dr(1), dr(2), dr(3))
            Loop
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub SearchBorrower()
        Try
            sqL = "SELECT BorrowerNo, CONCAT(Lastname, ', ', Firstname, ' ', MI) as BorName, CONCAT(Address, ', ', Area ) as Address, Collector FROM borrower WHERE lastname LIKE '" & TextBox1.Text & "%' Order by Lastname, Firstname"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            Do While dr.Read = True
                dgw.Rows.Add(dr(0), dr(1), dr(2), dr(3))
            Loop
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Function IsNotPaid() As Boolean
        Dim retBol As Boolean
        Try
            sqL = "SELECT * FROM loan where borrowerNo = '" & dgw.CurrentRow.Cells(0).Value & "' AND Balance > 0"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader

            If dr.Read = True Then
                retBol = True
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try

        Return retBol
    End Function



    Private Sub frmLoadBorrower_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadBorrower()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        SearchBorrower()
    End Sub

    Private Sub dgw_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgw.CellDoubleClick
        If frmBorrower.bSearch = True Then
            frmBorrower.txtBorrowerNo.Text = dgw.CurrentRow.Cells(0).Value

            Me.Close()
        End If

        If frmLoan.LSearch = True Then
            If IsNotPaid() = True Then
                MsgBox("This borrower is not paid yet.", MsgBoxStyle.Exclamation, "Warning")
                Exit Sub
            End If
            frmLoan.txtBorrowerNo.Text = dgw.CurrentRow.Cells(0).Value
            frmLoan.txtName.Text = dgw.CurrentRow.Cells(1).Value
            frmLoan.LSearch = False
            frmLoan.txtDuration.Focus()
            Me.Close()
        End If
    End Sub
End Class
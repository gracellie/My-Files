
Imports MySql.Data.MySqlClient
Public Class frmChangePwd








    Private Sub ValidatePassword()
        Try
            sqL = "SELECT * FROM users WHERE Username = '" & txtUsername.Text & "' AND Pwd = '" & txtOldPwd.Text & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)

            If dr.Read = True Then
                ChangePassword()
                Me.Close()
            Else
                MsgBox("Old password is incorrect.", MsgBoxStyle.Critical, "Old Password")
                txtOldPwd.Select()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub ChangePassword()
        Try
            sqL = "UPDATE users SET Pwd = '" & txtConNewPwd.Text & "' WHERE Username ='" & txtUsername.Text & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            Dim i As Integer
            i = cmd.ExecuteNonQuery
            If i > 0 Then
                MsgBox("Password successfully changed..", MsgBoxStyle.Information, "New Password")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()

        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        If txtNewPwd.Text <> txtConNewPwd.Text Then
            MsgBox(" New password did match.", MsgBoxStyle.Critical, "New Password")
            txtNewPwd.SelectAll()
            Exit Sub
        End If
        ValidatePassword()

    End Sub

    Private Sub frmChangePwd_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtConNewPwd.Text = ""
        txtNewPwd.Text = ""
        txtOldPwd.Text = ""
        txtUsername.Text = ""
        With frmUser
            txtUsername.Text = .txtusername.Text
        End With
    End Sub
End Class
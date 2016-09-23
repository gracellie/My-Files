
Imports MySql.Data.MySqlClient

Public Class frmLogin

    Dim strInput As String = ""


  




    Private Sub Login()
        Try
            sqL = "SELECT * FROM users WHERE Username ='" & txtUser.Text & "' AND PWd = '" & txtPassword.Text & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)

            If dr.Read = True Then
                frmMain.Show()
                Me.Hide()
                txtUser.Text = ""
                txtPassword.Text = ""


            Else
                MsgBox("Incorrect username or password!", MsgBoxStyle.Critical, "Login")
                txtPassword.SelectAll()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        Login()

    End Sub

    Private Sub txtUser_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUser.GotFocus
        AcceptButton = btnLogin
    End Sub

    Private Sub txtPassword_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPassword.GotFocus
        AcceptButton = btnLogin
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        End
    End Sub

    Private Sub frmLogin_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F1 Then
            frmDatabase.ShowDialog()
        End If
    End Sub

    Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Fadein()
        getData()
    End Sub

   
End Class
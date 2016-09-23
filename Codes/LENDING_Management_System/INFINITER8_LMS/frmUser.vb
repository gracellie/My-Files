
Imports MySql.Data.MySqlClient
Public Class frmUser

    Dim adding As Boolean
    Dim updating As Boolean

    Private Sub AddUser()
        Try
            sqL = "INSERT INTO users VALUES('" & txtusername.Text & "', '" & txtCpwd.Text & "', '" & cmbRole.Text & "')"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            Dim i As Integer
            i = cmd.ExecuteNonQuery
            If i > 0 Then
                MsgBox("User added successfully..", MsgBoxStyle.Information, "Add user")
                LoadUser()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()

        End Try
    End Sub

    Private Sub LoadUser()
        Try
            sqL = "SELECT username FROM users Order By username"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)

            ListBox1.Items.Clear()
            Do While dr.Read = True
                ListBox1.Items.Add(dr(0))
            Loop
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub UpdateUser()
        Try
            sqL = "UPDATE users Set pwd = '" & txtCpwd.Text & "' WHERE username = '" & txtusername.Text & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            Dim i As Integer
            i = cmd.ExecuteNonQuery
            If i > 0 Then
                MsgBox("Password successfully changed..", MsgBoxStyle.Information, "Change Password")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub DeleteUser()
        Try
            sqL = "DELETE FROM users Where Username = '" & ListBox1.SelectedItem & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            Dim i As Integer
            i = cmd.ExecuteNonQuery
            If i > 0 Then
                MsgBox("user deleted successfully..", MsgBoxStyle.Information, "Delet User")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub EnabledField()
        txtusername.Enabled = True
        txtCpwd.Enabled = True
        txtPwd.Enabled = True
        cmbRole.Enabled = True
    End Sub

    Private Sub DisAbledField()
        txtusername.Enabled = False
        txtCpwd.Enabled = False
        txtPwd.Enabled = False
        cmbRole.Enabled = False
    End Sub

    Private Sub ClearField()
        txtCpwd.Text = ""
        txtPwd.Text = ""
        txtusername.Text = ""
        cmbRole.Text = ""
    End Sub

    Private Sub EnabledButton()
        btnAdd.Enabled = True
        btnDelete.Enabled = True

        btnSave.Enabled = False
        btnCancel.Enabled = False
    End Sub

    Private Sub DisabledButton()
        btnAdd.Enabled = False
        btnDelete.Enabled = False

        btnSave.Enabled = True
        btnCancel.Enabled = True
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        ClearField()
        EnabledButton()
        DisAbledField()
    End Sub

    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        EnabledButton()
        DisAbledField()
        ClearField()
        LoadUser()
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        adding = True
        updating = False

        ClearField()
        DisabledButton()
        EnabledField()
        txtusername.Focus()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If adding = True Then

            If txtCpwd.Text <> txtPwd.Text Then
                MsgBox("Password did not match..!", MsgBoxStyle.Critical, "Mismatch")
                Exit Sub
            End If
            AddUser()
            DisAbledField()
            ClearField()
            EnabledButton()
            LoadUser()
        Else
            If txtCpwd.Text <> txtPwd.Text Then
                MsgBox("Password did not match..!", MsgBoxStyle.Critical, "Mismatch")
                Exit Sub
            End If
            UpdateUser()
            DisAbledField()
            ClearField()
            EnabledButton()
            LoadUser()
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click

        If ListBox1.SelectedItem = "" Then
            MsgBox("Please select user..", MsgBoxStyle.Critical, "Select")
            Exit Sub
        End If
        If ListBox1.SelectedItem = "admin" Or ListBox1.SelectedItem = "Admin" Or ListBox1.SelectedItem = "ADMIN" Then
            MsgBox("Admin Cannot deleted..", MsgBoxStyle.Critical, "Delete")
            Exit Sub
        End If

        If MsgBox("Are you sure you want to delete " & ListBox1.SelectedItem & "", MsgBoxStyle.YesNo, "Delete") = MsgBoxResult.Yes Then
            DeleteUser()
            ClearField()
            LoadUser()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ListBox1.SelectedItem = "" Then
            MsgBox("Please select user..", MsgBoxStyle.Critical, "Select")
            Exit Sub
        End If
        frmChangePwd.ShowDialog()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Try
            sqL = "SELECT Username, Pwd, Role FROM users where username = '" & ListBox1.SelectedItem & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If dr.Read = True Then
                txtusername.Text = dr(0)
                cmbRole.Text = dr(2)
                txtCpwd.Text = dr(1)
                txtPwd.Text = dr(1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub
End Class
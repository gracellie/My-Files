
Imports MySql.Data.MySqlClient
Public Class frmBorrower
    Dim adding As Boolean
    Dim updating As Boolean
    Public bSearch As Boolean

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



    Private Sub GetBorrowerNo()
        Try
            sqL = "SELECT BorrowerNo FROM borrower Order By BorrowerNo Desc"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If dr.Read = True Then
                txtBorrowerNo.Text = Val(dr(0)) + 1
            Else
                txtBorrowerNo.Text = 11000001000
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub


    Private Sub AddBorrower()
        Try
            sqL = "INSERT INTO borrower VALUES('" & txtBorrowerNo.Text & "', '" & txtLastname.Text & "', '" & txtFirstName.Text & "','" & txtMI.Text & "', '" & txtAddress.Text & "', '" & cmbArea.Text & "', '" & cmbCollector.Text & "')"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            Dim i As Integer
            i = cmd.ExecuteNonQuery
            If i > 0 Then
                MsgBox("Borrower successfully saved..", MsgBoxStyle.Information, "Add Borrower")

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub UpdateBorrower()
        Try
            sqL = "Update borrower SET Lastname = '" & txtLastname.Text & "', Firstname = '" & txtFirstName.Text & "', MI = '" & txtMI.Text & "', Address ='" & txtAddress.Text & "', Area = '" & cmbArea.Text & "', Collector ='" & cmbCollector.Text & "' WHERE BorrowerNo = '" & txtBorrowerNo.Text & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            Dim i As Integer
            i = cmd.ExecuteNonQuery
            If i > 0 Then
                MsgBox("Changes successfully saved..", MsgBoxStyle.Information, "Update Borrower")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub GetBorrowerInfo()
        Try
            sqL = "SELECT Lastname, Firstname, MI, Address, Area, Collector FROM borrower WHERE BorrowerNo ='" & txtBorrowerNo.Text & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)

            If dr.Read = True Then
                txtLastname.Text = dr(0)
                txtFirstName.Text = dr(1)
                txtMI.Text = dr(2)
                txtAddress.Text = dr(3)
                cmbArea.Text = dr(4)
                cmbCollector.Text = dr(5)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub



    Private Sub EnabledField()
        txtLastname.Enabled = True
        txtFirstName.Enabled = True
        txtMI.Enabled = True
        txtAddress.Enabled = True
        cmbArea.Enabled = True
        cmbCollector.Enabled = True
    End Sub

    Private Sub DisabledField()
        txtLastname.Enabled = False
        txtFirstName.Enabled = False
        txtMI.Enabled = False
        txtAddress.Enabled = False
        cmbArea.Enabled = False
        cmbCollector.Enabled = False
    End Sub

    Private Sub ClearFields()
        txtBorrowerNo.Text = ""
        txtLastname.Text = ""
        txtFirstName.Text = ""
        txtMI.Text = ""
        txtAddress.Text = ""
        cmbArea.Text = ""
        cmbCollector.Text = ""
    End Sub

    Private Sub EnabledButtons()
        btnAdd.Enabled = True
        btnUpdate.Enabled = True
        btnSearch.Enabled = True
        btnClose.Enabled = True

        btnSave.Enabled = False
        btnCancel.Enabled = False
    End Sub

    Private Sub DisabledButtons()
        btnAdd.Enabled = False
        btnUpdate.Enabled = False
        btnSearch.Enabled = False
        btnClose.Enabled = False

        btnSave.Enabled = True
        btnCancel.Enabled = True
    End Sub


    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmBorrower_Load_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DisabledField()
        EnabledButtons()
        ClearFields()
        LoadCollectors()
        LoadArea()
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        adding = True
        updating = False

        EnabledField()
        ClearFields()
        DisabledButtons()
        GetBorrowerNo()
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click

        If txtLastname.Text = "" Then
            MsgBox("Please select record to change..", MsgBoxStyle.Critical, "Update Record")
            Exit Sub
        End If

        adding = False
        updating = True

        EnabledField()
        DisabledButtons()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If adding = True Then
            AddBorrower()
        Else
            UpdateBorrower()

        End If

        DisabledField()
        ClearFields()
        EnabledButtons()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        DisabledField()
        ClearFields()
        EnabledButtons()
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        bSearch = True
        frmLoadBorrower.ShowDialog()
    End Sub

    Private Sub txtBorrowerNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GetBorrowerInfo()
    End Sub

    Private Sub txtBorrowerNo_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBorrowerNo.TextChanged
        If bSearch = True Then
            GetBorrowerInfo()
            bSearch = False
        End If

    End Sub
End Class
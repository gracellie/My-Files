Imports MySql.Data.MySqlClient
Public Class frmConfiguration



#Region "AREA"


    Private Sub AddArea()
        Try
            sqL = "INSERT INTO areas(AreaName) VALUES('" & txtArea.Text & "')"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            Dim i As Integer
            i = cmd.ExecuteNonQuery
            If i > 0 Then
                MsgBox("Area successfully added..", MsgBoxStyle.Information, "Add Area")
                LoadArea()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub DeleteArea()
        Try
            sqL = "DELETE FROM areas WHERE AreaName = '" & ListBox1.SelectedItem & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            Dim i As Integer
            i = cmd.ExecuteNonQuery
            If i > 0 Then
                MsgBox("Area Deleted", MsgBoxStyle.Information, "Delete")
                LoadArea()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub


    Private Sub LoadArea()
        Try
            sqL = "SELECT AreaName FROM areas Order By Areaname"
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




    Private Sub btnAddArea_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddArea.Click
        If btnAddArea.Text = "Save" Then
            AddArea()
            btnAddArea.Text = "Add"
            btnDeleteArea.Text = "Delete"
            txtArea.Text = ""
            txtArea.Enabled = False
        Else
            txtArea.Enabled = True
            txtArea.Focus()

            btnAddArea.Text = "Save"
            btnDeleteArea.Text = "Cancel"
        End If

       
    End Sub

    Private Sub btnDeleteArea_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteArea.Click
        If btnDeleteArea.Text = "Cancel" Then
            btnAddArea.Text = "Add"
            btnDeleteArea.Text = "Delete"
            txtArea.Text = ""
            txtArea.Enabled = False
        Else
            If Me.ListBox1.SelectedItem = "" Then
                MsgBox("Please Select Area to Delete", MsgBoxStyle.Exclamation, "Delete Collector")
            Else
                If MsgBox("Are you sure you want to delete?", MsgBoxStyle.YesNo, "Delete Area") = MsgBoxResult.Yes Then
                    DeleteArea()
                End If

            End If
        End If
       
    End Sub
#End Region

#Region "COLLECTOR"



    Private Sub AddCollector()
        Try
            sqL = "INSERT INTO collector(CollectorName) VALUES('" & txtCollector.Text & "')"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            Dim i As Integer
            i = cmd.ExecuteNonQuery
            If i > 0 Then
                MsgBox("Collector successfully added..", MsgBoxStyle.Information, "Add Collector")
                LoadCollector()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub DeleteCollector()
        Try
            sqL = "DELETE From collector WHERE CollectorName = '" & ListBox2.SelectedItem & "'"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            Dim i As Integer
            i = cmd.ExecuteNonQuery
            If i > 0 Then
                MsgBox("Collector deleted", MsgBoxStyle.Information, "Delete Collector")
                LoadCollector()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub LoadCollector()
        Try
            sqL = "SELECT CollectorName FROM collector Order By CollectorName"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            ListBox2.Items.Clear()
            Do While dr.Read = True
                ListBox2.Items.Add(dr(0))
            Loop
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub btnAddCollector_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddCollector.Click
        If btnAddCollector.Text = "Save" Then
            AddCollector()
            btnAddCollector.Text = "Add"
            btnDeleteCollector.Text = "Delete"
            txtCollector.Text = ""
            txtCollector.Enabled = False
        Else
            txtCollector.Enabled = True
            txtCollector.Focus()

            btnAddCollector.Text = "Save"
            btnDeleteCollector.Text = "Cancel"
        End If

       
    End Sub

    Private Sub btnDeleteCollector_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteCollector.Click
        If btnDeleteCollector.Text = "Cancel" Then
            btnAddCollector.Text = "Add"
            btnDeleteCollector.Text = "Delete"
            txtCollector.Text = ""
            txtCollector.Enabled = False
        Else
            If Me.ListBox2.SelectedItem = "" Then
                MsgBox("Please Select Collector to Delete", MsgBoxStyle.Exclamation, "Delete Collector")
            Else
                If MsgBox("Are you sure you want to delete?", MsgBoxStyle.YesNo, "Delete Collector") = MsgBoxResult.Yes Then
                    DeleteCollector()
                End If

            End If
        End If
    End Sub


#End Region


    Private Sub AddInterestRate()
        Try
            sqL = "INSERT INTO interestrate(IntRate) VALUES(" & txtInterest.Text & ")"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            Dim i As Integer
            i = cmd.ExecuteNonQuery
            If i > 0 Then
                MsgBox("Interest rate added successfully..", MsgBoxStyle.Information, "Add rate")
                LoadInterest()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub LoadInterest()
        Try
            sqL = "SELECT IntRate FROm interestrate Order By IntRate"
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            ListBox3.Items.Clear()
            Do While dr.Read
                ListBox3.Items.Add(dr(0))
            Loop
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub

    Private Sub DeleteInterest()
        Try
            sqL = "DELETE FROM interestrate WHERE IntRate = " & ListBox3.SelectedItem & ""
            ConnDB()
            cmd = New MySqlCommand(sqL, conn)
            Dim i As Integer
            i = cmd.ExecuteNonQuery
            If i > 0 Then
                MsgBox("Interest rate deleted..", MsgBoxStyle.Information, "Delete rate")
                LoadInterest()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            cmd.Dispose()
            conn.Close()
        End Try
    End Sub


    Private Sub frmConfiguration_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtArea.Enabled = False
        txtCollector.Enabled = False
        txtInterest.Enabled = False

        LoadArea()
        LoadCollector()
        LoadInterest()
    End Sub

    Private Sub btnCloseArea_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCloseArea.Click
        Me.Close()
    End Sub

    
    Private Sub btnCloseCollector_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCloseCollector.Click
        Me.Close()
    End Sub

    Private Sub btnAddInterest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddInterest.Click
        If btnAddInterest.Text = "Save" Then
            AddInterestRate()
            btnAddInterest.Text = "Add"
            btnDeleteInterest.Text = "Delete"
            txtInterest.Text = ""
            txtInterest.Enabled = False
        Else
            txtInterest.Enabled = True
            txtInterest.Focus()

            btnAddInterest.Text = "Save"
            btnDeleteInterest.Text = "Cancel"
        End If
    End Sub

    Private Sub btnDeleteInterest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteInterest.Click
        If btnDeleteInterest.Text = "Cancel" Then
            btnAddInterest.Text = "Add"
            btnDeleteInterest.Text = "Delete"
            txtInterest.Text = ""
            txtInterest.Enabled = False
        Else
            If Me.ListBox3.SelectedItem = 0 Then
                MsgBox("Please Select Interest rate to Delete", MsgBoxStyle.Exclamation, "Delete rate")
            Else
                If MsgBox("Are you sure you want to delete?", MsgBoxStyle.YesNo, "Delete rate") = MsgBoxResult.Yes Then
                    DeleteInterest()
                End If

            End If
        End If
    End Sub

    Private Sub btnCloseInterest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCloseInterest.Click
        Me.Close()
    End Sub
End Class
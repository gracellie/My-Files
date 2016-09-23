Public Class frmMain

    Private Sub btnConvert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConvert.Click
        If cmbWareHouse.Text = "" Or txtDest.Text = "" Or txtSalesFile.Text = "" Then Exit Sub

        Me.Enabled = False
        ConvertToPTU(txtSalesFile.Text, txtDest.Text)
        Me.Enabled = True

        cmbWareHouse.Text = ""
        txtSalesFile.Text = ""
        txtDest.Text = ""
    End Sub

    Private Sub ofdOpen_FileOk(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles ofdOpen.FileOk
        txtSalesFile.Text = ofdOpen.FileName
    End Sub

    Private Sub btnBrowse_Click(sender As System.Object, e As System.EventArgs) Handles btnBrowse.Click
        ofdOpen.ShowDialog()
    End Sub

    Private Sub btnBrowse2_Click(sender As System.Object, e As System.EventArgs) Handles btnBrowse2.Click
        sfdSave.ShowDialog()
    End Sub

    Private Sub sfdSave_FileOk(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles sfdSave.FileOk
        txtDest.Text = sfdSave.FileName
    End Sub

End Class

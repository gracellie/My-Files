<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSalesFile = New System.Windows.Forms.TextBox()
        Me.btnConvert = New System.Windows.Forms.Button()
        Me.txtDest = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.btnBrowse2 = New System.Windows.Forms.Button()
        Me.ofdOpen = New System.Windows.Forms.OpenFileDialog()
        Me.sfdSave = New System.Windows.Forms.SaveFileDialog()
        Me.cmbWareHouse = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 53)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(89, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Branch Sales File"
        '
        'txtSalesFile
        '
        Me.txtSalesFile.BackColor = System.Drawing.Color.White
        Me.txtSalesFile.Location = New System.Drawing.Point(15, 69)
        Me.txtSalesFile.Name = "txtSalesFile"
        Me.txtSalesFile.ReadOnly = True
        Me.txtSalesFile.Size = New System.Drawing.Size(302, 20)
        Me.txtSalesFile.TabIndex = 1
        '
        'btnConvert
        '
        Me.btnConvert.Location = New System.Drawing.Point(286, 140)
        Me.btnConvert.Name = "btnConvert"
        Me.btnConvert.Size = New System.Drawing.Size(75, 30)
        Me.btnConvert.TabIndex = 2
        Me.btnConvert.Text = "Convert"
        Me.btnConvert.UseVisualStyleBackColor = True
        '
        'txtDest
        '
        Me.txtDest.BackColor = System.Drawing.Color.White
        Me.txtDest.Location = New System.Drawing.Point(15, 110)
        Me.txtDest.Name = "txtDest"
        Me.txtDest.ReadOnly = True
        Me.txtDest.Size = New System.Drawing.Size(302, 20)
        Me.txtDest.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 94)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Destination"
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(323, 63)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(38, 30)
        Me.btnBrowse.TabIndex = 0
        Me.btnBrowse.Text = "..."
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'btnBrowse2
        '
        Me.btnBrowse2.Location = New System.Drawing.Point(323, 104)
        Me.btnBrowse2.Name = "btnBrowse2"
        Me.btnBrowse2.Size = New System.Drawing.Size(38, 30)
        Me.btnBrowse2.TabIndex = 1
        Me.btnBrowse2.Text = "..."
        Me.btnBrowse2.UseVisualStyleBackColor = True
        '
        'ofdOpen
        '
        Me.ofdOpen.Filter = "Excel Files 2000|*.xls"
        '
        'sfdSave
        '
        Me.sfdSave.Filter = "PTU File|*.ptu"
        '
        'cmbWareHouse
        '
        Me.cmbWareHouse.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbWareHouse.FormattingEnabled = True
        Me.cmbWareHouse.Items.AddRange(New Object() {"CHIC-BOY", "COFFEE LOVER", "FOODCOURT", "PBA SPORTS CAFE", "PTU KTV", "PTU WAVE RESTO"})
        Me.cmbWareHouse.Location = New System.Drawing.Point(15, 28)
        Me.cmbWareHouse.Name = "cmbWareHouse"
        Me.cmbWareHouse.Size = New System.Drawing.Size(139, 21)
        Me.cmbWareHouse.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(14, 10)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Store"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(378, 186)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cmbWareHouse)
        Me.Controls.Add(Me.btnBrowse2)
        Me.Controls.Add(Me.btnBrowse)
        Me.Controls.Add(Me.txtDest)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnConvert)
        Me.Controls.Add(Me.txtSalesFile)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Converter"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtSalesFile As System.Windows.Forms.TextBox
    Friend WithEvents btnConvert As System.Windows.Forms.Button
    Friend WithEvents txtDest As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents btnBrowse2 As System.Windows.Forms.Button
    Friend WithEvents ofdOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents sfdSave As System.Windows.Forms.SaveFileDialog
    Friend WithEvents cmbWareHouse As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label

End Class

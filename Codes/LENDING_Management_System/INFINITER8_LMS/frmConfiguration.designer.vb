<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConfiguration
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmConfiguration))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label9 = New System.Windows.Forms.Label
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.btnCloseArea = New System.Windows.Forms.Button
        Me.btnDeleteArea = New System.Windows.Forms.Button
        Me.btnAddArea = New System.Windows.Forms.Button
        Me.txtArea = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.btnCloseCollector = New System.Windows.Forms.Button
        Me.btnDeleteCollector = New System.Windows.Forms.Button
        Me.btnAddCollector = New System.Windows.Forms.Button
        Me.txtCollector = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.ListBox2 = New System.Windows.Forms.ListBox
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.btnCloseInterest = New System.Windows.Forms.Button
        Me.btnDeleteInterest = New System.Windows.Forms.Button
        Me.btnAddInterest = New System.Windows.Forms.Button
        Me.txtInterest = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.ListBox3 = New System.Windows.Forms.ListBox
        Me.Panel1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel1.BackgroundImage = CType(resources.GetObject("Panel1.BackgroundImage"), System.Drawing.Image)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Location = New System.Drawing.Point(-3, -1)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(491, 54)
        Me.Panel1.TabIndex = 1
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Font = New System.Drawing.Font("Britannic Bold", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Teal
        Me.Label9.Location = New System.Drawing.Point(8, 11)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(220, 30)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "CONFIGURATIONS"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabControl1.Location = New System.Drawing.Point(4, 59)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(467, 280)
        Me.TabControl1.TabIndex = 2
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.btnCloseArea)
        Me.TabPage1.Controls.Add(Me.btnDeleteArea)
        Me.TabPage1.Controls.Add(Me.btnAddArea)
        Me.TabPage1.Controls.Add(Me.txtArea)
        Me.TabPage1.Controls.Add(Me.Label6)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.ListBox1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(459, 254)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "AREA"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'btnCloseArea
        '
        Me.btnCloseArea.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.btnCloseArea.BackgroundImage = CType(resources.GetObject("btnCloseArea.BackgroundImage"), System.Drawing.Image)
        Me.btnCloseArea.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnCloseArea.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnCloseArea.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnCloseArea.Location = New System.Drawing.Point(380, 215)
        Me.btnCloseArea.Name = "btnCloseArea"
        Me.btnCloseArea.Size = New System.Drawing.Size(73, 28)
        Me.btnCloseArea.TabIndex = 10
        Me.btnCloseArea.Text = "&Close "
        Me.btnCloseArea.UseVisualStyleBackColor = False
        '
        'btnDeleteArea
        '
        Me.btnDeleteArea.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.btnDeleteArea.BackgroundImage = CType(resources.GetObject("btnDeleteArea.BackgroundImage"), System.Drawing.Image)
        Me.btnDeleteArea.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnDeleteArea.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnDeleteArea.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnDeleteArea.Location = New System.Drawing.Point(246, 215)
        Me.btnDeleteArea.Name = "btnDeleteArea"
        Me.btnDeleteArea.Size = New System.Drawing.Size(73, 28)
        Me.btnDeleteArea.TabIndex = 9
        Me.btnDeleteArea.Text = "&Delete"
        Me.btnDeleteArea.UseVisualStyleBackColor = False
        '
        'btnAddArea
        '
        Me.btnAddArea.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.btnAddArea.BackgroundImage = CType(resources.GetObject("btnAddArea.BackgroundImage"), System.Drawing.Image)
        Me.btnAddArea.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnAddArea.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnAddArea.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnAddArea.Location = New System.Drawing.Point(170, 215)
        Me.btnAddArea.Name = "btnAddArea"
        Me.btnAddArea.Size = New System.Drawing.Size(73, 28)
        Me.btnAddArea.TabIndex = 8
        Me.btnAddArea.Text = "&Add"
        Me.btnAddArea.UseVisualStyleBackColor = False
        '
        'txtArea
        '
        Me.txtArea.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtArea.Location = New System.Drawing.Point(182, 28)
        Me.txtArea.Name = "txtArea"
        Me.txtArea.Size = New System.Drawing.Size(171, 21)
        Me.txtArea.TabIndex = 7
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(133, 31)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(43, 13)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "Area :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 14)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "AREAS :"
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(6, 28)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(120, 212)
        Me.ListBox1.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.btnCloseCollector)
        Me.TabPage2.Controls.Add(Me.btnDeleteCollector)
        Me.TabPage2.Controls.Add(Me.btnAddCollector)
        Me.TabPage2.Controls.Add(Me.txtCollector)
        Me.TabPage2.Controls.Add(Me.Label5)
        Me.TabPage2.Controls.Add(Me.Label2)
        Me.TabPage2.Controls.Add(Me.ListBox2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(459, 254)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "COLLECTOR"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'btnCloseCollector
        '
        Me.btnCloseCollector.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.btnCloseCollector.BackgroundImage = CType(resources.GetObject("btnCloseCollector.BackgroundImage"), System.Drawing.Image)
        Me.btnCloseCollector.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnCloseCollector.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnCloseCollector.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnCloseCollector.Location = New System.Drawing.Point(380, 215)
        Me.btnCloseCollector.Name = "btnCloseCollector"
        Me.btnCloseCollector.Size = New System.Drawing.Size(73, 28)
        Me.btnCloseCollector.TabIndex = 13
        Me.btnCloseCollector.Text = "&Close "
        Me.btnCloseCollector.UseVisualStyleBackColor = False
        '
        'btnDeleteCollector
        '
        Me.btnDeleteCollector.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.btnDeleteCollector.BackgroundImage = CType(resources.GetObject("btnDeleteCollector.BackgroundImage"), System.Drawing.Image)
        Me.btnDeleteCollector.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnDeleteCollector.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnDeleteCollector.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnDeleteCollector.Location = New System.Drawing.Point(246, 215)
        Me.btnDeleteCollector.Name = "btnDeleteCollector"
        Me.btnDeleteCollector.Size = New System.Drawing.Size(73, 28)
        Me.btnDeleteCollector.TabIndex = 12
        Me.btnDeleteCollector.Text = "&Delete"
        Me.btnDeleteCollector.UseVisualStyleBackColor = False
        '
        'btnAddCollector
        '
        Me.btnAddCollector.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.btnAddCollector.BackgroundImage = CType(resources.GetObject("btnAddCollector.BackgroundImage"), System.Drawing.Image)
        Me.btnAddCollector.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnAddCollector.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnAddCollector.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnAddCollector.Location = New System.Drawing.Point(170, 215)
        Me.btnAddCollector.Name = "btnAddCollector"
        Me.btnAddCollector.Size = New System.Drawing.Size(73, 28)
        Me.btnAddCollector.TabIndex = 11
        Me.btnAddCollector.Text = "&Add"
        Me.btnAddCollector.UseVisualStyleBackColor = False
        '
        'txtCollector
        '
        Me.txtCollector.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCollector.Location = New System.Drawing.Point(206, 27)
        Me.txtCollector.Name = "txtCollector"
        Me.txtCollector.Size = New System.Drawing.Size(177, 21)
        Me.txtCollector.TabIndex = 7
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(133, 30)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(67, 13)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "Collector :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(3, 10)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(99, 14)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "COLLECTORS :"
        '
        'ListBox2
        '
        Me.ListBox2.FormattingEnabled = True
        Me.ListBox2.Location = New System.Drawing.Point(6, 27)
        Me.ListBox2.Name = "ListBox2"
        Me.ListBox2.Size = New System.Drawing.Size(120, 212)
        Me.ListBox2.TabIndex = 2
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.btnCloseInterest)
        Me.TabPage3.Controls.Add(Me.btnDeleteInterest)
        Me.TabPage3.Controls.Add(Me.btnAddInterest)
        Me.TabPage3.Controls.Add(Me.txtInterest)
        Me.TabPage3.Controls.Add(Me.Label4)
        Me.TabPage3.Controls.Add(Me.Label3)
        Me.TabPage3.Controls.Add(Me.ListBox3)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(459, 254)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "INTEREST"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'btnCloseInterest
        '
        Me.btnCloseInterest.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.btnCloseInterest.BackgroundImage = CType(resources.GetObject("btnCloseInterest.BackgroundImage"), System.Drawing.Image)
        Me.btnCloseInterest.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnCloseInterest.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnCloseInterest.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnCloseInterest.Location = New System.Drawing.Point(380, 215)
        Me.btnCloseInterest.Name = "btnCloseInterest"
        Me.btnCloseInterest.Size = New System.Drawing.Size(73, 28)
        Me.btnCloseInterest.TabIndex = 13
        Me.btnCloseInterest.Text = "&Close "
        Me.btnCloseInterest.UseVisualStyleBackColor = False
        '
        'btnDeleteInterest
        '
        Me.btnDeleteInterest.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.btnDeleteInterest.BackgroundImage = CType(resources.GetObject("btnDeleteInterest.BackgroundImage"), System.Drawing.Image)
        Me.btnDeleteInterest.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnDeleteInterest.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnDeleteInterest.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnDeleteInterest.Location = New System.Drawing.Point(246, 215)
        Me.btnDeleteInterest.Name = "btnDeleteInterest"
        Me.btnDeleteInterest.Size = New System.Drawing.Size(73, 28)
        Me.btnDeleteInterest.TabIndex = 12
        Me.btnDeleteInterest.Text = "&Delete"
        Me.btnDeleteInterest.UseVisualStyleBackColor = False
        '
        'btnAddInterest
        '
        Me.btnAddInterest.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.btnAddInterest.BackgroundImage = CType(resources.GetObject("btnAddInterest.BackgroundImage"), System.Drawing.Image)
        Me.btnAddInterest.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnAddInterest.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnAddInterest.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnAddInterest.Location = New System.Drawing.Point(170, 215)
        Me.btnAddInterest.Name = "btnAddInterest"
        Me.btnAddInterest.Size = New System.Drawing.Size(73, 28)
        Me.btnAddInterest.TabIndex = 11
        Me.btnAddInterest.Text = "&Add"
        Me.btnAddInterest.UseVisualStyleBackColor = False
        '
        'txtInterest
        '
        Me.txtInterest.Location = New System.Drawing.Point(229, 28)
        Me.txtInterest.Name = "txtInterest"
        Me.txtInterest.Size = New System.Drawing.Size(126, 21)
        Me.txtInterest.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(132, 31)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(91, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Interest Rate :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(3, 14)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(119, 14)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "INTEREST RATES :"
        '
        'ListBox3
        '
        Me.ListBox3.FormattingEnabled = True
        Me.ListBox3.Location = New System.Drawing.Point(6, 31)
        Me.ListBox3.Name = "ListBox3"
        Me.ListBox3.Size = New System.Drawing.Size(120, 212)
        Me.ListBox3.TabIndex = 2
        '
        'frmConfiguration
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(474, 346)
        Me.ControlBox = False
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmConfiguration"
        Me.Opacity = 0.9
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ListBox2 As System.Windows.Forms.ListBox
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ListBox3 As System.Windows.Forms.ListBox
    Friend WithEvents txtArea As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtCollector As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtInterest As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnCloseArea As System.Windows.Forms.Button
    Friend WithEvents btnDeleteArea As System.Windows.Forms.Button
    Friend WithEvents btnAddArea As System.Windows.Forms.Button
    Friend WithEvents btnCloseCollector As System.Windows.Forms.Button
    Friend WithEvents btnDeleteCollector As System.Windows.Forms.Button
    Friend WithEvents btnAddCollector As System.Windows.Forms.Button
    Friend WithEvents btnCloseInterest As System.Windows.Forms.Button
    Friend WithEvents btnDeleteInterest As System.Windows.Forms.Button
    Friend WithEvents btnAddInterest As System.Windows.Forms.Button
End Class

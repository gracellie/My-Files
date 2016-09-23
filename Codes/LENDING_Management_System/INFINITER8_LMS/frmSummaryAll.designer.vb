<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSummaryAll
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSummaryAll))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.chkFullyPaid = New System.Windows.Forms.CheckBox()
        Me.chkByMonth = New System.Windows.Forms.CheckBox()
        Me.chkUnPaid = New System.Windows.Forms.CheckBox()
        Me.chkByDate = New System.Windows.Forms.CheckBox()
        Me.dtpWeek = New System.Windows.Forms.DateTimePicker()
        Me.cmbTerm = New System.Windows.Forms.ComboBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.cmbArea = New System.Windows.Forms.ComboBox()
        Me.txtLastname = New System.Windows.Forms.TextBox()
        Me.cmbCollector = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.dgw = New System.Windows.Forms.DataGridView()
        Me.collastname = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DateLoan = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Address = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Area = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Weekly = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Collector = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IntEarned = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Status = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DateofFullPayment = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.lblTotalNoofLoans = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblPrincipal = New System.Windows.Forms.Label()
        Me.lblInterestEarned = New System.Windows.Forms.Label()
        Me.lblOustandingBal = New System.Windows.Forms.Label()
        Me.lblAccruedInterest = New System.Windows.Forms.Label()
        Me.lblOustandingCap = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Panel2.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgw, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel2.BackgroundImage = CType(resources.GetObject("Panel2.BackgroundImage"), System.Drawing.Image)
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel2.Controls.Add(Me.Button1)
        Me.Panel2.Controls.Add(Me.btnPrint)
        Me.Panel2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel2.Location = New System.Drawing.Point(-16, 619)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1051, 58)
        Me.Panel2.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Transparent
        Me.Button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.DarkCyan
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(959, 11)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 39)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "&Close "
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.UseVisualStyleBackColor = False
        '
        'btnPrint
        '
        Me.btnPrint.BackColor = System.Drawing.Color.Transparent
        Me.btnPrint.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnPrint.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnPrint.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnPrint.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.ForeColor = System.Drawing.Color.DarkCyan
        Me.btnPrint.Image = CType(resources.GetObject("btnPrint.Image"), System.Drawing.Image)
        Me.btnPrint.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPrint.Location = New System.Drawing.Point(878, 11)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 39)
        Me.btnPrint.TabIndex = 0
        Me.btnPrint.Text = "&Print  "
        Me.btnPrint.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnPrint.UseVisualStyleBackColor = False
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox2.Controls.Add(Me.chkFullyPaid)
        Me.GroupBox2.Controls.Add(Me.chkByMonth)
        Me.GroupBox2.Controls.Add(Me.chkUnPaid)
        Me.GroupBox2.Controls.Add(Me.chkByDate)
        Me.GroupBox2.Controls.Add(Me.dtpWeek)
        Me.GroupBox2.Controls.Add(Me.cmbTerm)
        Me.GroupBox2.Controls.Add(Me.Label14)
        Me.GroupBox2.Controls.Add(Me.cmbArea)
        Me.GroupBox2.Controls.Add(Me.txtLastname)
        Me.GroupBox2.Controls.Add(Me.cmbCollector)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(6, 71)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(167, 267)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Filter By"
        '
        'chkFullyPaid
        '
        Me.chkFullyPaid.AutoSize = True
        Me.chkFullyPaid.Location = New System.Drawing.Point(88, 143)
        Me.chkFullyPaid.Name = "chkFullyPaid"
        Me.chkFullyPaid.Size = New System.Drawing.Size(80, 17)
        Me.chkFullyPaid.TabIndex = 13
        Me.chkFullyPaid.Text = "Fully Paid"
        Me.chkFullyPaid.UseVisualStyleBackColor = True
        '
        'chkByMonth
        '
        Me.chkByMonth.AutoSize = True
        Me.chkByMonth.Location = New System.Drawing.Point(90, 174)
        Me.chkByMonth.Name = "chkByMonth"
        Me.chkByMonth.Size = New System.Drawing.Size(79, 17)
        Me.chkByMonth.TabIndex = 13
        Me.chkByMonth.Text = "By Month"
        Me.chkByMonth.UseVisualStyleBackColor = True
        '
        'chkUnPaid
        '
        Me.chkUnPaid.AutoSize = True
        Me.chkUnPaid.Location = New System.Drawing.Point(14, 142)
        Me.chkUnPaid.Name = "chkUnPaid"
        Me.chkUnPaid.Size = New System.Drawing.Size(69, 17)
        Me.chkUnPaid.TabIndex = 13
        Me.chkUnPaid.Text = "Unpaid "
        Me.chkUnPaid.UseVisualStyleBackColor = True
        '
        'chkByDate
        '
        Me.chkByDate.AutoSize = True
        Me.chkByDate.Location = New System.Drawing.Point(14, 173)
        Me.chkByDate.Name = "chkByDate"
        Me.chkByDate.Size = New System.Drawing.Size(72, 17)
        Me.chkByDate.TabIndex = 13
        Me.chkByDate.Text = "By Date"
        Me.chkByDate.UseVisualStyleBackColor = True
        '
        'dtpWeek
        '
        Me.dtpWeek.CustomFormat = "MM/dd/yyyy"
        Me.dtpWeek.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpWeek.Location = New System.Drawing.Point(16, 195)
        Me.dtpWeek.Name = "dtpWeek"
        Me.dtpWeek.Size = New System.Drawing.Size(129, 21)
        Me.dtpWeek.TabIndex = 12
        '
        'cmbTerm
        '
        Me.cmbTerm.FormattingEnabled = True
        Me.cmbTerm.Items.AddRange(New Object() {"All", "Daily", "Weekly", "Monthly", "Semi Monthly"})
        Me.cmbTerm.Location = New System.Drawing.Point(16, 109)
        Me.cmbTerm.Name = "cmbTerm"
        Me.cmbTerm.Size = New System.Drawing.Size(130, 21)
        Me.cmbTerm.TabIndex = 11
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(5, 93)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(94, 13)
        Me.Label14.TabIndex = 10
        Me.Label14.Text = "Payment Term "
        '
        'cmbArea
        '
        Me.cmbArea.FormattingEnabled = True
        Me.cmbArea.Location = New System.Drawing.Point(16, 69)
        Me.cmbArea.Name = "cmbArea"
        Me.cmbArea.Size = New System.Drawing.Size(131, 21)
        Me.cmbArea.TabIndex = 5
        '
        'txtLastname
        '
        Me.txtLastname.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtLastname.Location = New System.Drawing.Point(16, 238)
        Me.txtLastname.Name = "txtLastname"
        Me.txtLastname.Size = New System.Drawing.Size(129, 21)
        Me.txtLastname.TabIndex = 1
        '
        'cmbCollector
        '
        Me.cmbCollector.FormattingEnabled = True
        Me.cmbCollector.Location = New System.Drawing.Point(16, 33)
        Me.cmbCollector.Name = "cmbCollector"
        Me.cmbCollector.Size = New System.Drawing.Size(130, 21)
        Me.cmbCollector.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 222)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(62, 13)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "Lastname"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(4, 53)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(34, 13)
        Me.Label10.TabIndex = 0
        Me.Label10.Text = "Area"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(4, 17)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(58, 13)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Collector"
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.dgw)
        Me.GroupBox1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(179, 72)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(842, 541)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        '
        'dgw
        '
        Me.dgw.AllowUserToAddRows = False
        Me.dgw.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FloralWhite
        Me.dgw.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgw.BackgroundColor = System.Drawing.Color.White
        Me.dgw.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.dgw.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlLightLight
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgw.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.dgw.ColumnHeadersHeight = 40
        Me.dgw.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.collastname, Me.DateLoan, Me.Address, Me.Area, Me.Weekly, Me.Collector, Me.IntEarned, Me.Status, Me.DateofFullPayment})
        Me.dgw.Cursor = System.Windows.Forms.Cursors.Hand
        Me.dgw.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgw.EnableHeadersVisualStyles = False
        Me.dgw.GridColor = System.Drawing.Color.PaleTurquoise
        Me.dgw.Location = New System.Drawing.Point(3, 17)
        Me.dgw.MultiSelect = False
        Me.dgw.Name = "dgw"
        Me.dgw.ReadOnly = True
        Me.dgw.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.InactiveCaption
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgw.RowHeadersDefaultCellStyle = DataGridViewCellStyle8
        Me.dgw.RowHeadersVisible = False
        Me.dgw.RowHeadersWidth = 25
        Me.dgw.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridViewCellStyle9.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle9.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.ActiveCaption
        DataGridViewCellStyle9.SelectionForeColor = System.Drawing.Color.White
        Me.dgw.RowsDefaultCellStyle = DataGridViewCellStyle9
        Me.dgw.RowTemplate.Height = 18
        Me.dgw.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgw.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgw.Size = New System.Drawing.Size(836, 521)
        Me.dgw.TabIndex = 40
        '
        'collastname
        '
        Me.collastname.HeaderText = "Name"
        Me.collastname.Name = "collastname"
        Me.collastname.ReadOnly = True
        Me.collastname.Width = 170
        '
        'DateLoan
        '
        Me.DateLoan.HeaderText = "Date of Loan"
        Me.DateLoan.Name = "DateLoan"
        Me.DateLoan.ReadOnly = True
        Me.DateLoan.Width = 80
        '
        'Address
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Address.DefaultCellStyle = DataGridViewCellStyle3
        Me.Address.HeaderText = "Principal"
        Me.Address.Name = "Address"
        Me.Address.ReadOnly = True
        Me.Address.Width = 70
        '
        'Area
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Area.DefaultCellStyle = DataGridViewCellStyle4
        Me.Area.HeaderText = "Oustanding Cap."
        Me.Area.Name = "Area"
        Me.Area.ReadOnly = True
        Me.Area.Width = 80
        '
        'Weekly
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Weekly.DefaultCellStyle = DataGridViewCellStyle5
        Me.Weekly.HeaderText = "Oustanding Bal. (MV)"
        Me.Weekly.Name = "Weekly"
        Me.Weekly.ReadOnly = True
        Me.Weekly.Width = 80
        '
        'Collector
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Collector.DefaultCellStyle = DataGridViewCellStyle6
        Me.Collector.HeaderText = "Accrued Interest"
        Me.Collector.Name = "Collector"
        Me.Collector.ReadOnly = True
        Me.Collector.Width = 80
        '
        'IntEarned
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.IntEarned.DefaultCellStyle = DataGridViewCellStyle7
        Me.IntEarned.HeaderText = "Interest Earned"
        Me.IntEarned.Name = "IntEarned"
        Me.IntEarned.ReadOnly = True
        Me.IntEarned.Width = 80
        '
        'Status
        '
        Me.Status.HeaderText = "Status"
        Me.Status.Name = "Status"
        Me.Status.ReadOnly = True
        Me.Status.Width = 70
        '
        'DateofFullPayment
        '
        Me.DateofFullPayment.HeaderText = "Date of Full Payment"
        Me.DateofFullPayment.Name = "DateofFullPayment"
        Me.DateofFullPayment.ReadOnly = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel1.BackgroundImage = CType(resources.GetObject("Panel1.BackgroundImage"), System.Drawing.Image)
        Me.Panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Location = New System.Drawing.Point(-3, -1)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1038, 69)
        Me.Panel1.TabIndex = 8
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Font = New System.Drawing.Font("Britannic Bold", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Teal
        Me.Label9.Location = New System.Drawing.Point(11, 10)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(153, 36)
        Me.Label9.TabIndex = 1
        Me.Label9.Text = "SUMMARY"
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox3.Controls.Add(Me.lblTotalNoofLoans)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Controls.Add(Me.lblPrincipal)
        Me.GroupBox3.Controls.Add(Me.lblInterestEarned)
        Me.GroupBox3.Controls.Add(Me.lblOustandingBal)
        Me.GroupBox3.Controls.Add(Me.lblAccruedInterest)
        Me.GroupBox3.Controls.Add(Me.lblOustandingCap)
        Me.GroupBox3.Controls.Add(Me.Label12)
        Me.GroupBox3.Controls.Add(Me.Label8)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(6, 344)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(167, 269)
        Me.GroupBox3.TabIndex = 13
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Totals"
        '
        'lblTotalNoofLoans
        '
        Me.lblTotalNoofLoans.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalNoofLoans.Location = New System.Drawing.Point(18, 235)
        Me.lblTotalNoofLoans.Name = "lblTotalNoofLoans"
        Me.lblTotalNoofLoans.Size = New System.Drawing.Size(130, 20)
        Me.lblTotalNoofLoans.TabIndex = 8
        Me.lblTotalNoofLoans.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(2, 219)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(111, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Total No. Of Loans"
        '
        'lblPrincipal
        '
        Me.lblPrincipal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPrincipal.Location = New System.Drawing.Point(18, 35)
        Me.lblPrincipal.Name = "lblPrincipal"
        Me.lblPrincipal.Size = New System.Drawing.Size(130, 20)
        Me.lblPrincipal.TabIndex = 6
        Me.lblPrincipal.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblInterestEarned
        '
        Me.lblInterestEarned.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblInterestEarned.Location = New System.Drawing.Point(19, 190)
        Me.lblInterestEarned.Name = "lblInterestEarned"
        Me.lblInterestEarned.Size = New System.Drawing.Size(130, 20)
        Me.lblInterestEarned.TabIndex = 6
        Me.lblInterestEarned.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblOustandingBal
        '
        Me.lblOustandingBal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOustandingBal.Location = New System.Drawing.Point(20, 113)
        Me.lblOustandingBal.Name = "lblOustandingBal"
        Me.lblOustandingBal.Size = New System.Drawing.Size(130, 20)
        Me.lblOustandingBal.TabIndex = 6
        Me.lblOustandingBal.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAccruedInterest
        '
        Me.lblAccruedInterest.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAccruedInterest.Location = New System.Drawing.Point(20, 154)
        Me.lblAccruedInterest.Name = "lblAccruedInterest"
        Me.lblAccruedInterest.Size = New System.Drawing.Size(130, 20)
        Me.lblAccruedInterest.TabIndex = 6
        Me.lblAccruedInterest.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblOustandingCap
        '
        Me.lblOustandingCap.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOustandingCap.Location = New System.Drawing.Point(21, 77)
        Me.lblOustandingCap.Name = "lblOustandingCap"
        Me.lblOustandingCap.Size = New System.Drawing.Size(130, 20)
        Me.lblOustandingCap.TabIndex = 6
        Me.lblOustandingCap.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(5, 19)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(55, 13)
        Me.Label12.TabIndex = 5
        Me.Label12.Text = "Principal"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 174)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(96, 13)
        Me.Label8.TabIndex = 5
        Me.Label8.Text = "Interest Earned"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(2, 97)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(155, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Outstanding Balance (MV)"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(5, 138)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(102, 13)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "Accrued Interest"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 61)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(115, 13)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Oustanding Capital"
        '
        'frmSummaryAll
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1031, 676)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.GroupBox3)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmSummaryAll"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel2.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.dgw, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmbCollector As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dgw As System.Windows.Forms.DataGridView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents lblOustandingCap As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lblPrincipal As System.Windows.Forms.Label
    Friend WithEvents lblInterestEarned As System.Windows.Forms.Label
    Friend WithEvents lblOustandingBal As System.Windows.Forms.Label
    Friend WithEvents lblAccruedInterest As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtLastname As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmbArea As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cmbTerm As System.Windows.Forms.ComboBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents chkFullyPaid As System.Windows.Forms.CheckBox
    Friend WithEvents chkByMonth As System.Windows.Forms.CheckBox
    Friend WithEvents chkUnPaid As System.Windows.Forms.CheckBox
    Friend WithEvents chkByDate As System.Windows.Forms.CheckBox
    Friend WithEvents dtpWeek As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblTotalNoofLoans As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents collastname As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DateLoan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Address As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Area As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Weekly As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Collector As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IntEarned As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Status As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DateofFullPayment As System.Windows.Forms.DataGridViewTextBoxColumn
End Class

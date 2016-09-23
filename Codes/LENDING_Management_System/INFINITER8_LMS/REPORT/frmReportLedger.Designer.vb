<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReportLedger
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
        Me.components = New System.ComponentModel.Container()
        Dim ReportDataSource1 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
        Me.ReportViewer1 = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.infiniter8_dbDataSet = New INFINITER8_LMS.infiniter8_dbDataSet()
        Me.ledgerBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        ' Me.ledgerTableAdapter = New INFINITER8_LMS.infiniter8_dbDataSetTableAdapters.ledgerTableAdapter()
        CType(Me.infiniter8_dbDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ledgerBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ReportViewer1
        '
        Me.ReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        ReportDataSource1.Name = "dsLedger"
        ReportDataSource1.Value = Me.ledgerBindingSource
        Me.ReportViewer1.LocalReport.DataSources.Add(ReportDataSource1)
        Me.ReportViewer1.LocalReport.ReportEmbeddedResource = "INFINITER8_LMS.rptLedger.rdlc"
        Me.ReportViewer1.Location = New System.Drawing.Point(0, 0)
        Me.ReportViewer1.Name = "ReportViewer1"
        Me.ReportViewer1.Size = New System.Drawing.Size(553, 325)
        Me.ReportViewer1.TabIndex = 0
        '
        'infiniter8_dbDataSet
        '
        Me.infiniter8_dbDataSet.DataSetName = "infiniter8_dbDataSet"
        Me.infiniter8_dbDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'ledgerBindingSource
        '
        Me.ledgerBindingSource.DataMember = "ledger"
        Me.ledgerBindingSource.DataSource = Me.infiniter8_dbDataSet
        '
        'ledgerTableAdapter
        '
        ' Me.ledgerTableAdapter.ClearBeforeFill = True
        '
        'frmReportLedger
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(553, 325)
        Me.Controls.Add(Me.ReportViewer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmReportLedger"
        Me.Text = "Borrower's Ledger"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.infiniter8_dbDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ledgerBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ReportViewer1 As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents ledgerBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents infiniter8_dbDataSet As INFINITER8_LMS.infiniter8_dbDataSet
    'Friend WithEvents ledgerTableAdapter As INFINITER8_LMS.infiniter8_dbDataSetTableAdapters.ledgerTableAdapter
End Class

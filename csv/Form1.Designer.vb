<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.btnSelectFile = New System.Windows.Forms.Button()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.txtFilePath = New System.Windows.Forms.TextBox()
        Me.dgvPreview = New System.Windows.Forms.DataGridView()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.txtInvoiceNo = New System.Windows.Forms.TextBox()
        Me.txtDate = New System.Windows.Forms.TextBox()
        Me.txtCompanyName = New System.Windows.Forms.TextBox()
        Me.lblTotalAmount = New System.Windows.Forms.Label()
        Me.dgvMapping = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnSaveFormat = New System.Windows.Forms.Button()
        Me.txtFormatName = New System.Windows.Forms.TextBox()
        Me.cmbSavedFormats = New System.Windows.Forms.ComboBox()
        Me.btnLoadFormat = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.SetUpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SelectDatabaseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.cmbBillType = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.EXT = New System.Windows.Forms.Button()
        CType(Me.dgvPreview, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvMapping, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnSelectFile
        '
        Me.btnSelectFile.Location = New System.Drawing.Point(174, 57)
        Me.btnSelectFile.Name = "btnSelectFile"
        Me.btnSelectFile.Size = New System.Drawing.Size(118, 35)
        Me.btnSelectFile.TabIndex = 0
        Me.btnSelectFile.Text = "Select CSV File"
        Me.btnSelectFile.UseVisualStyleBackColor = True
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(321, 447)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(147, 59)
        Me.btnImport.TabIndex = 1
        Me.btnImport.Text = "Save to Database"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'txtFilePath
        '
        Me.txtFilePath.Location = New System.Drawing.Point(372, 65)
        Me.txtFilePath.Name = "txtFilePath"
        Me.txtFilePath.Size = New System.Drawing.Size(182, 20)
        Me.txtFilePath.TabIndex = 2
        '
        'dgvPreview
        '
        Me.dgvPreview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPreview.Location = New System.Drawing.Point(321, 149)
        Me.dgvPreview.Name = "dgvPreview"
        Me.dgvPreview.Size = New System.Drawing.Size(769, 292)
        Me.dgvPreview.TabIndex = 3
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'txtInvoiceNo
        '
        Me.txtInvoiceNo.Location = New System.Drawing.Point(648, 65)
        Me.txtInvoiceNo.Name = "txtInvoiceNo"
        Me.txtInvoiceNo.Size = New System.Drawing.Size(132, 20)
        Me.txtInvoiceNo.TabIndex = 4
        '
        'txtDate
        '
        Me.txtDate.Location = New System.Drawing.Point(813, 65)
        Me.txtDate.Name = "txtDate"
        Me.txtDate.Size = New System.Drawing.Size(150, 20)
        Me.txtDate.TabIndex = 4
        '
        'txtCompanyName
        '
        Me.txtCompanyName.Location = New System.Drawing.Point(400, 109)
        Me.txtCompanyName.Name = "txtCompanyName"
        Me.txtCompanyName.Size = New System.Drawing.Size(179, 20)
        Me.txtCompanyName.TabIndex = 4
        '
        'lblTotalAmount
        '
        Me.lblTotalAmount.AutoSize = True
        Me.lblTotalAmount.Location = New System.Drawing.Point(28, 65)
        Me.lblTotalAmount.Name = "lblTotalAmount"
        Me.lblTotalAmount.Size = New System.Drawing.Size(39, 13)
        Me.lblTotalAmount.TabIndex = 5
        Me.lblTotalAmount.Text = "Label1"
        '
        'dgvMapping
        '
        Me.dgvMapping.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvMapping.Location = New System.Drawing.Point(16, 149)
        Me.dgvMapping.Name = "dgvMapping"
        Me.dgvMapping.Size = New System.Drawing.Size(264, 415)
        Me.dgvMapping.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 133)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(89, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Set your Columns"
        '
        'btnSaveFormat
        '
        Me.btnSaveFormat.Location = New System.Drawing.Point(174, 600)
        Me.btnSaveFormat.Name = "btnSaveFormat"
        Me.btnSaveFormat.Size = New System.Drawing.Size(106, 45)
        Me.btnSaveFormat.TabIndex = 8
        Me.btnSaveFormat.Text = "Save Format"
        Me.btnSaveFormat.UseVisualStyleBackColor = True
        '
        'txtFormatName
        '
        Me.txtFormatName.Location = New System.Drawing.Point(117, 574)
        Me.txtFormatName.Name = "txtFormatName"
        Me.txtFormatName.Size = New System.Drawing.Size(163, 20)
        Me.txtFormatName.TabIndex = 9
        '
        'cmbSavedFormats
        '
        Me.cmbSavedFormats.FormattingEnabled = True
        Me.cmbSavedFormats.Location = New System.Drawing.Point(31, 27)
        Me.cmbSavedFormats.Name = "cmbSavedFormats"
        Me.cmbSavedFormats.Size = New System.Drawing.Size(188, 21)
        Me.cmbSavedFormats.TabIndex = 10
        '
        'btnLoadFormat
        '
        Me.btnLoadFormat.Location = New System.Drawing.Point(235, 25)
        Me.btnLoadFormat.Name = "btnLoadFormat"
        Me.btnLoadFormat.Size = New System.Drawing.Size(75, 23)
        Me.btnLoadFormat.TabIndex = 8
        Me.btnLoadFormat.Text = "Load Format"
        Me.btnLoadFormat.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(318, 68)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Location"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(560, 68)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(82, 13)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Invoice Number"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(786, 68)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Date"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(318, 112)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(76, 13)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Supplier Name"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(14, 577)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(76, 13)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Formate Name"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SetUpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1370, 24)
        Me.MenuStrip1.TabIndex = 12
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'SetUpToolStripMenuItem
        '
        Me.SetUpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SelectDatabaseToolStripMenuItem})
        Me.SetUpToolStripMenuItem.Name = "SetUpToolStripMenuItem"
        Me.SetUpToolStripMenuItem.Size = New System.Drawing.Size(50, 20)
        Me.SetUpToolStripMenuItem.Text = "SetUp"
        '
        'SelectDatabaseToolStripMenuItem
        '
        Me.SelectDatabaseToolStripMenuItem.Name = "SelectDatabaseToolStripMenuItem"
        Me.SelectDatabaseToolStripMenuItem.Size = New System.Drawing.Size(156, 22)
        Me.SelectDatabaseToolStripMenuItem.Text = "Database Select"
        '
        'cmbBillType
        '
        Me.cmbBillType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbBillType.FormattingEnabled = True
        Me.cmbBillType.Items.AddRange(New Object() {"I", "E"})
        Me.cmbBillType.Location = New System.Drawing.Point(682, 109)
        Me.cmbBillType.Name = "cmbBillType"
        Me.cmbBillType.Size = New System.Drawing.Size(75, 21)
        Me.cmbBillType.TabIndex = 13
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(597, 112)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(79, 13)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Purchase Type"
        '
        'EXT
        '
        Me.EXT.Location = New System.Drawing.Point(982, 109)
        Me.EXT.Name = "EXT"
        Me.EXT.Size = New System.Drawing.Size(108, 37)
        Me.EXT.TabIndex = 15
        Me.EXT.Text = "Exit"
        Me.EXT.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(1370, 749)
        Me.Controls.Add(Me.EXT)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.cmbBillType)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cmbSavedFormats)
        Me.Controls.Add(Me.txtFormatName)
        Me.Controls.Add(Me.btnLoadFormat)
        Me.Controls.Add(Me.btnSaveFormat)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgvMapping)
        Me.Controls.Add(Me.lblTotalAmount)
        Me.Controls.Add(Me.txtCompanyName)
        Me.Controls.Add(Me.txtDate)
        Me.Controls.Add(Me.txtInvoiceNo)
        Me.Controls.Add(Me.dgvPreview)
        Me.Controls.Add(Me.txtFilePath)
        Me.Controls.Add(Me.btnImport)
        Me.Controls.Add(Me.btnSelectFile)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " "
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.dgvPreview, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvMapping, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnSelectFile As Button
    Friend WithEvents btnImport As Button
    Friend WithEvents txtFilePath As TextBox
    Friend WithEvents dgvPreview As DataGridView
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents txtInvoiceNo As TextBox
    Friend WithEvents txtDate As TextBox
    Friend WithEvents txtCompanyName As TextBox
    Friend WithEvents lblTotalAmount As Label
    Friend WithEvents dgvMapping As DataGridView
    Friend WithEvents Label1 As Label
    Friend WithEvents btnSaveFormat As Button
    Friend WithEvents txtFormatName As TextBox
    Friend WithEvents cmbSavedFormats As ComboBox
    Friend WithEvents btnLoadFormat As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents SetUpToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SelectDatabaseToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents cmbBillType As ComboBox
    Friend WithEvents Label7 As Label
    Friend WithEvents EXT As Button
End Class

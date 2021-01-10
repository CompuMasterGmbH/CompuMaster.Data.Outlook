<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
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

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenOutlookPSTOSTToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.QuitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OutlookPSTOSTToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SelectedOutlookFolderForOperationsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripComboBoxOutlookFolderOperationTarget = New System.Windows.Forms.ToolStripComboBox()
        Me.DataGridView = New System.Windows.Forms.DataGridView()
        Me.ExportToExcelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExportToCSVToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.DataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.OutlookPSTOSTToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(800, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenOutlookPSTOSTToolStripMenuItem, Me.ToolStripSeparator1, Me.QuitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'OpenOutlookPSTOSTToolStripMenuItem
        '
        Me.OpenOutlookPSTOSTToolStripMenuItem.Name = "OpenOutlookPSTOSTToolStripMenuItem"
        Me.OpenOutlookPSTOSTToolStripMenuItem.Size = New System.Drawing.Size(197, 22)
        Me.OpenOutlookPSTOSTToolStripMenuItem.Text = "&Open Outlook PST/OST"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(194, 6)
        '
        'QuitToolStripMenuItem
        '
        Me.QuitToolStripMenuItem.Name = "QuitToolStripMenuItem"
        Me.QuitToolStripMenuItem.Size = New System.Drawing.Size(197, 22)
        Me.QuitToolStripMenuItem.Text = "&Quit"
        '
        'OutlookPSTOSTToolStripMenuItem
        '
        Me.OutlookPSTOSTToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SelectedOutlookFolderForOperationsToolStripMenuItem, Me.ToolStripComboBoxOutlookFolderOperationTarget, Me.ExportToExcelToolStripMenuItem, Me.ExportToCSVToolStripMenuItem})
        Me.OutlookPSTOSTToolStripMenuItem.Name = "OutlookPSTOSTToolStripMenuItem"
        Me.OutlookPSTOSTToolStripMenuItem.Size = New System.Drawing.Size(110, 20)
        Me.OutlookPSTOSTToolStripMenuItem.Text = "Out&look PST/OST"
        '
        'SelectedOutlookFolderForOperationsToolStripMenuItem
        '
        Me.SelectedOutlookFolderForOperationsToolStripMenuItem.Name = "SelectedOutlookFolderForOperationsToolStripMenuItem"
        Me.SelectedOutlookFolderForOperationsToolStripMenuItem.Size = New System.Drawing.Size(560, 22)
        Me.SelectedOutlookFolderForOperationsToolStripMenuItem.Text = "Selected Outlook folder for operations"
        '
        'ToolStripComboBoxOutlookFolderOperationTarget
        '
        Me.ToolStripComboBoxOutlookFolderOperationTarget.Name = "ToolStripComboBoxOutlookFolderOperationTarget"
        Me.ToolStripComboBoxOutlookFolderOperationTarget.Size = New System.Drawing.Size(500, 23)
        '
        'DataGridView
        '
        Me.DataGridView.AllowUserToAddRows = False
        Me.DataGridView.AllowUserToDeleteRows = False
        Me.DataGridView.AllowUserToOrderColumns = True
        Me.DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView.Location = New System.Drawing.Point(0, 24)
        Me.DataGridView.Name = "DataGridView"
        Me.DataGridView.ReadOnly = True
        Me.DataGridView.Size = New System.Drawing.Size(800, 426)
        Me.DataGridView.TabIndex = 1
        '
        'ExportToExcelToolStripMenuItem
        '
        Me.ExportToExcelToolStripMenuItem.Name = "ExportToExcelToolStripMenuItem"
        Me.ExportToExcelToolStripMenuItem.Size = New System.Drawing.Size(560, 22)
        Me.ExportToExcelToolStripMenuItem.Text = "Export folder items to Excel"
        '
        'ExportToCSVToolStripMenuItem
        '
        Me.ExportToCSVToolStripMenuItem.Name = "ExportToCSVToolStripMenuItem"
        Me.ExportToCSVToolStripMenuItem.Size = New System.Drawing.Size(560, 22)
        Me.ExportToCSVToolStripMenuItem.Text = "Export folder items to CSV"
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.DataGridView)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "MainForm"
        Me.Text = "Form1"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.DataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents QuitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OpenOutlookPSTOSTToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents OutlookPSTOSTToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SelectedOutlookFolderForOperationsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripComboBoxOutlookFolderOperationTarget As ToolStripComboBox
    Friend WithEvents DataGridView As DataGridView
    Friend WithEvents ExportToExcelToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExportToCSVToolStripMenuItem As ToolStripMenuItem
End Class

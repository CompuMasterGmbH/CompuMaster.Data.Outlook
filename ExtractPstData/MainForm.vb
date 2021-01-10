Imports System.ComponentModel

Public Class MainForm

    Private OutlookPst As New OutlookPstDatabase()

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = My.Application.Info.Title
    End Sub

    Private Sub QuitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub OpenOutlookPstToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenOutlookPstToolStripMenuItem.Click
        Try
            Dim f As New System.Windows.Forms.OpenFileDialog()
            f.CheckFileExists = True
            f.Filter = "Outlook Mailboxes (*.pst)|*.pst"
            If My.Settings.LastOpenedPstFile <> Nothing Then
                f.InitialDirectory = System.IO.Path.GetDirectoryName(My.Settings.LastOpenedPstFile)
                f.FileName = System.IO.Path.GetFileName(My.Settings.LastOpenedPstFile)
            Else
                f.InitialDirectory = My.Application.Info.DirectoryPath
            End If
            If f.ShowDialog() = DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                Me.OutlookPst.OutlookPstFile = f.FileName
                Me.FillOutlookFolderListForOperationTargets()
                Me.UpdateSelectedFolderOperationTarget(Me.OutlookPst.OutlookPstRootFolder)
                My.Settings.LastOpenedPstFile = f.FileName
                My.Settings.Save()
                Me.OutlookPstToolStripMenuItem.Available = True
                Me.Cursor = Cursors.Default
            End If
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MessageBox.Show("ERROR: " & ex.ToString, "Open Outlook PST file", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub UpdateSelectedFolderOperationTarget(folder As CompuMaster.Data.Outlook.Directory)
        Me.OutlookPst.OutlookPstOperationFolder = folder
        Me.ToolStripComboBoxOutlookFolderOperationTarget.SelectedItem = folder.Path
        Me.Text = My.Application.Info.Title & " - " & System.IO.Path.GetFileName(Me.OutlookPst.OutlookPstFile)
        If Me.OutlookPst.OutlookPstOperationFolderPath <> Nothing Then Me.Text &= ":" & Me.OutlookPst.OutlookPstOperationFolderPath
        Me.LoadFolderItems()
    End Sub

    Private Sub FillOutlookFolderListForOperationTargets()
        Me.ToolStripComboBoxOutlookFolderOperationTarget.Items.Clear()
        For Each FolderName As String In Me.OutlookPst.AvailableOutlookFolderPaths
            Me.ToolStripComboBoxOutlookFolderOperationTarget.Items.Add(FolderName)
        Next
    End Sub

    Private Sub MainForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Try
            Me.OutlookPst.QuitOutlookApp()
        Catch
        End Try
    End Sub

    Private Sub ToolStripComboBoxOutlookFolderOperationTarget_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripComboBoxOutlookFolderOperationTarget.SelectedIndexChanged
        Me.Cursor = Cursors.WaitCursor
        UpdateSelectedFolderOperationTarget(Me.OutlookPst.OutlookPstRootFolder.SelectSubFolder(CType(Me.ToolStripComboBoxOutlookFolderOperationTarget.SelectedItem, String), False, False))
        Me.Cursor = Cursors.Default
    End Sub

    Private CurrentFolderItems As DataTable = Nothing

    Private Sub LoadFolderItems()
        Me.CurrentFolderItems = Me.OutlookPst.OutlookPstOperationFolder.ItemsAllAsDataTable()
        Me.DataGridView.DataSource = CurrentFolderItems
    End Sub

    Private Sub ExportToExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportToExcelToolStripMenuItem.Click
        Try
            Dim f As New System.Windows.Forms.SaveFileDialog
            f.DefaultExt = ".xlsx"
            f.Filter = "Microsoft Excel (*.xlsx)|*.xlsx"
            f.InitialDirectory = My.Application.Info.DirectoryPath
            f.FileName = Me.OutlookPst.OutlookPstOperationFolder.DisplayName
            If f.ShowDialog() = DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                If System.IO.File.Exists(f.FileName) Then System.IO.File.Delete(f.FileName)
                CompuMaster.Data.XlsEpplus.WriteDataTableToXlsFileAndFirstSheet(f.FileName, Me.PreparedExportDataTable)
                System.Diagnostics.Process.Start(f.FileName)
                Me.Cursor = Cursors.Default
            End If
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MessageBox.Show("ERROR: " & ex.ToString, "Save to Microsoft Excel XLSX file", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ExportToCSVToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportToCSVToolStripMenuItem.Click
        Try
            Dim f As New System.Windows.Forms.SaveFileDialog
            f.DefaultExt = ".csv"
            f.Filter = "Text/Csv (*.csv; *.txt)|*.csv;*.txt"
            f.InitialDirectory = My.Application.Info.DirectoryPath
            f.FileName = Me.OutlookPst.OutlookPstOperationFolder.DisplayName
            If f.ShowDialog() = DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                If System.IO.File.Exists(f.FileName) Then System.IO.File.Delete(f.FileName)
                CompuMaster.Data.Csv.WriteDataTableToCsvFile(f.FileName, Me.PreparedExportDataTable)
                Me.Cursor = Cursors.Default
            End If
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MessageBox.Show("ERROR: " & ex.ToString, "Save to CSV file", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function PreparedExportDataTable() As DataTable
        Dim Result As DataTable = CompuMaster.Data.DataTables.CreateDataTableClone(Me.CurrentFolderItems)
        'Remove all binary fields
        For ColCounter As Integer = Result.Columns.Count - 1 To 0 Step -1
            If Result.Columns(ColCounter).DataType Is GetType(Byte()) Then
                CompuMaster.Data.DataTables.RemoveColumns(Result, New String() {Result.Columns(ColCounter).ColumnName})
            End If
        Next
        'Replace array columns by their string representation
        For ColCounter As Integer = Result.Columns.Count - 1 To 0 Step -1
            If Result.Columns(ColCounter).DataType.IsArray Then
                'Convert cell content to string represenation
                Const TempColName As String = "TempXXXXXXXXXXXXXXXXXX"
                Result.Columns.Add(TempColName, GetType(String))
                For RowCounter As Integer = 0 To Result.Rows.Count - 1
                    If Not IsDBNull(Result.Rows(RowCounter)(ColCounter)) Then
                        Dim Items As Array = CType(Result.Rows(RowCounter)(ColCounter), Array)
                        Dim ItemsAsStrings As New List(Of String)
                        For ItemCounter As Integer = 0 To Items.Length - 1
                            ItemsAsStrings.Add(Items(ItemCounter).ToString)
                        Next
                        Result.Rows(RowCounter)(TempColName) = String.Join("; ", ItemsAsStrings)
                    End If
                Next
                'Drop old column, rename temp column to origin column name
                Dim OldColName As String = Result.Columns(ColCounter).ColumnName
                CompuMaster.Data.DataTables.RemoveColumns(Result, New String() {OldColName})
                Result.Columns(TempColName).ColumnName = OldColName
            End If
        Next
        Return Result
    End Function

    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.OutlookPst.QuitOutlookApp()
        Catch
        End Try
    End Sub

End Class

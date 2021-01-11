Imports System.ComponentModel

Public Class MainForm

    Private OutlookPst As New OutlookPstDatabase()

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = My.Application.Info.Title
    End Sub

    Private Sub QuitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub OpenOutlookPSTStoreToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenOutlookPSTStoreToolStripMenuItem.Click
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
                Me.FolderToolStripMenuItem.Available = True
                Me.OpenOutlookStoreToolStripMenuItem.Available = False
                Me.Cursor = Cursors.Default
            End If
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MessageBox.Show("ERROR: " & ex.ToString, "Open Outlook PST file", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub UpdateSelectedFolderOperationTarget(folder As CompuMaster.Data.Outlook.Directory)
        Me.OutlookPst.OutlookPstOperationFolder = folder
        Me.Text = My.Application.Info.Title & " - "
        If Me.OutlookPst.OutlookPstFile <> Nothing Then
            Me.Text &= System.IO.Path.GetFileName(Me.OutlookPst.OutlookPstFile)
        ElseIf Me.OutlookPst.OutlookStore.FilePath <> Nothing Then
            Me.Text &= System.IO.Path.GetFileName(Me.OutlookPst.OutlookStore.FilePath)
        Else
            Me.Text &= Me.OutlookPst.OutlookStore.DisplayName
        End If
        If Me.OutlookPst.OutlookPstOperationFolderPath <> Nothing Then Me.Text &= ":" & Me.OutlookPst.OutlookPstOperationFolderPath
        For MyCounter As Integer = 0 To Me.FolderToolStripMenuItem.DropDownItems.Count - 1
            If Me.FolderToolStripMenuItem.DropDownItems(MyCounter).Name = "FolderName:" & folder.Path Then
                CType(Me.FolderToolStripMenuItem.DropDownItems(MyCounter), ToolStripMenuItem).CheckState = CheckState.Checked
            Else
                CType(Me.FolderToolStripMenuItem.DropDownItems(MyCounter), ToolStripMenuItem).CheckState = CheckState.Unchecked
            End If
        Next
        Me.LoadFolderItems()
    End Sub

    Private Sub FillOutlookFolderListForOperationTargets()
        Me.FolderToolStripMenuItem.DropDownItems.Clear()
        For Each FolderName As String In Me.OutlookPst.AvailableOutlookFolderPaths
            If FolderName = "" Then
                Me.FolderToolStripMenuItem.DropDownItems.Add(New ToolStripMenuItem("{Root}", Nothing, AddressOf OpenOutlookFolderToolStripMenuItem_Click, "FolderName:" & FolderName))
            Else
                Me.FolderToolStripMenuItem.DropDownItems.Add(New ToolStripMenuItem(FolderName, Nothing, AddressOf OpenOutlookFolderToolStripMenuItem_Click, "FolderName:" & FolderName))
            End If
        Next
    End Sub

    Private Sub MainForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Try
            Me.OutlookPst.QuitOutlookApp()
        Catch
        End Try
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

    Private Sub ConnectToOutlookToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConnectToOutlookToolStripMenuItem.Click
        Me.Cursor = Cursors.WaitCursor
        Me.OpenOutlookStoreToolStripMenuItem.DropDownItems.Clear()
        For Each OStore As NetOffice.OutlookApi.Store In Me.OutlookPst.OutlookApp.Stores
            Me.OpenOutlookStoreToolStripMenuItem.DropDownItems.Add(New ToolStripMenuItem(OStore.DisplayName, Nothing, AddressOf OpenOutlookStoreToolStripMenuItem_Click, "OutlookStore:" & OStore.StoreID))
        Next
        Me.OpenOutlookStoreToolStripMenuItem.Available = True
        Me.Cursor = Cursors.Default
    End Sub

    Public Sub OpenOutlookStoreToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.WaitCursor
        Dim MenuItem As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim SelectedOutlookStore As NetOffice.OutlookApi.Store = Nothing
        For Each OStore As NetOffice.OutlookApi.Store In Me.OutlookPst.OutlookApp.Stores
            If MenuItem.Name = "OutlookStore:" & OStore.StoreID Then
                SelectedOutlookStore = OStore
            End If
        Next
        If SelectedOutlookStore Is Nothing Then Throw New NullReferenceException("Outlook store not found")
        For MyCounter As Integer = 0 To Me.OpenOutlookStoreToolStripMenuItem.DropDownItems.Count - 1
            If Me.OpenOutlookStoreToolStripMenuItem.DropDownItems(MyCounter).Name = "OutlookStore:" & SelectedOutlookStore.StoreID Then
                CType(Me.OpenOutlookStoreToolStripMenuItem.DropDownItems(MyCounter), ToolStripMenuItem).CheckState = CheckState.Checked
            Else
                CType(Me.OpenOutlookStoreToolStripMenuItem.DropDownItems(MyCounter), ToolStripMenuItem).CheckState = CheckState.Unchecked
            End If
        Next
        Me.OutlookPst.OutlookStore = SelectedOutlookStore
        Me.FillOutlookFolderListForOperationTargets()
        Me.UpdateSelectedFolderOperationTarget(Me.OutlookPst.OutlookPstRootFolder)
        'Me.UpdateSelectedFolderOperationTarget(Me.OutlookPst.OutlookApp.LookupFolder(Me.OutlookPst.OutlookStore, CompuMaster.Data.Outlook.OutlookApp.WellKnownFolderName.Calendar).Directory)
        Me.FolderToolStripMenuItem.Available = True
        Me.Cursor = Cursors.Default
    End Sub

    Public Sub OpenOutlookFolderToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.WaitCursor
        Dim MenuItem As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim SelectedFolderPath As String = Nothing
        If MenuItem.Name.StartsWith("FolderName:") Then
            SelectedFolderPath = MenuItem.Name.Substring("FolderName:".Length)
        End If
        If SelectedFolderPath Is Nothing Then Throw New NullReferenceException("Outlook folder not found")
        UpdateSelectedFolderOperationTarget(Me.OutlookPst.OutlookPstRootFolder.SelectSubFolder(SelectedFolderPath, False, False))
        Me.Cursor = Cursors.Default
    End Sub

End Class

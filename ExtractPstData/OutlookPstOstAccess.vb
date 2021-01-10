Option Explicit On
Option Strict On

Imports CompuMaster.Data.Outlook

Public Class OutlookPstDatabase

    Private _OutlookApp As CompuMaster.Data.Outlook.OutlookApp = Nothing
    Public ReadOnly Property OutlookApp As CompuMaster.Data.Outlook.OutlookApp
        Get
            If _OutlookApp Is Nothing Then
                _OutlookApp = New CompuMaster.Data.Outlook.OutlookApp(12)
            End If
            Return _OutlookApp
        End Get
    End Property

    Private _OutlookPstFile As String
    Public Property OutlookPstFile As String
        Get
            Return _OutlookPstFile
        End Get
        Set(value As String)
            _OutlookPstFile = value
            Me.OutlookPstRootFolderPath = OutlookApp.LookupRootFolder(System.IO.Path.Combine(System.Environment.CurrentDirectory, value))
            Me.OutlookPstOperationFolder = Me.OutlookPstRootFolder
        End Set
    End Property

    Public Property OutlookPstRootFolderPath As CompuMaster.Data.Outlook.FolderPathRepresentation = Nothing

    Public ReadOnly Property OutlookPstRootFolder As CompuMaster.Data.Outlook.Directory
        Get
            Return Me.OutlookPstRootFolderPath.Directory
        End Get
    End Property

    Private _OutlookPstOperationFolderName As String
    Public Property OutlookPstOperationFolderPath As String
        Get
            Return _OutlookPstOperationFolderName
        End Get
        Set(value As String)
            _OutlookPstOperationFolderName = value
            _OutlookPstOperationFolder = Me.OutlookPstRootFolder.SelectSubFolder(value, False, False)
        End Set
    End Property

    Private _OutlookPstOperationFolder As CompuMaster.Data.Outlook.Directory
    Public Property OutlookPstOperationFolder As CompuMaster.Data.Outlook.Directory
        Get
            Return _OutlookPstOperationFolder
        End Get
        Set(value As CompuMaster.Data.Outlook.Directory)
            _OutlookPstOperationFolder = value
            _OutlookPstOperationFolderName = value.Path
        End Set
    End Property

    Public Sub QuitOutlookApp()
        If Me._OutlookApp IsNot Nothing Then Me._OutlookApp.Application.Quit()
    End Sub

    Public Function AvailableOutlookFolderPaths() As List(Of String)
        Dim Result As New List(Of String)
        ForDirectoryAndEachSubDirectory(
                Me.OutlookPstRootFolder,
                Sub(dir As Directory)
                    Result.Add(dir.Path)
                End Sub)
        Return Result
    End Function

    Private Delegate Sub DirectoryAction(dir As Directory)

    Private Sub ForDirectoryAndEachSubDirectory(dir As Directory, actions As DirectoryAction)
        actions(dir)
        For Each dirItem As Directory In dir.SubFolders
            ForDirectoryAndEachSubDirectory(dirItem, actions)
        Next
    End Sub

    Private Sub ForEachSubDirectory(dir As Directory, actions As DirectoryAction)
        For Each dirItem As Directory In dir.SubFolders
            actions(dir)
            ForEachSubDirectory(dirItem, actions)
        Next
    End Sub

End Class

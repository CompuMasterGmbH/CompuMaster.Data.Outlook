Option Explicit On
Option Strict On

Imports CompuMaster.Data.Outlook

Public Class OutlookPstOstAccess

    Private _OutlookApp As CompuMaster.Data.Outlook.OutlookApp = Nothing
    Public ReadOnly Property OutlookApp As CompuMaster.Data.Outlook.OutlookApp
        Get
            If _OutlookApp Is Nothing Then
                _OutlookApp = New CompuMaster.Data.Outlook.OutlookApp(12)
            End If
            Return _OutlookApp
        End Get
    End Property

    Private _OutlookPstOstFile As String
    Public Property OutlookPstOstFile As String
        Get
            Return _OutlookPstOstFile
        End Get
        Set(value As String)
            _OutlookPstOstFile = value
            Me.OutlookPstOstRootFolderPath = OutlookApp.LookupRootFolder(System.IO.Path.Combine(System.Environment.CurrentDirectory, value))
            Me.OutlookPstOstOperationFolder = Me.OutlookPstOstRootFolder
        End Set
    End Property

    Public Property OutlookPstOstRootFolderPath As CompuMaster.Data.Outlook.FolderPathRepresentation = Nothing

    Public ReadOnly Property OutlookPstOstRootFolder As CompuMaster.Data.Outlook.Directory
        Get
            Return Me.OutlookPstOstRootFolderPath.Directory
        End Get
    End Property

    Private _OutlookPstOstOperationFolderName As String
    Public Property OutlookPstOstOperationFolderPath As String
        Get
            Return _OutlookPstOstOperationFolderName
        End Get
        Set(value As String)
            _OutlookPstOstOperationFolderName = value
            _OutlookPstOstOperationFolder = Me.OutlookPstOstRootFolder.SelectSubFolder(value, False, False)
        End Set
    End Property

    Private _OutlookPstOstOperationFolder As CompuMaster.Data.Outlook.Directory
    Public Property OutlookPstOstOperationFolder As CompuMaster.Data.Outlook.Directory
        Get
            Return _OutlookPstOstOperationFolder
        End Get
        Set(value As CompuMaster.Data.Outlook.Directory)
            _OutlookPstOstOperationFolder = value
            _OutlookPstOstOperationFolderName = value.Path
        End Set
    End Property

    Public Sub QuitOutlookApp()
        If Me._OutlookApp IsNot Nothing Then Me._OutlookApp.Application.Quit()
    End Sub

    Public Function AvailableOutlookFolderPaths() As List(Of String)
        Dim Result As New List(Of String)
        ForDirectoryAndEachSubDirectory(
                Me.OutlookPstOstRootFolder,
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

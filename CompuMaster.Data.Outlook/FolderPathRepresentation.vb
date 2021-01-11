Option Strict On
Option Explicit On

Imports System
Imports NetOffice.OutlookApi
Imports System.Net

Namespace CompuMaster.Data.Outlook

    ''' <summary>
    ''' A representation of a folder path/ID
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FolderPathRepresentation

        Private _root As CompuMaster.Data.Outlook.OutlookApp.WellKnownFolderName = Nothing
        'Private _subfolder As String = Nothing
        Private _rootFolderID As String = Nothing
        Private _subFolderID As String = Nothing
        Private _outlookWrapper As CompuMaster.Data.Outlook.OutlookApp = Nothing
        Private _outlookRootFolder As MAPIFolder
        Private _outlookSubFolder As MAPIFolder
        Private _store As NetOffice.OutlookApi.Store

        Public Sub New(ByVal outlook As OutlookApp, store As NetOffice.OutlookApi.Store, rootFolderID As String, ByVal subFolderID As String)
            _outlookWrapper = outlook
            _rootFolderID = rootFolderID
            _subFolderID = subFolderID
            _store = store
        End Sub

        Friend Sub New(ByVal outlook As OutlookApp, store As NetOffice.OutlookApi.Store, ByVal subFolder As MAPIFolder)
            Me.New(outlook, store, store.GetRootFolder, subFolder)
        End Sub

        Friend Sub New(ByVal outlook As OutlookApp, store As NetOffice.OutlookApi.Store, ByVal rootFolder As MAPIFolder, ByVal subFolder As MAPIFolder)
            Me.New(outlook, store, rootFolder.EntryID, subFolder.EntryID)
            _outlookRootFolder = rootFolder
            _outlookSubFolder = subFolder
        End Sub

        'Friend Sub New(ByVal outlook As OutlookApp, ByVal folderID As FolderId)
        '    Me.New(outlook, folderID.UniqueId)
        'End Sub

        Public Sub New(ByVal outlookApplication As OutlookApp, ByVal root As CompuMaster.Data.Outlook.OutlookApp.WellKnownFolderName)
            _root = root
            _outlookWrapper = outlookApplication
        End Sub

        'Public Sub New(ByVal exchange As Exchange2007SP1OrHigher, ByVal root As CompuMaster.Data.MsExchange.Exchange2007SP1OrHigher.WellKnownFolderName, ByVal subfolderName As String)
        '    _root = root
        '    _subfolder = subfolderName
        '    _exchangeWrapper = exchange
        'End Sub

        Public ReadOnly Property RootFolder As MAPIFolder
            Get
                If _outlookRootFolder Is Nothing Then
                    _outlookRootFolder = Me._store.GetRootFolder
                End If
                Return _outlookRootFolder
            End Get
        End Property

        Public ReadOnly Property Folder As MAPIFolder
            Get
                If _outlookSubFolder Is Nothing Then
                    _outlookSubFolder = _outlookWrapper.LookupFolder(Me._store, Me._root).Folder
                End If
                Return _outlookSubFolder
            End Get
        End Property

        ''' <summary>
        ''' The folder ID as used in Exchange
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function FolderID() As String
            If Not _subFolderID Is Nothing Then
                Return _subFolderID
            Else
                Return _outlookWrapper.LookupFolder(Me._store, Me._root).FolderID
            End If
        End Function

        Private _Directory As Directory
        Public ReadOnly Property Directory As Directory
            Get
                If _Directory Is Nothing Then
                    If _subFolderID = _rootFolderID Then
                        'RootFolder
                        _Directory = New Directory(_outlookWrapper, Me.RootFolder, CType(Nothing, Folder))
                    Else
                        'SubFolder of RootFolder
                        _Directory = New Directory(_outlookWrapper, Me.RootFolder, CType(Nothing, Folder)).LookupSubDirectory(_subFolderID)
                    End If
                End If
                Return _Directory
            End Get
        End Property
    End Class

End Namespace
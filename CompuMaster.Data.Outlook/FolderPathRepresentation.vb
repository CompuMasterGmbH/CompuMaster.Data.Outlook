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
        Private _folderID As String = Nothing
        Private _outlookWrapper As CompuMaster.Data.Outlook.OutlookApp = Nothing
        Private _outlookFolder As MAPIFolder
        Private _store As NetOffice.OutlookApi.Store

        Public Sub New(ByVal outlook As OutlookApp, store As NetOffice.OutlookApi.Store, ByVal folderID As String)
            _outlookWrapper = outlook
            _folderID = folderID
            _store = store
        End Sub

        Friend Sub New(ByVal outlook As OutlookApp, store As NetOffice.OutlookApi.Store, ByVal folder As MAPIFolder)
            Me.New(outlook, store, folder.EntryID)
            _outlookFolder = folder
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

        Public ReadOnly Property Folder As MAPIFolder
            Get
                If _outlookFolder Is Nothing Then
                    _outlookFolder = _outlookWrapper.LookupFolder(Me._store, Me._root).Folder
                End If
                Return _outlookFolder
            End Get
        End Property


        ''' <summary>
        ''' The folder ID as used in Exchange
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function FolderID() As String
            If Not _folderID Is Nothing Then
                Return _folderID
            Else
                Return _outlookWrapper.LookupFolder(Me._store, Me._root).FolderID
            End If
        End Function

        Private _Directory As Directory
        Public ReadOnly Property Directory As Directory
            Get
                If _Directory Is Nothing Then

                    _Directory = New Directory(_outlookWrapper, Me.Folder, CType(Nothing, Folder))
                End If
                Return _Directory
            End Get
        End Property
    End Class

End Namespace
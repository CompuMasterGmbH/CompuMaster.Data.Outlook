Option Strict On
Option Explicit On

Imports System
Imports NetOffice.OutlookApi
Imports System.Net

Namespace CompuMaster.Data.Outlook

    ''' <summary>
    ''' This solution works only for Outlook
    ''' </summary>
    ''' <remarks></remarks>
    Public Class OutlookApp

        Public Shared Function CreateFactory(minRequiredMajorVersion As Integer) As OutlookApp
            Return New OutlookApp(minRequiredMajorVersion)
        End Function

        Protected _app As New NetOffice.OutlookApi.Application
        Public ReadOnly Property Application As NetOffice.OutlookApi.Application
            Get
                Return _app
            End Get
        End Property

        Friend ReadOnly Property OutlookVersion As Version
            Get
                Return New Version(_app.Version)
            End Get
        End Property

        Public Sub New(minRequiredMajorVersion As Integer)
            If minRequiredMajorVersion < 12 Then Throw New NotSupportedException("Outlook application must be at least version 12")
            If Me.OutlookVersion.Major < minRequiredMajorVersion Then Throw New NotSupportedException("Currently installed outlook application version doesn't match with required version " & minRequiredMajorVersion)
        End Sub

        ''' <summary>
        ''' The list of well known folder names in Exchange
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum WellKnownFolderName As Integer
            'ArchiveDeletedItems = 22
            'ArchiveMsgFolderRoot = 21
            'ArchiveRecoverableItemsDeletions = 24
            'ArchiveRecoverableItemsPurges = 26
            'ArchiveRecoverableItemsRoot = 23
            'ArchiveRecoverableItemsVersions = 25
            'ArchiveRoot = 20
            'Calendar = 0
            'Contacts = 1
            'DeletedItems = 2
            'Drafts = 3
            'Inbox = 4
            'Journal = 5
            'JunkEmail = 13
            'MsgFolderRoot = 10
            'Notes = 6
            'Outbox = 7
            'PublicFoldersRoot = 11
            'RecoverableItemsDeletions = 17
            'RecoverableItemsPurges = 19
            'RecoverableItemsRoot = 16
            'RecoverableItemsVersions = 18
            Root = 12
            'SearchFolders = 14
            'SentItems = 8
            'Tasks = 9
            'VoiceMail = 15
        End Enum

        '''' <summary>
        '''' Send an e-mail message
        '''' </summary>
        '''' <param name="subject"></param>
        '''' <param name="bodyPlainText"></param>
        '''' <param name="recipientsTo"></param>
        '''' <param name="recipientsCc"></param>
        '''' <param name="recipientsBcc"></param>
        '''' <remarks></remarks>
        'Public Sub SendMail(ByVal subject As String, ByVal bodyPlainText As String, ByVal recipientsTo As Recipient(), ByVal recipientsCc As Recipient(), ByVal recipientsBcc As Recipient())
        '    CreateMessage(subject, bodyPlainText, String.Empty, recipientsTo, recipientsCc, recipientsBcc, Nothing).SendAndSaveCopy()
        'End Sub

        '''' <summary>
        '''' Send an e-mail message
        '''' </summary>
        '''' <param name="subject"></param>
        '''' <param name="bodyPlainText"></param>
        '''' <param name="bodyHtml"></param>
        '''' <param name="recipientsTo"></param>
        '''' <param name="recipientsCc"></param>
        '''' <param name="recipientsBcc"></param>
        '''' <remarks></remarks>
        'Public Sub SendMail(ByVal subject As String, ByVal bodyPlainText As String, ByVal bodyHtml As String, ByVal recipientsTo As Recipient(), ByVal recipientsCc As Recipient(), ByVal recipientsBcc As Recipient())
        '    CreateMessage(subject, bodyPlainText, bodyHtml, recipientsTo, recipientsCc, recipientsBcc, Nothing).SendAndSaveCopy()
        'End Sub

        '''' <summary>
        '''' Send an e-mail message with attachment
        '''' </summary>
        '''' <param name="subject"></param>
        '''' <param name="bodyPlainText"></param>
        '''' <param name="bodyHtml"></param>
        '''' <param name="recipientsTo"></param>
        '''' <param name="recipientsCc"></param>
        '''' <param name="recipientsBcc"></param>
        '''' <param name="attachment"></param>
        '''' <remarks></remarks>
        'Public Sub SendMail(ByVal subject As String, ByVal bodyPlainText As String, ByVal bodyHtml As String, ByVal recipientsTo As Recipient(), ByVal recipientsCc As Recipient(), ByVal recipientsBcc As Recipient(), ByVal attachment() As EMailAttachment)
        '    CreateMessage(subject, bodyPlainText, bodyHtml, recipientsTo, recipientsCc, recipientsBcc, attachment).SendAndSaveCopy()
        'End Sub

        '''' <summary>
        '''' Create a new e-mail message
        '''' </summary>
        '''' <param name="subject"></param>
        '''' <param name="bodyPlainText"></param>
        '''' <param name="bodyHtml"></param>
        '''' <param name="recipientsTo"></param>
        '''' <param name="recipientsCc"></param>
        '''' <param name="recipientsBcc"></param>
        '''' <param name="attachment"></param>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Private Function CreateMessage(ByVal subject As String, ByVal bodyPlainText As String, ByVal bodyHtml As String, ByVal recipientsTo As Recipient(), ByVal recipientsCc As Recipient(), ByVal recipientsBcc As Recipient(), ByVal attachment() As EMailAttachment) As EmailMessage
        '    If bodyPlainText = Nothing And bodyHtml = Nothing Then
        '        Throw New ArgumentNullException("plain text or html body required")
        '    ElseIf Not (bodyPlainText = Nothing Xor bodyHtml = Nothing) Then
        '        Throw New ArgumentException("either plain text or html body required, but not both")
        '    End If
        '    If recipientsTo Is Nothing Then recipientsTo = New Recipient() {}
        '    If recipientsCc Is Nothing Then recipientsCc = New Recipient() {}
        '    If recipientsBcc Is Nothing Then recipientsBcc = New Recipient() {}
        '    'Create the e-mail message, set its properties, and send it to user2@contoso.com, saving a copy to the Sent Items folder. 
        '    Dim message As New EmailMessage(Me.CreateConfiguredExchangeService())
        '    message.Subject = subject
        '    If bodyHtml <> Nothing Then
        '        message.Body = New MessageBody(BodyType.HTML, bodyHtml)
        '    Else
        '        message.Body = New MessageBody(BodyType.Text, bodyPlainText)
        '    End If
        '    For Each recipient As Recipient In recipientsTo
        '        If recipient.Name = Nothing Then
        '            message.ToRecipients.Add(recipient.EMailAddress)
        '        Else
        '            message.ToRecipients.Add(recipient.Name, recipient.EMailAddress)
        '        End If
        '    Next
        '    For Each recipient As Recipient In recipientsCc
        '        If recipient.Name = Nothing Then
        '            message.CcRecipients.Add(recipient.EMailAddress)
        '        Else
        '            message.CcRecipients.Add(recipient.Name, recipient.EMailAddress)
        '        End If
        '    Next
        '    For Each recipient As Recipient In recipientsBcc
        '        If recipient.Name = Nothing Then
        '            message.BccRecipients.Add(recipient.EMailAddress)
        '        Else
        '            message.BccRecipients.Add(recipient.Name, recipient.EMailAddress)
        '        End If
        '    Next
        '    If Not attachment Is Nothing Then
        '        For Each Item As EMailAttachment In attachment
        '            If Not Item Is Nothing Then
        '                If Item.FilePath <> "" AndAlso Item.FileName = Nothing Then
        '                    message.Attachments.AddFileAttachment(Item.FilePath)
        '                ElseIf Item.FileName <> "" AndAlso Not Item.FileData Is Nothing Then
        '                    message.Attachments.AddFileAttachment(Item.FileName, Item.FileData)
        '                ElseIf Item.FilePath <> "" AndAlso Item.FileName <> "" Then
        '                    message.Attachments.AddFileAttachment(Item.FileName, Item.FilePath)
        '                ElseIf Item.FileName <> "" AndAlso Not Item.FileStream Is Nothing Then
        '                    message.Attachments.AddFileAttachment(Item.FileName, Item.FileStream)
        '                End If
        '            End If
        '        Next
        '    End If
        '    Return message
        'End Function

        '''' <summary>
        '''' Save an e-mail message as draft in the drafts folder
        '''' </summary>
        '''' <param name="subject"></param>
        '''' <param name="bodyPlainText"></param>
        '''' <param name="bodyHtml"></param>
        '''' <param name="recipientsTo"></param>
        '''' <param name="recipientsCc"></param>
        '''' <param name="recipientsBcc"></param>
        '''' <remarks></remarks>
        'Public Function SaveMailAsDraft(ByVal subject As String, ByVal bodyPlainText As String, ByVal bodyHtml As String, ByVal recipientsTo As Recipient(), ByVal recipientsCc As Recipient(), ByVal recipientsBcc As Recipient()) As Uri
        '    Return SaveMailAsDraft(subject, bodyPlainText, bodyHtml, recipientsTo, recipientsCc, recipientsBcc, Nothing, Nothing)
        'End Function

        '''' <summary>
        '''' Save an e-mail message with attachments as draft in the drafts folder
        '''' </summary>
        '''' <param name="subject"></param>
        '''' <param name="bodyPlainText"></param>
        '''' <param name="bodyHtml"></param>
        '''' <param name="recipientsTo"></param>
        '''' <param name="recipientsCc"></param>
        '''' <param name="recipientsBcc"></param>
        '''' <param name="attachment"></param>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Function SaveMailAsDraft(ByVal subject As String, ByVal bodyPlainText As String, ByVal bodyHtml As String, ByVal recipientsTo As Recipient(), ByVal recipientsCc As Recipient(), ByVal recipientsBcc As Recipient(), ByVal attachment() As EMailAttachment) As Uri
        '    Return SaveMailAsDraft(subject, bodyPlainText, bodyHtml, recipientsTo, recipientsCc, recipientsBcc, Nothing, attachment)
        'End Function

        '''' <summary>
        '''' Save an e-mail message as a draft
        '''' </summary>
        '''' <param name="subject"></param>
        '''' <param name="bodyPlainText"></param>
        '''' <param name="bodyHtml"></param>
        '''' <param name="recipientsTo"></param>
        '''' <param name="recipientsCc"></param>
        '''' <param name="recipientsBcc"></param>
        '''' <param name="folder"></param>
        '''' <remarks></remarks>
        'Public Function SaveMailAsDraft(ByVal subject As String, ByVal bodyPlainText As String, ByVal bodyHtml As String, ByVal recipientsTo As Recipient(), ByVal recipientsCc As Recipient(), ByVal recipientsBcc As Recipient(), ByVal folder As CompuMaster.Data.Outlook.FolderPathRepresentation, ByVal attatchment() As EMailAttachment) As Uri
        '    Dim message As EmailMessage = CreateMessage(subject, bodyPlainText, bodyHtml, recipientsTo, recipientsCc, recipientsBcc, attatchment)
        '    If folder Is Nothing Then
        '        message.Save()
        '    Else
        '        message.Save(folder.FolderID)
        '    End If
        '    If Me._exchangeVersion = ExchangeVersion.Exchange2007_SP1 Then
        '        Return Nothing 'Not supported for exchange 2007
        '    Else
        '        'exchange 2010 supports the lookup of a web client url
        '        Dim url As String = message.WebClientEditFormQueryString
        '        Dim uri As New Uri(url)
        '        Throw New NotImplementedException("URL lookup of mail still to be implemented")
        '        Return uri
        '    End If
        'End Function

        '''' <summary>
        '''' Create an appointment
        '''' </summary>
        '''' <param name="subject"></param>
        '''' <param name="location"></param>
        '''' <param name="body"></param>
        '''' <param name="start"></param>
        '''' <param name="duration"></param>
        '''' <returns>The unique ID of the appointment for later reference</returns>
        '''' <remarks></remarks>
        'Public Function CreateAppointment(ByVal subject As String, ByVal location As String, ByVal body As String, ByVal start As DateTime, ByVal duration As TimeSpan) As String
        '    Return CreateMeetingAppointment(subject, location, body, start, duration, New Recipient() {}, New Recipient() {}, New Recipient() {})
        'End Function

        '''' <summary>
        '''' Create a meeting appointment
        '''' </summary>
        '''' <param name="subject"></param>
        '''' <param name="location"></param>
        '''' <param name="body"></param>
        '''' <param name="start"></param>
        '''' <param name="duration"></param>
        '''' <param name="requiredAttendees"></param>
        '''' <param name="optionalAttendees"></param>
        '''' <param name="resources"></param>
        '''' <returns>The unique ID of the appointment for later reference</returns>
        '''' <remarks></remarks>
        'Public Function CreateMeetingAppointment(ByVal subject As String, ByVal location As String, ByVal body As String, ByVal start As DateTime, ByVal duration As TimeSpan, ByVal requiredAttendees As Recipient(), ByVal optionalAttendees As Recipient(), ByVal resources As Recipient()) As String
        '    If start = Nothing Then Throw New ArgumentNullException("start")
        '    If requiredAttendees Is Nothing Then requiredAttendees = New Recipient() {}
        '    If optionalAttendees Is Nothing Then optionalAttendees = New Recipient() {}
        '    If resources Is Nothing Then resources = New Recipient() {}
        '    Dim appointment As New Appointment(Me.CreateConfiguredExchangeService())
        '    appointment.Subject = subject
        '    appointment.Body = body
        '    appointment.Location = location
        '    appointment.Start = start
        '    appointment.End = appointment.Start.Add(duration)
        '    For Each Attendee As Recipient In requiredAttendees
        '        If Attendee.Name = Nothing Then
        '            appointment.RequiredAttendees.Add(Attendee.EMailAddress)
        '        Else
        '            appointment.RequiredAttendees.Add(Attendee.Name, Attendee.EMailAddress)
        '        End If
        '    Next
        '    For Each Attendee As Recipient In optionalAttendees
        '        If Attendee.Name = Nothing Then
        '            appointment.OptionalAttendees.Add(Attendee.EMailAddress)
        '        Else
        '            appointment.OptionalAttendees.Add(Attendee.Name, Attendee.EMailAddress)
        '        End If
        '    Next
        '    For Each Attendee As Recipient In resources
        '        If Attendee.Name = Nothing Then
        '            appointment.Resources.Add(Attendee.EMailAddress)
        '        Else
        '            appointment.Resources.Add(Attendee.Name, Attendee.EMailAddress)
        '        End If
        '    Next
        '    If requiredAttendees.Length = 0 AndAlso optionalAttendees.Length = 0 AndAlso resources.Length = 0 Then
        '        appointment.Save(SendInvitationsMode.SendToNone)
        '    Else
        '        appointment.Save(SendInvitationsMode.SendToAllAndSaveCopy)
        '    End If
        '    Return appointment.Id.UniqueId
        'End Function

        '''' <summary>
        '''' Save a contact
        '''' </summary>
        '''' <param name="contact"></param>
        '''' <param name="folder"></param>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Function SaveContact(ByVal contact As ContactItem, ByVal folder As FolderPathRepresentation) As String
        '    If contact.IsNew AndAlso Not folder Is Nothing Then
        '        'Create new entry in specified folder
        '        contact.Save(folder.FolderID)
        '    ElseIf contact.IsNew AndAlso folder Is Nothing Then
        '        'Create new entry in default folder
        '        contact.Save()
        '    ElseIf folder Is Nothing Then
        '        'Overwrite existing item
        '        contact.Update(ConflictResolutionMode.AutoResolve)
        '    ElseIf Not folder Is Nothing AndAlso contact.ParentFolderId.UniqueId <> folder.FolderID Then
        '        'Save additional item in different folder instead of overwriting existing item
        '        contact.Save(folder.FolderID)
        '    Else
        '        'Overwrite existing item
        '        contact.Update(ConflictResolutionMode.AutoResolve)
        '    End If
        '    Return contact.Id.UniqueId
        'End Function

        '''' <summary>
        '''' Create a new contact
        '''' </summary>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Function CreateNewContact() As ContactItem
        '    Return New Contact(Me.CreateConfiguredExchangeService())
        'End Function

        '''' <summary>
        '''' Load a contact
        '''' </summary>
        '''' <param name="itemID"></param>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Function LoadContactData(ByVal itemID As ItemID) As ContactItem
        '    'Return Contact.Bind(Me.CreateConfiguredExchangeService, itemID)
        'End Function

        '''' <summary>
        '''' Load a contact
        '''' </summary>
        '''' <param name="itemID"></param>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Function LoadContactData(ByVal itemID As String) As Contact
        '    Return Me.LoadContactData(New Microsoft.Exchange.WebServices.Data.ItemId(itemID))
        'End Function

        ''' <summary>
        ''' Well known folder classes
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum FolderClass As Integer
            Undefined = 0
            Generic = 1
            Contacts = 5
            Tasks = 2
            Search = 3
            Calendar = 4
            Notices = 6
            Journal = 7
            Configuration = 8
            Custom = 9
        End Enum

        '''' <summary>
        '''' Lookup the folder class
        '''' </summary>
        '''' <param name="folder"></param>
        '''' <returns>The well known folder class or otherwise Custom for a custom FolderClass name at Exchange</returns>
        '''' <remarks></remarks>
        '<Obsolete("Better use Directory class", False)> Public Function LookupFolderClass(ByVal folder As FolderPathRepresentation) As OutlookApp.FolderClass
        '    Select Case LookupFolderClassName(folder)
        '        Case "IPF.Appointment"
        '            Return OutlookApp.FolderClass.Calendar
        '        Case "IPF.Contact"
        '            Return OutlookApp.FolderClass.Contacts
        '        Case "IPF.Note"
        '            Return OutlookApp.FolderClass.Generic
        '        Case "IPF.Journal"
        '            Return OutlookApp.FolderClass.Journal
        '            'Case ""
        '            '    Return Exchange2007SP1OrHigher.FolderClass.Search
        '        Case "IPF.Task"
        '            Return OutlookApp.FolderClass.Tasks
        '        Case "IPF.StickyNote"
        '            Return OutlookApp.FolderClass.Notices
        '        Case "IPF.Configuration"
        '            Return OutlookApp.FolderClass.Configuration
        '        Case Else
        '            Return FolderClass.Custom
        '    End Select
        'End Function

        '''' <summary>
        '''' Lookup the folder class
        '''' </summary>
        '''' <param name="folder"></param>
        '''' <returns>The real folder class name as defined at Exchange</returns>
        '''' <remarks></remarks>
        '<Obsolete("Better use Directory class", False)> Public Function LookupFolderClassName(ByVal folder As FolderPathRepresentation) As String
        '    Dim lookupfolder As NetOffice.OutlookApi.MAPIFolder ' Microsoft.Exchange.WebServices.Data.Folder
        '    lookupfolder = _app.GetNamespace("MAPIFolders").GetFolderFromID(folder.FolderID)
        '    Return lookupfolder.Class.ToString
        'End Function

        ''' <summary>
        ''' Convert the well known folder class into the official folder class name as Exchange
        ''' </summary>
        ''' <param name="folderClass"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function FolderClassName(ByVal folderClass As FolderClass) As String
            Select Case folderClass
                Case OutlookApp.FolderClass.Calendar
                    Return "IPF.Appointment"
                Case OutlookApp.FolderClass.Contacts
                    Return "IPF.Contact"
                Case OutlookApp.FolderClass.Generic
                    Return "IPF.Note"
                Case OutlookApp.FolderClass.Journal
                    Return "IPF.Journal"
                Case OutlookApp.FolderClass.Search
                    Throw New NotSupportedException("Search classes use custom names, name resolution not supported")
                Case OutlookApp.FolderClass.Tasks
                    Return "IPF.Task"
                Case OutlookApp.FolderClass.Notices
                    Return "IPF.StickyNote"
                Case OutlookApp.FolderClass.Configuration
                    Return "IPF.Configuration"
                Case OutlookApp.FolderClass.Custom
                    Throw New NotSupportedException("A custom folder class requires you to use custom folder class name. The purpose of this method is not intended for this folder class type.")
                Case Else
                    Throw New ArgumentOutOfRangeException("folderClass")
            End Select
        End Function

        Private Function LookupOutlookStore(outlookNamespace As NetOffice.OutlookApi._NameSpace, filepath As String) As NetOffice.OutlookApi.Store
            For Each store As Store In outlookNamespace.Stores
                If filepath = store.FilePath Then
                    Return store
                End If
            Next
            Return Nothing
        End Function

        Private ReadOnly Property NamespaceMapi As NetOffice.OutlookApi._NameSpace
            Get
                Static _Result As NetOffice.OutlookApi._NameSpace
                If _Result Is Nothing Then
                    _Result = Me._app.Application.GetNamespace("MAPI")
                End If
                Return _Result
            End Get
        End Property

        Function OpenStore(pstFilePath As String) As NetOffice.OutlookApi.Store
            If LookupOutlookStore(NamespaceMapi, pstFilePath) Is Nothing Then
                Me._app.Session.AddStore(pstFilePath)
            End If
            Dim SourceStore As NetOffice.OutlookApi.Store = LookupOutlookStore(NamespaceMapi, pstFilePath)
            If SourceStore Is Nothing Then
                Throw New System.Exception("Adding store " & pstFilePath & " failed")
            Else
                Return SourceStore
            End If
        End Function

        Public Function LookupRootFolder(pstFilePath As String) As FolderPathRepresentation
            If System.IO.File.Exists(pstFilePath) = False Then
                Throw New System.Exception("File not found: " & pstFilePath)
            ElseIf pstFilePath.ToLowerInvariant.EndsWith(".pst") = False Then
                Throw New System.Exception("File not a .pst file: " & pstFilePath)
            End If
            Dim SourceStore As NetOffice.OutlookApi.Store = OpenStore(pstFilePath)
            Return LookupFolder(SourceStore, WellKnownFolderName.Root)
        End Function

        ''' <summary>
        ''' Lookup a folder path representation based on its directory structure
        ''' </summary>
        ''' <param name="baseFolder"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function LookupFolder(store As NetOffice.OutlookApi.Store, ByVal baseFolder As WellKnownFolderName) As FolderPathRepresentation
            Dim folder As MAPIFolder
            Select Case baseFolder
                Case WellKnownFolderName.Root
                    folder = store.GetRootFolder()
                Case Else
                    Throw New ArgumentException("Invalid baseFolder: " & GetType(WellKnownFolderName).GetEnumName(baseFolder))
            End Select
            '= Microsoft.Exchange.WebServices.Data.Folder.Bind(Me.CreateConfiguredExchangeService, CType(baseFolder, Microsoft.Exchange.WebServices.Data.WellKnownFolderName))
            Return New FolderPathRepresentation(Me, store, folder)
        End Function

        Public Function Stores() As Stores
            Return Me.NamespaceMapi.Stores
        End Function

        Private _DirectorySeparatorChar As Char = "\"c
        ''' <summary>
        ''' The directory separator char which shall be used in all method calls to this class with parameters specifying a directory structure
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>The exchange server supports typical directory separator chars like \, / as a normal character within a directory name. That's why the programmer may want to define his own separator char to support required directories correctly.</remarks>
        Public Property DirectorySeparatorChar() As Char
            Get
                Return _DirectorySeparatorChar
            End Get
            Set(ByVal value As Char)
                _DirectorySeparatorChar = value
            End Set
        End Property

        '''' <summary>
        '''' Enumerates possible matches of mailbox accounts/contacts for the searched name
        '''' </summary>
        '''' <param name="searchedName"></param>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Function ResolveMailboxOrContactNames(ByVal searchedName As String) As Mailbox()
        '    'Identify the mailbox folders to search for potential name resolution matches.
        '    Dim folders As List(Of FolderId) = New List(Of FolderId)
        '    folders.Add(New FolderId(Microsoft.Exchange.WebServices.Data.WellKnownFolderName.Contacts))

        '    'Search for all contact entries in the default mailbox contacts folder and in Active Directory. This results in a call to EWS.
        '    Dim coll As NameResolutionCollection = Me.CreateConfiguredExchangeService.ResolveName(searchedName, folders, ResolveNameSearchLocation.ContactsThenDirectory, False)

        '    Dim Results As New ArrayList
        '    For Each nameRes As NameResolution In coll
        '        Results.Add(nameRes.Mailbox)
        '        Console.WriteLine("Contact name: " + nameRes.Mailbox.Name)
        '        Console.WriteLine("Contact e-mail address: " + nameRes.Mailbox.Address)
        '        Console.WriteLine("Mailbox type: " + nameRes.Mailbox.MailboxType.ToString)
        '    Next
        '    Return CType(Results.ToArray(GetType(Mailbox)), Mailbox())
        'End Function

    End Class

End Namespace
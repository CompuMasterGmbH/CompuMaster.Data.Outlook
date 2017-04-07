Option Explicit On
Option Strict On

Imports CompuMaster.Data.Outlook
Imports CompuMaster.Data.Outlook.OutlookApp

Public Class HlsMsExchangeDataAccess

    Private _ExchangeServer As String

    Public Sub New(exchangeServer As String)
        _ExchangeServer = exchangeServer
    End Sub

    'Private Function SearchFilterFactory(activityDateFrom As Date, activityDateTo As Date, schema As Microsoft.Exchange.WebServices.Data.PropertyDefinition) As Microsoft.Exchange.WebServices.Data.SearchFilter.SearchFilterCollection
    '    Dim searchFilterCollection As New Microsoft.Exchange.WebServices.Data.SearchFilter.SearchFilterCollection(Microsoft.Exchange.WebServices.Data.LogicalOperator.And)
    '    Dim searchFilterEarlierDate As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsGreaterThanOrEqualTo(schema, activityDateFrom)
    '    Dim searchFilterLaterDate As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsLessThanOrEqualTo(schema, activityDateTo)
    '    searchFilterCollection.Add(searchFilterEarlierDate)
    '    searchFilterCollection.Add(searchFilterLaterDate)
    '    Return searchFilterCollection
    'End Function

    '''' <summary>
    '''' CAUTION: returned data must be cloned before further processing
    '''' </summary>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function MsExchangeActivities(activityDateFrom As Date, activityDateTo As Date) As DataTable
    '    'prepare search filter
    '    activityDateFrom = Calendar.DateInformation.BeginOfDay(activityDateFrom)
    '    activityDateTo = Calendar.DateInformation.EndOfDay(activityDateTo, Calendar.DateInformation.Accuracy.Second)
    '    Dim searchFilterCollection As New Microsoft.Exchange.WebServices.Data.SearchFilter.SearchFilterCollection(Microsoft.Exchange.WebServices.Data.LogicalOperator.Or)
    '    'searchFilterCollection.Add(SearchFilterFactory(activityDateFrom, activityDateTo, Microsoft.Exchange.WebServices.Data.ItemSchema.DateTimeCreated))
    '    searchFilterCollection.Add(SearchFilterFactory(activityDateFrom, activityDateTo, Microsoft.Exchange.WebServices.Data.ItemSchema.DateTimeReceived))
    '    searchFilterCollection.Add(SearchFilterFactory(activityDateFrom, activityDateTo, Microsoft.Exchange.WebServices.Data.ItemSchema.DateTimeSent))

    '    'retrieve the list of folders of user's mailbox
    '    Dim e2007 As New CompuMaster.Data.Outlook.OutlookApp()
    '    Dim itemView As New Microsoft.Exchange.WebServices.Data.ItemView(Integer.MaxValue, 0, Microsoft.Exchange.WebServices.Data.OffsetBasePoint.Beginning)
    '    Dim folderMsgRoot As CompuMaster.Data.Outlook.FolderPathRepresentation = e2007.LookupFolder(WellKnownFolderName.MsgFolderRoot)
    '    Dim folderRoot As CompuMaster.Data.Outlook.FolderPathRepresentation = e2007.LookupFolder(WellKnownFolderName.Root)
    '    Dim Items As Item() = Nothing '= folderRoot.Directory.MailboxItems(searchFilterCollection, itemView)

    '    'prepare Result table
    '    Dim Result As New DataTable("MsOutlook")
    '    Result.Columns.Add("DateCreated", GetType(DateTime))
    '    Result.Columns.Add("DateTimeReceived", GetType(DateTime))
    '    Result.Columns.Add("DateTimeSent", GetType(DateTime))
    '    Result.Columns.Add("Subject", GetType(String))
    '    Result.Columns.Add("FromSender", GetType(String))
    '    Result.Columns.Add("DisplayTo", GetType(String))
    '    Result.Columns.Add("DisplayCc", GetType(String))
    '    Result.Columns.Add("MailboxFolder", GetType(String))
    '    For Each entryItem As Item In Items
    '        If entryItem.Subject = "" Then 'AndAlso entryItem.DisplayTo = "" AndAlso entryItem.DisplayCc = "" Then
    '            'ignore / don't display item - might be e.g. a contact recipient cache entry - waste for review usage
    '        Else
    '            Dim row As DataRow = Result.NewRow
    '            'row("DateCreated") = Utils.ValueNotNothingOrDBNull(entryItem.DateTimeCreated)
    '            'row("DateTimeReceived") = Utils.ValueNotNothingOrDBNull(entryItem.DateTimeReceived)
    '            'row("DateTimeSent") = Utils.ValueNotNothingOrDBNull(entryItem.DateTimeSent)
    '            row("Subject") = entryItem.Subject
    '            'Try
    '            '    Dim ExchangeSender As String = entryItem.FromExchangeSender
    '            '    If ExchangeSender <> "" Then
    '            '        If ExchangeSender.Contains("<SMTP:") Then
    '            '            'Sample: From=DiskSpaceMonitor <SMTP:noreply@bomag.com>
    '            '            ExchangeSender = ExchangeSender.Replace("<SMTP:", "<")
    '            '        Else
    '            '            'Sample: From=Jochen Wezel - CompuMaster GmbH <EX:/O=COMPUMASTER GMBH - EMMELSHAUSEN/OU=ERSTE ADMINISTRATIVE GRUPPE/CN=RECIPIENTS/CN=JWEZEL>
    '            '            ExchangeSender = ExchangeSender.Substring(0, ExchangeSender.IndexOf("<"c))
    '            '        End If
    '            '    End If
    '            '    row("FromSender") = ExchangeSender
    '            'Catch ex As Exception
    '            '    row("FromSender") = "{ERROR: " & ex.Message & "}"
    '            'End Try
    '            'row("DisplayTo") = entryItem.DisplayTo
    '            'row("DisplayCc") = entryItem.DisplayCc
    '            Dim DisplayPath As String = entryItem.ParentDirectory.DisplayPath
    '            If DisplayPath.StartsWith(folderMsgRoot.Directory.DisplayPath & "\") Then
    '                DisplayPath = DisplayPath.Substring(Len(folderMsgRoot.Directory.DisplayPath & "\"))
    '            End If
    '            row("MailboxFolder") = DisplayPath
    '            Result.Rows.Add(row)
    '        End If
    '    Next
    '    Return Result
    'End Function

End Class

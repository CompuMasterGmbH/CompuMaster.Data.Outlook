Option Explicit On
Option Strict On

Imports CompuMaster.Data.Outlook
Imports CompuMaster.Data.Outlook.OutlookApp

Module Module1

    Sub Main()
        'Try
        '    Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " Execute TestSuite 'MsExchangeActivities 2016-03 (partly)' (Y/N)?")
        '    If Console.ReadKey().KeyChar.ToString.ToLowerInvariant = "y" Then
        '        Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " AppStart")
        '        Dim he As New HlsMsExchangeDataAccess("server-test-exchange")
        '        Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " BeforeQuery")
        '        Dim t As DataTable ' = he.MsExchangeActivities(New Date(2016, 3, 3), New Date(2016, 3, 30, 23, 59, 59))
        '        Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " AfterQuery")
        '        Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " TableOutput (Y/N)?")
        '        If Console.ReadKey().KeyChar.ToString.ToLowerInvariant = "y" Then
        '            Console.WriteLine()
        '            Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " BeforeOutput")
        '            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(t))
        '        End If
        '        Console.WriteLine()
        '        Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " RowCount=" & t.Rows.Count)
        '        Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " AppEnd")
        '    End If
        'Catch ex As Exception
        '    Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " AppError")
        '    Console.WriteLine(ex.ToString)
        'End Try

        Try
            Console.WriteLine()
            Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " Execute TestSuite 'TestExchange2007' (Y/N)?")
            If Console.ReadKey().KeyChar.ToString.ToLowerInvariant = "y" Then
                TestExchange2007()
            End If
        Catch ex As Exception
            Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & " AppError")
            Console.WriteLine(ex.ToString)
        End Try
    End Sub



    Sub TestExchange2007()
        Try
            Dim oApp As New CompuMaster.Data.Outlook.OutlookApp(12)

            'Dim folderInbox As CompuMaster.Data.Outlook.FolderPathRepresentation = oApp.LookupFolder(WellKnownFolderName.Inbox)

            Dim folderRoot As CompuMaster.Data.Outlook.FolderPathRepresentation = oApp.LookupRootFolder("R:\Private\Jähnke\Outlook\Archive.pst")
            Dim dirRoot As Directory = folderRoot.Directory ' folderRoot.Directory.SelectSubFolder("AllItems", False, oApp.DirectorySeparatorChar)

            'ShowItems(dirRoot, e2007)

            Console.WriteLine(dirRoot.DisplayPath)
            ForEachSubDirectory(dirRoot.InitialRootDirectory, oApp)
            Console.WriteLine("Total count of subfolders for " & dirRoot.DisplayPath & ": " & folderRoot.Directory.SubFolderCount)
            Console.WriteLine("Total count of queried subfolders for " & dirRoot.DisplayPath & ": " & folderRoot.Directory.SubFolders.Length)

            'Dim folderRoot As CompuMaster.Data.MsExchange.FolderPathRepresentation = e2007.LookupFolder(WellKnownFolderName.MsgFolderRoot)
            'Dim dirRoot As Directory = folderRoot.Directory.SelectSubFolder("Inbox", False, e2007.DirectorySeparatorChar)

            'Dim dirInbox As Directory = dirRoot.InitialRootDirectory.SelectSubFolder("Oberste Ebene des Informationsspeichers\Inbox", False, e2007.DirectorySeparatorChar)
            Console.WriteLine()
            Dim dirInbox As Directory = dirRoot.SelectSubFolder("Posteingang\Heinrich-Haus", True, "\"c)
            Console.WriteLine("Inbox(manual lookup)=" & dirInbox.DisplayPath)
            ShowItems(dirInbox, oApp)
            'ShowItems(Convert2Items(dirRoot, e2007, New Microsoft.Exchange.WebServices.Data.Item() {dirInbox.ItemsAsExchangeItem()(0)}))
            'ShowItems(New Item() {dirInbox.Items()(0)}) 

            'ShowItems(New Item() {dirInbox.MailboxItems(SearchDefault, ItemViewDefault)(0)})

            'Console.WriteLine()
            'Console.WriteLine("Calendar appointments:")
            'ShowItems(New Item() {dirRoot.MailboxItems(SearchCalendar, ItemViewDefault)(0)})
            'ShowItems(New Item() {dirRoot.MailboxItems(SearchCalendar, ItemViewCalendarDefault)(0)})
            'ShowItems(dirRoot.MailboxItems(SearchCalendar, ItemViewDefault))
            'ShowItems(dirRoot.MailboxItems(SearchCalendar, ItemViewCalendarDefault))
            'ShowItems(dirRoot.MailboxItems(SearchInclCalendarEntries, ItemViewCalendarDefault))

            Console.WriteLine()
            'Dim foldersBelowRoot As Directory() = oApp.ListFolderItems(folderRoot)
            'Dim foldersBelowRoot As Directory() = e2007.ListSubFoldersRecursively(folderRoot)
            'Dim foldersBelowRoot As Directory() = dirRoot.SubFolders
            Dim testSubFolder As Directory = dirRoot
            Console.WriteLine("TEST SUBS FOR: " & testSubFolder.DisplayName)
            Console.WriteLine("TEST SUBS FOR: " & testSubFolder.FolderID)
            Console.WriteLine("TEST SUBS FOR: " & testSubFolder.SubFolderCount)
            'Console.WriteLine("TEST SUBS FOR: " & testSubFolder.SubFolderCoun)
            'foldersBelowRoot = e2007.ListSubFolders(New FolderPathRepresentation(testSubFolder.ExchangeFolder.))

            'Dim itemView As New Microsoft.Exchange.WebServices.Data.ItemView(Integer.MaxValue, 0, Microsoft.Exchange.WebServices.Data.OffsetBasePoint.Beginning)
            'Dim searchFilter As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsEqualTo(Microsoft.Exchange.WebServices.Data.ItemSchema.DateTimeCreated, New DateTime(2016 - 03 - 18))
            'Items = folderRoot.ExchangeFolder.FindItems(searchFilter, itemView)



            End

            'Dim u As Uri = e2007.SaveMailAsDraft("test", "test <b>plain</b>", "", Nothing, Nothing, Nothing)
            'e2007.SaveMailAsDraft("test", "", "text <b>html</b>", Nothing, Nothing, Nothing)
            'Console.WriteLine(u.ToString)
            End
            'e2007.ResolveMailboxOrContactNames("jochen")
            'e2007.CreateFolder("Test", e2007.LookupFolder(Microsoft.Exchange.WebServices.Data.WellKnownFolderName.Inbox, "CS\Sub\!Archiv", False))
            'e2007.CreateFolder("CS\Sub\!Archiv\Test\Sub-Test", e2007.LookupFolder(Microsoft.Exchange.WebServices.Data.WellKnownFolderName.Inbox, "", False))
            'e2007.EmptyFolder(e2007.LookupFolder(Microsoft.Exchange.WebServices.Data.WellKnownFolderName.Inbox, "CS\Sub\!Archiv\Test", False), DeleteMode.MoveToDeletedItems, False)
            'e2007.DeleteFolder(e2007.LookupFolder(Microsoft.Exchange.WebServices.Data.WellKnownFolderName.Inbox, "CS\Sub\!Archiv\Test", False), DeleteMode.MoveToDeletedItems)
            'Dim MyFolder As FolderPathRepresentation = e2007.LookupFolder(WellKnownFolderName.PublicFoldersRoot, "Company Contacts", False)
            Dim MyFolder As Directory = dirRoot.SelectSubFolder("Inbox", False, oApp.DirectorySeparatorChar)
            'Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTable(e2007.ListFolderItems(MyFolder)))
            Dim dt As DataTable
            'dt = Directory.ItemsAsDataTable(MyFolder.Items)
            'dt = CompuMaster.Data.DataTables.CreateDataTableClone(e2007.ListFolderItems(MyFolder), "subject like '*sürüm*' or subject like '*rund um berlin*'", "", 3)
            'dt = CompuMaster.Data.DataTables.CreateDataTableClone(e2007.ListFolderItems(MyFolder), "subject='Michael Pöfler' or subject = 'Elena Lamberti'", "", 3)
            'CompuMaster.Data.Csv.WriteDataTableToCsvFile("g:\cc.csv", dt)
            Dim ht As Hashtable = CompuMaster.Data.DataTables.FindDuplicates(dt.Columns("ID"))

            'dt.Rows.Add(dt.NewRow)
            'Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTable(dt))
            'Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTable(dt.Clone))
            Console.WriteLine(vbNewLine & "Data Rows: 2 first exemplary IDs:")
            Console.WriteLine(dt.Rows(0)("ID"))
            Console.WriteLine(dt.Rows(1)("ID"))
            Dim IDsAreEqual As Boolean = (dt.Rows(0)("ID").ToString = dt.Rows(1)("ID").ToString)
            If IDsAreEqual = False Then Console.WriteLine(Space(FirstDifferentChar(dt.Rows(0)("ID").ToString, dt.Rows(1)("ID").ToString)) & "^")
            Console.WriteLine("IDs are equal=" & IDsAreEqual.ToString.ToUpper)

            Console.WriteLine(vbNewLine & "DUPS:")
            For Each key As Object In ht.Keys
                Console.WriteLine(key.ToString & "=" & ht(key).ToString)
            Next
            'e2007.VerifyUniqueItemIDs(dt)

            'Console.WriteLine(vbnewline & "Re-Loading ID")
            'Dim c As Microsoft.Exchange.WebServices.Data.Contact = e2007.LoadContactData(Utils.NoDBNull(dt.Rows(0)("ID").ToString, ""))
            'Console.WriteLine(c.Subject)
            'c.Update(Microsoft.Exchange.WebServices.Data.ConflictResolutionMode.AutoResolve)

            End
            'e2007.SendMail("Test", "from CompuMaster.Data.Exchange2007SP1OrHigher" & vbNewLine & "on " & Now.ToString, New Recipient() {New Recipient("jwezel@compumaster.de")}, Nothing, Nothing)
            'e2007.CreateAppointment("Test-Appointment", "nowhere", "from CompuMaster.Data.Exchange2007SP1OrHigher" & vbNewLine & "on " & Now.ToString, Now.AddMinutes(5), New TimeSpan(0, 30, 0))
            'e2007.CreateMeetingAppointment("Test-Meeting", "nowhere", "from CompuMaster.Data.Exchange2007SP1OrHigher" & vbNewLine & "on " & Now.ToString, Now.AddMinutes(5), New TimeSpan(0, 30, 0), New Recipient() {New Recipient("jwezel@compumaster.de")}, Nothing, Nothing)
        Catch ex As Exception
            Console.WriteLine("Error: " + ex.ToString)
        End Try
    End Sub

    'Private Function SearchDefault() As Microsoft.Exchange.WebServices.Data.SearchFilter
    '    Dim searchFilterCollection As New Microsoft.Exchange.WebServices.Data.SearchFilter.SearchFilterCollection(Microsoft.Exchange.WebServices.Data.LogicalOperator.And)
    '    Dim searchFilterEarlierDate As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsGreaterThanOrEqualTo(Microsoft.Exchange.WebServices.Data.ItemSchema.DateTimeCreated, New DateTime(2016, 3, 10, 14, 0, 0))
    '    Dim searchFilterLaterDate As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsLessThanOrEqualTo(Microsoft.Exchange.WebServices.Data.ItemSchema.DateTimeCreated, New DateTime(2016, 3, 24, 14, 59, 59))
    '    searchFilterCollection.Add(searchFilterEarlierDate)
    '    searchFilterCollection.Add(searchFilterLaterDate)
    '    Return searchFilterCollection
    'End Function

    'Private Function SearchInclCalendarEntries() As Microsoft.Exchange.WebServices.Data.SearchFilter
    '    Dim searchFilterCollection As New Microsoft.Exchange.WebServices.Data.SearchFilter.SearchFilterCollection(Microsoft.Exchange.WebServices.Data.LogicalOperator.Or)
    '    searchFilterCollection.Add(SearchDefault)
    '    searchFilterCollection.Add(SearchCalendar)
    '    Return searchFilterCollection
    'End Function

    'Private Function SearchCalendar() As Microsoft.Exchange.WebServices.Data.SearchFilter
    '    Dim calEntriesSearchFilterCollection As New Microsoft.Exchange.WebServices.Data.SearchFilter.SearchFilterCollection(Microsoft.Exchange.WebServices.Data.LogicalOperator.And)
    '    Dim calItemClass As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsEqualTo(Microsoft.Exchange.WebServices.Data.ItemSchema.ItemClass, "IPM.Appointment")
    '    Dim calItemEventLatestStart As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsLessThanOrEqualTo(Microsoft.Exchange.WebServices.Data.AppointmentSchema.Start, New DateTime(2016, 3, 24, 14, 59, 59))
    '    Dim calItemEventEarliestEnd As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsGreaterThanOrEqualTo(Microsoft.Exchange.WebServices.Data.AppointmentSchema.End, New DateTime(2016, 3, 10, 14, 0, 0))
    '    calEntriesSearchFilterCollection.Add(calItemClass)
    '    calEntriesSearchFilterCollection.Add(calItemEventLatestStart)
    '    calEntriesSearchFilterCollection.Add(calItemEventEarliestEnd)
    '    Return calEntriesSearchFilterCollection
    'End Function

    'Private Function ItemViewDefault() As Microsoft.Exchange.WebServices.Data.ItemView
    '    Dim itemView As New Microsoft.Exchange.WebServices.Data.ItemView(Integer.MaxValue, 0, Microsoft.Exchange.WebServices.Data.OffsetBasePoint.Beginning)
    '    itemView.OrderBy.Add(Microsoft.Exchange.WebServices.Data.ItemSchema.DateTimeCreated, Microsoft.Exchange.WebServices.Data.SortDirection.Descending)
    '    'itemView.Traversal = Microsoft.Exchange.WebServices.Data.ItemTraversal.Associated
    '    Return itemView
    'End Function

    'Private Function ItemViewCalendarDefault() As Microsoft.Exchange.WebServices.Data.ItemView
    '    Dim itemView As New Microsoft.Exchange.WebServices.Data.ItemView(Integer.MaxValue, 0, Microsoft.Exchange.WebServices.Data.OffsetBasePoint.Beginning)
    '    itemView.OrderBy.Add(Microsoft.Exchange.WebServices.Data.AppointmentSchema.End, Microsoft.Exchange.WebServices.Data.SortDirection.Descending)
    '    itemView.OrderBy.Add(Microsoft.Exchange.WebServices.Data.AppointmentSchema.Start, Microsoft.Exchange.WebServices.Data.SortDirection.Descending)
    '    Return itemView
    'End Function

    Private Sub ShowItems(dir As Directory, e2007 As OutlookApp)
        Dim items As NetOffice.OutlookApi._Items = dir.OutlookFolder.Items
        ShowItems(Convert2Items(dir, e2007, items))
    End Sub

    Private Function Convert2Items(dir As Directory, e2007 As OutlookApp, items As NetOffice.OutlookApi._Items) As Item()
        Dim Result As New List(Of Item)
        For Each item As Object In items
            Result.Add(New Item(e2007, CType(item, NetOffice.COMObject), dir))
        Next
        'For MyItemCounter As Integer = 0 To System.Math.Min(1, items.Count) - 1
        '    Result.Add(New Item(e2007, CType(items(MyItemCounter), NetOffice.COMObject), dir))
        'Next
        Return Result.ToArray
    End Function
    'Private Function Convert2Items(dir As Directory, e2007 As OutlookApp, items As List(Of Microsoft.Exchange.WebServices.Data.Item)) As Item()
    '    '    Dim Result As New List(Of Item)
    '    '    For MyItemCounter As Integer = 0 To System.Math.Min(1, items.Count) - 1
    '    '        Result.Add(New Item(e2007, items(MyItemCounter), dir))
    '    '    Next
    '    '    Return Result.ToArray
    'End Function

    Private Sub ShowItems(items As Item())

        Console.WriteLine("    ---")
        For MyItemCounter As Integer = 0 To System.Math.Min(3, items.Length) - 1
            Dim entryItem As Item = items(MyItemCounter)
            Console.WriteLine("    " & entryItem.Subject) '& " / DC:" & entryItem.DateTimeCreated '& " / DR:" & entryItem.DateTimeReceived & " / DS:" & entryItem.DateTimeSent)
            'Console.WriteLine("    TYPE:" & entryItem.ExchangeItem.ItemClass)
            'Console.WriteLine("    CalBeg:" & entryItem.CalendarEntryBegin)
            'Console.WriteLine("    CalEnd:" & entryItem.CalendarEntryEnd)
            'Console.WriteLine("    Co:" & entryItem.MimeContent)
            'Console.WriteLine("    BT: " & entryItem.BodyType)
            'Console.WriteLine("    BC: " & entryItem.Body)
            'Console.WriteLine("    Fr: " & Utils.ObjectNotNothingOrEmptyString(entryItem.FromSender).ToString)

            'Console.WriteLine("    Fr: " & entryItem.FromExchangeSender)
            'Console.WriteLine("    To: " & entryItem.DisplayTo)
            'Console.WriteLine("    Cc: " & entryItem.DisplayCc)
            Console.WriteLine("    Pa: " & entryItem.ParentDirectory.DisplayPath)
            Console.WriteLine("    Cl: " & entryItem.ObjectClassName)
            Console.WriteLine("    ---")
        Next
    End Sub

    Private Sub ForEachSubDirectory(dir As Directory, e2007 As OutlookApp)

        For Each dirItem As Directory In dir.SubFolders
            Console.Write(dirItem.ToString)
            'Console.WriteLine(" (F:" & dirItem.SubFolderCount & " / U:" & dirItem.ItemUnreadCount & " / T:" & dirItem.ItemCount & ")")
            'Console.WriteLine(" (F:" & dirItem.SubFolderCount & " / T:" & dirItem.ItemCount & ")")
            Console.WriteLine()

            'Dim itemView As New Microsoft.Exchange.WebServices.Data.ItemView(Integer.MaxValue, 0, Microsoft.Exchange.WebServices.Data.OffsetBasePoint.Beginning)
            'Dim searchFilterCollection As New Microsoft.Exchange.WebServices.Data.SearchFilter.SearchFilterCollection(Microsoft.Exchange.WebServices.Data.LogicalOperator.And)
            'Dim searchFilterEarlierDate As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsGreaterThanOrEqualTo(Microsoft.Exchange.WebServices.Data.ItemSchema.DateTimeCreated, New DateTime(2016, 03, 18, 14, 00, 0))
            'Dim searchFilterLaterDate As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsLessThanOrEqualTo(Microsoft.Exchange.WebServices.Data.ItemSchema.DateTimeCreated, New DateTime(2016, 03, 18, 14, 59, 59))
            'searchFilterCollection.Add(searchFilterEarlierDate)
            'searchFilterCollection.Add(searchFilterLaterDate)

            ''Dim itemsEApi As Microsoft.Exchange.WebServices.Data.FindItemsResults(Of Microsoft.Exchange.WebServices.Data.Item) = dirItem.ExchangeFolder.FindItems(searchFilterCollection, itemView)
            'Dim items As ObjectModel.Collection(Of Microsoft.Exchange.WebServices.Data.Item) = dirItem.Items(searchFilterCollection, itemView)
            ''Dim items As ObjectModel.Collection(Of Microsoft.Exchange.WebServices.Data.Item) = dirItem.Items()
            ''If itemsEApi.Items.Count <> items.Count Or items.Count <> e2007.ListFolderItemsAsExchangeItems(dirItem).Length Then
            ''    Console.WriteLine("!!" & dirItem.ToString & " (" & e2007.ListFolderItemsAsExchangeItems(dirItem).Length & " of " & dirItem.ItemCount & ")")
            ''End If
            ''Console.WriteLine("    FType: " & dirItem.FolderClass)

            'Dim EndCounter As Integer
            'EndCounter += 1

            ''For Each editem As Generic.KeyValuePair(Of String, Object) In dirItem.ExtendedData
            ''    If editem.Value Is Nothing Then
            ''        Console.WriteLine("         " & editem.Key & "={NULL}")
            ''    Else
            ''        Console.WriteLine("         " & editem.Key & "=" & editem.Value.ToString)
            ''    End If
            ''Next
            ''If EndCounter >= 10 Then End

            'If False AndAlso True OrElse dirItem.DisplayPath.Contains("Inbox") Then

            '    For MyItemCounter As Integer = 0 To System.Math.Min(1, items.Count) - 1
            '        Dim entryItem As Item
            '        entryItem = New Item(e2007, items.Item(MyItemCounter), dirItem)
            '        Console.WriteLine("    " & entryItem.Subject & " / DC:" & entryItem.DateTimeCreated & " / DR:" & entryItem.DateTimeReceived & " / DS:" & entryItem.DateTimeSent)
            '        'Console.WriteLine("    Co:" & entryItem.MimeContent)
            '        'Console.WriteLine("    BT: " & entryItem.BodyType)
            '        'Console.WriteLine("    BC: " & entryItem.Body)
            '        'Console.WriteLine("    Fr: " & Utils.ObjectNotNothingOrEmptyString(entryItem.FromSender).ToString)

            '        Console.WriteLine("    Fr: " & entryItem.FromExchangeSender)
            '        Console.WriteLine("    To: " & entryItem.DisplayTo)
            '        Console.WriteLine("    Cc: " & entryItem.DisplayCc)
            '        Console.WriteLine("    Fr: " & entryItem.ParentDirectory.DisplayPath)
            '        'For Each addr As System.Net.Mail.MailAddress In entryItem.RecipientTo
            '        '    Console.WriteLine("    TO: " & addr.ToString)
            '        'Next
            '        'For Each addr As System.Net.Mail.MailAddress In entryItem.RecipientCc
            '        '    Console.WriteLine("    CC: " & addr.ToString)
            '        'Next
            '        'For Each addr As System.Net.Mail.MailAddress In entryItem.RecipientBcc
            '        '    Console.WriteLine("    BCC: " & addr.ToString)
            '        'Next
            '        'For Each addr As System.Net.Mail.MailAddress In entryItem.ReplyTo
            '        '    Console.WriteLine("    Repl: " & addr.ToString)
            '        'Next

            '        ''Console.WriteLine("T: " & entryItem.BodyText)
            '        ''Console.WriteLine("H: " & entryItem.BodyHtml)
            '        'For Each editem As Generic.KeyValuePair(Of String, Object) In entryItem.ExtendedData
            '        '    If editem.Value Is Nothing Then
            '        '        Console.WriteLine("         " & editem.Key & "={NULL}")
            '        '    Else
            '        '        Console.WriteLine("         " & editem.Key & "=" & editem.Value.ToString)
            '        '    End If
            '        'Next

            '    Next

            'End If

            ForEachSubDirectory(dirItem, e2007)

            'If dirItem.DisplayPath.Contains("Technik") Then
            '    ForEachSubDirectory(dirItem, e2007)
            'Else
            '    ForEachSubDirectory(dirItem, e2007)
            'End If
        Next

    End Sub

    Public Function FirstDifferentChar(ByVal value1 As String, ByVal value2 As String) As Integer
        Dim charCounter As Integer
        For charCounter = 0 To value1.Length
            If value1(charCounter) <> value2(charCounter) Then Return charCounter
        Next
        Return charCounter
    End Function
End Module

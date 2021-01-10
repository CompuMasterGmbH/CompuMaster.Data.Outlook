Option Explicit On
Option Strict On

Imports CompuMaster
Imports CompuMaster.Data.Outlook
Imports CompuMaster.Data.Outlook.OutlookApp

Module MainModule

    Sub Main()
        Dim OutlookApp As New CompuMaster.Data.Outlook.OutlookApp(12)
        Try
            Dim PstRootFolderPath As CompuMaster.Data.Outlook.FolderPathRepresentation
            PstRootFolderPath = OutlookApp.LookupRootFolder(System.IO.Path.Combine(My.Application.Info.DirectoryPath, "SampleData", "Mailbox.pst"))
            PstRootFolderPath.Directory.ForDirectoryAndEachSubDirectory(
            Sub(dir As CompuMaster.Data.Outlook.Directory)
                Console.Write(dir.DisplayPath) 'Console.Write(dir.ToString)
                'Console.Write(" [" & dir..FolderClass & "]")
                Console.Write(" (SubFolders:" & dir.SubFolderCount & " / UnReadItems:" & dir.ItemUnreadCount & " / TotalItems:" & dir.ItemCount & ")")
                'Console.Write(" (SubFolders:" & dir.SubFolderCount & " / TotalItems:" & dir.ItemCount & ")")
                Console.WriteLine()
                CompuMaster.Console.CurrentIndentationLevel += 1
                'ShowItems_FormatList(dir)
                ShowItems_FormatTable(dir)
                CompuMaster.Console.CurrentIndentationLevel -= 1
            End Sub)
            Console.WriteLine()
        Finally
            OutlookApp.Application.Quit()
        End Try

    End Sub

    Private Sub ShowItems_FormatList(dir As Directory)
        Dim items As NetOffice.OutlookApi._Items = dir.OutlookFolder.Items
        ShowItems_FormatList(Convert2Items(dir, items))
    End Sub

    Private Sub ShowItems_FormatTable(dir As Directory)
        Dim FolderItems As DataTable = dir.ItemsAllAsDataTable
        CompuMaster.Data.DataTables.RemoveColumns(FolderItems, New String() {"Body", "HTMLBody", "RTFBody"}) 'Do not show multi-line fields following steps
        CompuMaster.Data.DataTables.RemoveColumns(FolderItems, New String() {"ParentFolderID", "EntryID"}) 'Do not show ID fields in following steps
        Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(FolderItems))
    End Sub

    Private Function Convert2Items(dir As Directory, items As NetOffice.OutlookApi._Items) As Item()
        Dim Result As New List(Of Item)
        For Each item As Object In items
            Result.Add(New Item(dir.OutlookApp, CType(item, NetOffice.COMObject), dir))
        Next
        Return Result.ToArray
    End Function

    Private Sub ShowItems_FormatList(items As Item())
        Console.WriteLine("---")
        For MyItemCounter As Integer = 0 To System.Math.Min(3, items.Length) - 1
            Dim entryItem As Item = items(MyItemCounter)
            Console.WriteLine("" & entryItem.Subject) '& " / DC:" & entryItem.DateTimeCreated '& " / DR:" & entryItem.DateTimeReceived & " / DS:" & entryItem.DateTimeSent)
            'Console.WriteLine("TYPE:" & entryItem.ExchangeItem.ItemClass)
            If entryItem.Start <> Nothing Then Console.WriteLine("CalBeg:" & entryItem.Start)
            If entryItem.End <> Nothing Then Console.WriteLine("CalEnd:" & entryItem.End)
            If entryItem.StartUtc <> Nothing Then Console.WriteLine("CalBegUtc:" & entryItem.StartUtc)
            If entryItem.EndUtc <> Nothing Then Console.WriteLine("CalEndUtc:" & entryItem.EndUtc)
            'Console.WriteLine("Co:" & entryItem.MimeContent)
            Console.WriteLine("BT: " & entryItem.BodyFormat.ToString)
            Console.WriteLine("BC: " & entryItem.Body)
            'Console.WriteLine("Fr: " & Utils.ObjectNotNothingOrEmptyString(entryItem.FromSender).ToString)

            'Console.WriteLine("Fr: " & entryItem.FromExchangeSender)
            Console.WriteLine("To: " & entryItem.To)
            Console.WriteLine("Cc: " & entryItem.CC)
            Console.WriteLine("Pa: " & entryItem.ParentDirectory.DisplayPath)
            Console.WriteLine("Cl: " & entryItem.ObjectClassName)
            Console.WriteLine("---")
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

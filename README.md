# CompuMaster.Data.Outlook
Simplified access to already connected **Outlook Mailboxes** or to additional **Outlook PST** files with performant data transfer (reduced roundtrips to server)
based on [NetOfficeFw.OutlookApi](https://github.com/NetOfficeFw/NetOffice) (requires installed MS Outlook)

[![Github Release](https://img.shields.io/github/release/CompuMasterGmbH/CompuMaster.Data.Outlook.svg?maxAge=2592000&label=GitHub%20Release)](https://github.com/CompuMasterGmbH/CompuMaster.Data.Outlook/releases) 
[![NuGet CompuMaster.Data.Outlook](https://img.shields.io/nuget/v/CompuMaster.Data.Outlook.svg?label=NuGet%20CM.Data.Outlook)](https://www.nuget.org/packages/CompuMaster.Data.Outlook/) 

## Sample application

* Sample application for quick starting your own development  
![Screenshot of sample application](ExtractionSampleApp.png)
* Sample extraction of MS Outlook calendar entries into Microsoft Excel XLSX file (works for CSV files, too)
![Screenshot of sample application](ExtractionSampleExcelTable.png)

## Sample code

It's easy to browse and access all folders of Outlook's PST files, it takes just a few lines of code (also see sample console/debug app)  

### Sample code (C#)

```C#
class Program
{
    static void Main(string[] args)
    {
        CompuMaster.Data.Outlook.OutlookApp OutlookApp = new CompuMaster.Data.Outlook.OutlookApp(12);
        try
        {
            CompuMaster.Data.Outlook.FolderPathRepresentation PstRootFolderPath;
            PstRootFolderPath = OutlookApp.LookupRootFolder(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SampleData", "Mailbox.pst"));
            PstRootFolderPath.Directory.ForDirectoryAndEachSubDirectory((CompuMaster.Data.Outlook.Directory dir) =>
            {
                Console.Write(dir.Path);
                Console.Write(" (SubFolders:" + dir.SubFolderCount.ToString() + " / UnReadItems:" + dir.ItemUnreadCount().ToString() + " / TotalItems:" + dir.ItemCount().ToString() + ")");
                Console.WriteLine();
                ShowItems_FormatTable(dir);
            });
        }
        finally
        {
            OutlookApp.Application.Quit();
        }
    }

    // Following method requires CompuMaster.Data library (see NuGet gallery) to render DataTable nicely for command line output
    private static void ShowItems_FormatTable(CompuMaster.Data.Outlook.Directory dir)
    {
        System.Data.DataTable FolderItems = dir.ItemsAllAsDataTable();
        CompuMaster.Data.DataTables.RemoveColumns(FolderItems, new string[] { "Body", "HTMLBody", "RTFBody" }); // Do not show multi-line fields following steps
        CompuMaster.Data.DataTables.RemoveColumns(FolderItems, new string[] { "ParentFolderID", "EntryID" }); // Do not show ID fields in following steps
        Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(FolderItems));
    }
}
```

### Sample code (VB.NET)

```vb.net
Sub Main()
    Dim OutlookApp As New CompuMaster.Data.Outlook.OutlookApp(12)
    Try
        Dim PstRootFolderPath As CompuMaster.Data.Outlook.FolderPathRepresentation 
        PstRootFolderPath = OutlookApp.LookupRootFolder(System.IO.Path.Combine(My.Application.Info.DirectoryPath, "SampleData", "Mailbox.pst"))
        PstRootFolderPath.Directory.ForDirectoryAndEachSubDirectory(
            Sub(dir As CompuMaster.Data.Outlook.Directory)
                Console.Write(dir.Path)
                Console.Write(" (SubFolders:" & dir.SubFolderCount & " / UnReadItems:" & dir.ItemUnreadCount & " / TotalItems:" & dir.ItemCount & ")")
                Console.WriteLine()
                ShowItems_FormatTable(dir)
            End Sub)
    Finally
        OutlookApp.Application.Quit()
    End Try
End Sub

'Following method requires CompuMaster.Data library (see NuGet gallery) to render DataTable nicely for command line output
Private Sub ShowItems_FormatTable(dir As CompuMaster.Data.Outlook.Directory)
    Dim FolderItems As System.Data.DataTable = dir.ItemsAllAsDataTable
    CompuMaster.Data.DataTables.RemoveColumns(FolderItems, New String() {"Body", "HTMLBody", "RTFBody"}) 'Do not show multi-line fields following steps
    CompuMaster.Data.DataTables.RemoveColumns(FolderItems, New String() {"ParentFolderID", "EntryID"}) 'Do not show ID fields in following steps
    Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(FolderItems))
End Sub

```

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleConsoleAppCSharp
{
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
}
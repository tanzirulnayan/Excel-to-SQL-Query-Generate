using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
namespace ExcelToSqlQueryGenerate
{
    class Program
    {
        string path = "";
        _Application excel = new _Excel.Application();
        private Workbook wb;
        private Worksheet ws;
        List<Notification> notificationsList = new List<Notification>();

        public Program(string path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public void ReadCell(int numberOfRows)
        {
            for (int i = 2; i <= numberOfRows; i++)
            {
                try
                {
                    var notification = new Notification();
                    notification.UserID = (int)ws.Cells[i, 2].Value2;
                    notification.Url = ws.Cells[i, 3].Value2;
                    notification.Message = ws.Cells[i, 4].Value2.ToString();
                    notification.StatusID = (int)ws.Cells[i, 5].Value2;
                    notification.PostedBy = (int)ws.Cells[i, 7].Value2;
                    if (ws.Cells[i, 6].Value2 == "NULL" || ws.Cells[i, 6].Value2 == null)
                    {
                        notification.ReadDate = "NULL";
                    }
                    else
                    {
                        notification.ReadDate = "'" + DateTime.Parse(ws.Cells[i, 6].Value2).ToString() + "'";
                    }
                    notification.PostedDate = DateTime.FromOADate(double.Parse(ws.Cells[i, 8].Value2.ToString()));

                    this.notificationsList.Add(notification);
                }
                catch (Exception e)
                {
                    continue;
                }

                System.Console.WriteLine(i);
            }

            CreateSqlQuery(notificationsList);
        }

        private void CreateSqlQuery(List<Notification> notificationsList)
        {
            var sqlString = string.Empty;
            var query = string.Empty;
            if (notificationsList != null)
            {
                foreach (var notification in notificationsList)
                {
                    query = "\nINSERT INTO [dbo].[Notification] ([UserID],[Url],[Message],[StatusID],[ReadDate],[PostedBy],[PostedDate]) VALUES (" +
                            notification.UserID + ", '" + notification.Url + "', '" + notification.Message + "', " + notification.StatusID + ", " +
                            notification.ReadDate + ", " + notification.PostedBy + ", '" + notification.PostedDate + "') \nGO\n";
                    sqlString = sqlString + Environment.NewLine + query;
                }

                var fullSavePath = @"C:\Users\Tanzirul Haque Nayan\Documents\query.txt";
                using (FileStream fs = File.Create(fullSavePath))
                {
                    // Add some text to file    
                    Byte[] title1 = new UTF8Encoding(true).GetBytes(sqlString);
                    fs.Write(title1, 0, title1.Length);
                }
                System.Console.WriteLine("Operation Complete!");
            }

            System.Console.ReadKey();
        }

        public static void openFile()
        {
            Program excel = new Program(@"C:\Users\Tanzirul Haque Nayan\Downloads\data.xlsx", 1);
            excel.ReadCell(19753);
        }

        static void Main(string[] args)
        {
            openFile();
        }
    }
}

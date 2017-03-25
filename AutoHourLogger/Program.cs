using System;
using System.Configuration;
using System.Linq;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;

namespace AutoHourLogger
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var spreadsheetId = ConfigurationManager.AppSettings["SheetId"];
            var sheetName = ConfigurationManager.AppSettings["SheetName"];
            var range = ConfigurationManager.AppSettings["ReaderRange"];

            var searchValue = "Lecture 103";


            string[] _scopes = {SheetsService.Scope.Spreadsheets};
            var ApplicationName = "Google Sheet Auto Hour Logger";

            var credential = UserAuthentication.Authenticate(_scopes);

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });


            var dataReader = new SheetDataReader(spreadsheetId, sheetName, range);
            var values = dataReader.ReadData(service);
            var sheetCell = dataReader.FindCell(values, searchValue);


            var rc = Helper.GetRowAndColumnNumber(sheetCell);
            var newCell = Helper.GetCellNumber(int.Parse(rc.Split('|').ToArray()[0]) + 1,
                int.Parse(rc.Split('|').ToArray()[1]));


            var dataWriter = new SheetDataWriter(spreadsheetId, sheetName, "123:254", newCell);
            dataWriter.WriteToSheet(service);


            Console.WriteLine("done!");
        }
    }
}
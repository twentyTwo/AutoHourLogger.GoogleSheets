using System;
using System.Configuration;
using System.Linq;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;

namespace AutoHourLogger
{
    using System.Globalization;

    internal class Program
    {
        private static void Main(string[] args)
        {

            var searchValue = "Lecture 103";
            var whatTowrite = DateTime.Today.ToString(CultureInfo.InvariantCulture);

            var spreadsheetId = ConfigurationManager.AppSettings["SheetId"];
            var sheetName = ConfigurationManager.AppSettings["SheetName"];
            var range = ConfigurationManager.AppSettings["ReaderRange"];




            string[] _scopes = {SheetsService.Scope.Spreadsheets};
            var ApplicationName = "Google Sheet Auto Hour Logger";

            var credential = UserAuthentication.Authenticate(_scopes);

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });


            // Read data and find the cell with search text

            var dataReader = new SheetDataReader(spreadsheetId, sheetName, range);
            var values = dataReader.ReadData(service);
            var sheetCellNumber = dataReader.FindCell(values, searchValue);


            var rc = Helper.GetRowAndColumnNumber(sheetCellNumber);
            var cellToWrite = Helper.GetCellNumber(rc.RowNumber + 1, rc.ColumnNumber); 


            var dataWriter = new SheetDataWriter(spreadsheetId, sheetName, cellToWrite, whatTowrite);
            dataWriter.WriteToSheet(service);


            Console.WriteLine("done!");
        }
    }
}
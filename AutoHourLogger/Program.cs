using System;
using System.Configuration;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;

namespace AutoHourLogger
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                var spreadsheetId = ConfigurationManager.AppSettings["SheetId"];
                var sheetName = ConfigurationManager.AppSettings["SheetName"];
                var range = ConfigurationManager.AppSettings["ReaderRange"];
                var dateFormat = ConfigurationManager.AppSettings["DateFormat"];
                var timeToWrite = ConfigurationManager.AppSettings["Time"];


                var searchValue = DateTime.Now.ToString(dateFormat);

                TimeSpan time;
                if (!TimeSpan.TryParse(timeToWrite, out time))
                {
                    throw new Exception("Time in the settings can not converted to time duration.");
                }

                var whatTowrite = time;


                string[] scopes = {SheetsService.Scope.Spreadsheets};
                var ApplicationName = "Google Sheet Auto Hour Logger";

                var credential = UserAuthentication.Authenticate(scopes);

                // Create Google Sheets API service.
                var service = new SheetsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName
                });


                // Read data and find the cell with search text.

                var dataReader = new SheetDataReader(service, spreadsheetId, sheetName, range);
                var sheetDataTask = dataReader.ReadDataAsync();
                var sheetData = sheetDataTask.Result;


                if (sheetData != null)
                {
                    var sheetCellNumber = dataReader.FindCell(sheetData, searchValue);

                    if (sheetCellNumber != null)
                    {
                        var rc = Helper.GetRowAndColumnNumber(sheetCellNumber);
                        var cellToWrite = Helper.GetCellNumber(rc.RowNumber + 1, rc.ColumnNumber);

                        var dataWriter = new SheetDataWriter(service, spreadsheetId, sheetName);
                        var result = dataWriter.WriteToSheetAsync(cellToWrite, whatTowrite);
                        Console.WriteLine($"{result.Result} is updated");
                    }
                    else
                    {
                        Console.WriteLine("The target data not found in the sheet");
                    }
                }
                else
                {
                    Console.WriteLine("Sheet is empty");
                }

                Console.WriteLine("Data written successfully");
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Sheet reading is not successfull.\n{exception.Message}");
            }

            Console.ReadKey();
        }
    }
}
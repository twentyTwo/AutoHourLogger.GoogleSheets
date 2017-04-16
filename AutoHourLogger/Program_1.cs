//using System;
//using System.Configuration;
//using System.Globalization;
//using Google.Apis.Services;
//using Google.Apis.Sheets.v4;

//namespace AutoHourLogger
//{
//    internal class Program_1
//    {
//        private static void Main1(string[] args)
//        {
//            try
//            {

//                var searchValue = "Lecture 103";
//                var whatTowrite = DateTime.Today.ToString(CultureInfo.InvariantCulture);

//                var spreadsheetId = ConfigurationManager.AppSettings["SheetId"];
//                var sheetName = ConfigurationManager.AppSettings["SheetName"];
//                var range = ConfigurationManager.AppSettings["ReaderRange"];


//                string[] scopes = { SheetsService.Scope.Spreadsheets };
//                var ApplicationName = "Google Sheet Auto Hour Logger";

//                var credential = UserAuthentication.Authenticate(scopes);

//                // Create Google Sheets API service.
//                var service = new SheetsService(new BaseClientService.Initializer
//                {
//                    HttpClientInitializer = credential,
//                    ApplicationName = ApplicationName
//                });


//                // Read data and find the cell with search text.

//                var dataReader = new SheetDataReader(service, spreadsheetId, sheetName, range);
//                var sheetData = dataReader.ReadData();

//                if (sheetData != null)
//                {
//                    var sheetCellNumber = dataReader.FindCell(sheetData, searchValue);

//                    if (sheetCellNumber!=null)
//                    {
//                        var rc = Helper.GetRowAndColumnNumber(sheetCellNumber);
//                        var cellToWrite = Helper.GetCellNumber(rc.RowNumber + 1, rc.ColumnNumber);

//                        var dataWriter = new SheetDataWriter(service, spreadsheetId, sheetName);
//                        var result = dataWriter.WriteToSheet(cellToWrite, whatTowrite);
//                        Console.WriteLine($"{result.UpdatedData} is updated in {result.UpdatedCells}");

//                    }
//                    else
//                    {
//                        Console.WriteLine("The target data not found in the sheet");
//                    }
//                }
//                else
//                {
//                    Console.WriteLine("Sheet is empty");
//                }

//                Console.WriteLine("Data written successfully");              

//            }
//            catch (Exception exception)
//            {
//                Console.WriteLine($"Sheet reading is not successfull.\n{exception.Message}");
                
//            }

//        }
//    }
//}
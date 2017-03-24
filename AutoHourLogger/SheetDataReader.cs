using System;
using System.Collections.Generic;
using System.Configuration;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace AutoHourLogger
{
    public static class SheetDataReader
    {
        public static void ReadData(SheetsService service)
        {
            // Define request parameters.
            string spreadsheetId = ConfigurationManager.AppSettings["SheetId"];
            string range = ConfigurationManager.AppSettings["ReaderRange"];
            string sheetName = ConfigurationManager.AppSettings["SheetName"];

            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, sheetName + "!"+range);

            ValueRange response = request.Execute();
            IList<IList<object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    // Find the desired string and calculate the cell number/range
                    Console.WriteLine("{0}, {1}", row[0], row[4]);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
        }
    }
}
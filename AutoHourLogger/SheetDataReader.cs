using System;
using System.Collections.Generic;
using System.Configuration;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace AutoHourLogger
{
    public static class SheetDataReader
    {
        public static string TargetString { get; set; }
        public static void ReadData(SheetsService service)
        {
            TargetString = "Abcd";
            // Define request parameters.
            // Call them from main
            string spreadsheetId = ConfigurationManager.AppSettings["SheetId"];
            string range = ConfigurationManager.AppSettings["ReaderRange"];
            string sheetName = ConfigurationManager.AppSettings["SheetName"];

            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, sheetName + "!"+range);

            ValueRange response = request.Execute();
            IList<IList<object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                FindCell(values, TargetString);
            }
            else
            {
                Console.WriteLine("No data found.");
            }
        }

        private static string FindCell(IList<IList<object>> values, string targetString)
        {
            int rowNumber = 0;
            foreach (var row in values)
            {
                // Find the desired string and calculate the cell number/range
                for (int i = 0; i < row.Count; i++)
                {
                    
                    if ((string) row[i] == targetString)
                    {
                        return CalculateCell(rowNumber, i);
                    }
                }
                rowNumber++;
                Console.WriteLine("{0}, {1}", row[0], row[4]);
                
            }
            return null;
        }

        private static string CalculateCell(int rowNumber, int i)
        {
            return "A2:B";

        }
    }
}
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace AutoHourLogger
{
    public class SheetDataReader
    {
        private readonly string _spreadsheetId;
        private readonly string _sheetName;
        private readonly string _range;

        public SheetDataReader(string spreadsheetId, string sheetName, string range)
        {
            _range = range;
            _spreadsheetId = spreadsheetId;
            _sheetName = sheetName;
        }

        public IList<IList<object>> ReadData(SheetsService service)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(_spreadsheetId, _sheetName + "!"+_range);

            ValueRange response = request.Execute();
            IList<IList<object>> values = response.Values;

            return values;
        }
        

        public string FindCell(IList<IList<object>> values, string targetString)
        {

            if(values ==null || values.Count == 0)
                throw new Exception("No data found in sheet.");

            int rowNumber = 0;
            foreach (var row in values)
            {
                for (int columnNumber = 0; columnNumber < row.Count; columnNumber++)
                {

                    if ((string)row[columnNumber] == targetString)
                    {
                        return CalculateCell(rowNumber, columnNumber);
                    }
                }
                rowNumber++;
            }

            return null;

        }

        private string CalculateCell(int rowNumber, int columnNumber)
        {
            var rowAndColumn = Helper.GetRowAndColumnNumber(_range.Split(':').ToArray()[0]);

            int row  = int.Parse(rowAndColumn.Split('|').ToArray()[0]);
            int column = int.Parse(rowAndColumn.Split('|').ToArray()[1]);

            return Helper.GetCellNumber(rowNumber+row, columnNumber+column);
        }
        
    }
}

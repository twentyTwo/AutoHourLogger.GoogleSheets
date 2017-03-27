using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Sheets.v4;

namespace AutoHourLogger
{
    public class SheetDataReader
    {
        private readonly string _range;
        private readonly SheetsService _service;
        private readonly string _sheetName;
        private readonly string _spreadsheetId;

        public SheetDataReader(SheetsService service, string spreadsheetId, string sheetName, string range)
        {
            _range = range;
            _service = service;
            _spreadsheetId = spreadsheetId;
            _sheetName = sheetName;
        }

        public IList<IList<object>> ReadData()
        {
            var request = _service.Spreadsheets.Values.Get(_spreadsheetId, _sheetName + "!" + _range);

            var response = request.Execute();
            var values = response.Values;

            return values;
        }

        public string FindCell(IList<IList<object>> values, string targetString)
        {
            if (values == null || values.Count == 0)
                throw new Exception("No data found in sheet.");

            var rowNumber = 0;
            foreach (var row in values)
            {
                for (var columnNumber = 0; columnNumber < row.Count; columnNumber++)
                {
                    if ((string) row[columnNumber] == targetString)
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

            var row = rowAndColumn.RowNumber;
            var column = rowAndColumn.ColumnNumber;

            return Helper.GetCellNumber(rowNumber + row, columnNumber + column);
        }
    }
}
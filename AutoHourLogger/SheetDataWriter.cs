using System;
using System.Collections.Generic;
using System.Configuration;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace AutoHourLogger
{
    public class SheetDataWriter
    {
        private readonly string _sheetId;
        private readonly string _sheetName;
        private readonly string _valueToWrite;
        private readonly string _sheetCellNumber;


        public SheetDataWriter(string sheetId, string sheetName, string valueToWrite, string sheetCellNumber)
        {
            _sheetId = sheetId;
            _sheetName = sheetName;
            _valueToWrite = valueToWrite;
            _sheetCellNumber = sheetCellNumber;
        }

        public void WriteToSheet(SheetsService service)
        {
            String range = _sheetName + "!" + _sheetCellNumber; // "Basic!B111";
            ValueRange valueRange = new ValueRange { MajorDimension = "COLUMNS" };


            var oblist = new List<object>() { _valueToWrite };
            valueRange.Values = new List<IList<object>> { oblist };

            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, _sheetId, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result = update.Execute();
            Console.WriteLine("{0} is updated to {1}",result.UpdatedRange, result.UpdatedData);
        }
    }
}
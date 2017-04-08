using System;
using System.Collections.Generic;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace AutoHourLogger
{
    public class SheetDataWriter
    {
        private readonly SheetsService _service;
        private readonly string _sheetId;
        private readonly string _sheetName;

        public SheetDataWriter(SheetsService service, string sheetId, string sheetName)
        {
            _service = service;
            _sheetId = sheetId;
            _sheetName = sheetName;
        }

        public UpdateValuesResponse WriteToSheet(string sheetCellNumber, string valueToWrite)
        {
            try
            {
                var range = _sheetName + "!" + sheetCellNumber; // "Basic!B111";
                var valueRange = new ValueRange { MajorDimension = "COLUMNS" };


                var objectList = new List<object> { valueToWrite };
                valueRange.Values = new List<IList<object>> { objectList };

                var update = _service.Spreadsheets.Values.Update(valueRange, _sheetId, range);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                return update.Execute();
                
            }
            catch (Exception )
            {
                
                throw new Exception("Error in writting data");
            }

            
        }
    }
}
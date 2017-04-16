// // --------------------------------------------------------------------------------------------------------------------
// // <copyright file="noor.alam.shuvo@gmail.com" company="">
// //   Copyright @ 2017
// // </copyright>
// <summary>
// // </summary>
// // --------------------------------------------------------------------------------------------------------------------

namespace AutoHourLogger
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;

    using Google.Apis.Sheets.v4;
    using Google.Apis.Sheets.v4.Data;

    public class SheetDataWriter
    {
        private readonly SheetsService _service;

        private readonly string _sheetId;

        private readonly string _sheetName;

        public SheetDataWriter(SheetsService service, string sheetId, string sheetName)
        {
            this._service = service;
            this._sheetId = sheetId;
            this._sheetName = sheetName;
        }

        public async Task<UpdateValuesResponse> WriteToSheetAsync(string sheetCellNumber, TimeSpan valueToWrite)
        {
            try
            {
                var range = this._sheetName + "!" + sheetCellNumber; // "Basic!B111";
                var valueRange = new ValueRange { MajorDimension = "COLUMNS" };

                var objectList = new List<object> { valueToWrite };
                valueRange.Values = new List<IList<object>> { objectList };

                var update = this._service.Spreadsheets.Values.Update(valueRange, this._sheetId, range);
                update.ValueInputOption =
                    SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

                return await update.ExecuteAsync();
            }
            catch (Exception)
            {
                throw new Exception("Error in writing data");
            }
        }
    }
}
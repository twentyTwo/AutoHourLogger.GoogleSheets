using System;
using System.Collections.Generic;
using System.Configuration;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace AutoHourLogger
{
    public static class SheetDataWriter
    {
        public static void WriteToSheet(SheetsService service)
        {
            string spreadsheetId = ConfigurationManager.AppSettings["SheetId"];
            string inputValue = ConfigurationManager.AppSettings["InputValue"]; 

            String range = GetTabAndCell();
            ValueRange valueRange = new ValueRange { MajorDimension = "COLUMNS" };


            var oblist = new List<object>() { inputValue };
            valueRange.Values = new List<IList<object>> { oblist };

            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();
        }
        private static string GetTabAndCell()
        {
            return "Basic!B111";
        }
    }
}
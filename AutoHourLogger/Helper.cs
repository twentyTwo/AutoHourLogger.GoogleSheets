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
    using System.Text;

    public static class Helper
    {
        public static string GetCellNumber(int row, int column)
        {
            var columnSheet = new StringBuilder();

            while (column > 0)
            {
                var cm = column % 26;

                if (cm == 0)
                {
                    column--;
                    columnSheet.Insert(0, 'Z');
                }
                else
                {
                    columnSheet.Insert(0, (char)(cm + 'A' - 1));
                }

                column /= 26;
            }

            return $"{columnSheet}{row}";
        }

        public static SheetCellNumeric GetRowAndColumnNumber(string sheetCellNumber)
        {
            var i = 0;

            for (; i < sheetCellNumber.Length; i++)
            {
                if (char.IsDigit(sheetCellNumber[i]))
                {
                    break;
                }
            }

            var column = 0;
            var row = sheetCellNumber.Substring(i);

            var cs = sheetCellNumber.Substring(0, i);
            var csl = cs.Length - 1;

            for (var j = 0; j < cs.Length; j++)
            {
                var cc = cs[j] - 'A' + 1;

                var multi = (int)Math.Pow(26, csl--);

                column += cc * multi;
            }

            return new SheetCellNumeric { RowNumber = int.Parse(row), ColumnNumber = column };
        }

        public class SheetCellNumeric
        {
            public int ColumnNumber;

            public int RowNumber;
        }
    }
}
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFilter
{
    class ExcelTable : DataTable
    {
        public static ExcelTable Load(string FileName)
        {
            XLWorkbook FilterTable = new XLWorkbook(FileName);
            IXLWorksheet FilterSheet = FilterTable.Worksheets.ToList()[0];

            ExcelTable Filters = new ExcelTable();

            var rows = FilterSheet.RangeUsed().RowsUsed().Skip(0); // Skip header row
            foreach (var row in rows)
            {
                var rowNumber = row.RowNumber();
                // Process the row
            }

            return Filters
        }
    }
}

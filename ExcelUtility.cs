using System.Collections.Generic;
using ClosedXML.Excel;

namespace ExcelLibrary
{
    public static class ExcelUtility
    {
        public static IXLWorkbook  WorksheetFromIEnumerable<T>(IEnumerable<T> source)
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add(typeof(T).ToString());
            var properties = typeof(T).GetProperties();
            var row = 1;
            var col = 1;
            foreach (var prop in properties)
            {
                worksheet.Cell(row, col++).Value = prop.Name;
            }

            foreach (var record in source)
            {
                row++;
                col = 1;
                foreach (var prop in properties)
                {
                    var cell = worksheet.Cell(row, col++);
                   cell.Value = prop.GetValue(record);
                    switch (prop.PropertyType.Name)
                    {
                        case "int":
                        case "double":
                        case "float":
                        case "decimal":
                            cell.SetDataType(XLDataType.Number);
                            break;
                        case "bool":
                            cell.SetDataType(XLDataType.Boolean);
                            break;
                    }
                }
            }

            return workbook;
        }
    }
}

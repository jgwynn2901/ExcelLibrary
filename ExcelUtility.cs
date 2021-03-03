using System;
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
                    switch (prop.PropertyType.Name.ToLower())
                    {
                        case "int":
                        case "double":
                        case "float":
                        case "decimal":
                            cell.SetDataType(XLDataType.Number);
                            break;
                        case "boolean":
                            cell.SetDataType(XLDataType.Boolean);
                            break;
                    }
                }
            }

            return workbook;
        }

        public static IEnumerable<T>  IEnumerableFromWorksheet<T> (IXLWorksheet sheet) where T : new()
        {
            var results = new List<T>();
            var properties = typeof(T).GetProperties();
            var lastRow = sheet.LastRowUsed().RowNumber();
            for (var row = 2; row <= lastRow; ++row)
            {
                var currentRow = sheet.Row(row);
                var data = new T();
                var col = 1;
                foreach (var prop in properties)
                {
                    var cell = currentRow.Cell(col++);
                    switch (prop.PropertyType.Name.ToLower())
                    {
                        case "int":
                            prop.SetValue(data, (int) cell.Value);
                            break;
                        case "decimal":
                            prop.SetValue(data, Convert.ToDecimal(cell.Value));
                            break;
                        case "boolean":
                            prop.SetValue(data, Convert.ToBoolean(cell.Value));
                            break;
                        default:
                            prop.SetValue(data, cell.Value);
                            break;
                    }
                }

                results.Add(data);
            }

            return results;
        }
    }
}

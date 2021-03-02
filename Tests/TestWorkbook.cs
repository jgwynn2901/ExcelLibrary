using ClosedXML.Excel;
using Xunit;

namespace ExcelLibrary.Tests
{
    public class TestWorkbook
    {
        [Fact]
        public void ExampleTest()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sample Sheet");
                worksheet.Cell("A1").Value = "Hello World!";
                worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
                workbook.SaveAs("HelloWorld.xlsx");
            }
        }
    }
}

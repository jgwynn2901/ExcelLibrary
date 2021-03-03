using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ExcelLibrary;
using Xunit;

namespace ExcelTests
{
    public class User
    {
        public string Name { get; set; }
        public string Email { get; set; }   
        public bool Registered { get; set; }
        public decimal Cost { get; set; }

    }

    public class TestWorkbook
    {
        [Fact]
        public void ExampleTest()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sample Sheet");
                worksheet.Cell(1,1).Value = "Hello World!";
                worksheet.Cell(1,2).FormulaA1 = "=MID(A1, 7, 5)";
                worksheet.Cell(2,1).Value = "Hello BWorld!";
                worksheet.Cell(2,2).FormulaA1 = "=MID(A2, 7, 5)";
                workbook.SaveAs("HelloWorld.xlsx");
            }
        }

        [Fact]
        public void TestFromIEnumerable()
        {
            var userList = new List<User>
            {
                new User {Name = "Fred Flintstone", Email = "FreddieTheFlint@sedgwick.com", Registered = false, Cost = 122.16M},
                new User {Name = "Barney Rubble", Email = "BarnieTheRub@sedgwick.com",Registered = true, Cost = 1200.20M},
                new User {Name = "Pebbles Flintstone", Email = "PebbyTheFlint@sedgwick.com",Registered = false, Cost = 990.99M},
                new User {Name = "Wilma Flintstone", Email = "WilmieTheFlint@sedgwick.com",Registered = false, Cost = 101.12M}
            };
            using (var book = ExcelUtility.WorksheetFromIEnumerable(userList))
            {
                book.SaveAs("Flintstones.xlsx");
            }
        }

        [Fact]
        public void TestToIEnumerable()
        {
            using (var book = new XLWorkbook("Flintstones.xlsx"))
            {
                var sheet = book.Worksheet(1);
                Assert.NotNull(sheet);
                var results = ExcelUtility.IEnumerableFromWorksheet<User>(sheet);
                Assert.True(results.Count() == 4);
            }
        }
    }
}

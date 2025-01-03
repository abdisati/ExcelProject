using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelLibray;
using ExcelDataReaderApp;

namespace ExcelDataReaderTest
{
    [TestClass]
    public class MainMethodTests
    {
        private TextWriter? originalOut;
        private StringWriter? stringWriter;
        private TextReader? originalIn;
        private string? tempFilePath1;
        private string? tempFilePath2;

        [TestInitialize]
        public void TestInitialize()
        {
            originalOut = Console.Out;
            stringWriter = new StringWriter();
            Console.SetOut(stringWriter);

            originalIn = Console.In;

            // Create temporary Excel files for testing
            tempFilePath1 = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
            tempFilePath2 = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");

            CreateTestExcelFile(tempFilePath1, new List<Person>
            {
                new Person { Name = "John Doe", Age = 30, City = "New York" },
                new Person { Name = "Jane Smith", Age = 25, City = "Los Angeles" }
            });

            CreateTestExcelFile(tempFilePath2, new List<Person>
            {
                new Person { Name = "John Doe", Age = 30, City = "New York" },
                new Person { Name = "Jane Smith", Age = 26, City = "Los Angeles" }
            });
        }

        [TestCleanup]
        public void TestCleanup()
        {
            if (originalOut != null)
            {
                Console.SetOut(originalOut);
            }
            stringWriter?.Dispose();
            stringWriter = null; // Ensure no residual state

            if (originalIn != null)
            {
                Console.SetIn(originalIn);
            }

            // Delete the temporary files after testing
            if (File.Exists(tempFilePath1))
            {
                File.Delete(tempFilePath1);
            }
            if (File.Exists(tempFilePath2))
            {
                File.Delete(tempFilePath2);
            }
        }

        private void CreateTestExcelFile(string filePath, List<Person> people)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");
                worksheet.Cell(1, 1).Value = "Name";
                worksheet.Cell(1, 2).Value = "Age";
                worksheet.Cell(1, 3).Value = "City";

                for (int i = 0; i < people.Count; i++)
                {
                    worksheet.Cell(i + 2, 1).Value = people[i].Name;
                    worksheet.Cell(i + 2, 2).Value = people[i].Age;
                    worksheet.Cell(i + 2, 3).Value = people[i].City;
                }

                workbook.SaveAs(filePath);
            }
        }

        [TestMethod]
        public void TestMainMethod_ValidInput()
        {
            // Arrange
            var input = $"{tempFilePath1}\nSheet1\n{tempFilePath2}\nSheet1\n";
            using (var stringReader = new StringReader(input))
            {
                Console.SetIn(stringReader);

                // Act
                ExcelReaderProgram.Main();

                // Assert
                var output = stringWriter?.ToString();
                Assert.IsNotNull(output);
                Assert.IsTrue(output.Contains("Exact match: Name: John Doe, Age: 30, City: New York"));
                Assert.IsTrue(output.Contains("Partial match for Jane Smith:"));
                Assert.IsTrue(output.Contains("Age differs: Sheet1=25, Sheet2=26"));
            }
        }

        [TestMethod]
        public void TestMainMethod_InvalidFilePath()
        {
            // Arrange
            var input = $"invalid_path.xlsx\nSheet1\n{tempFilePath2}\nSheet1\n";
            using (var stringReader = new StringReader(input))
            {
                Console.SetIn(stringReader);

                // Act
                ExcelReaderProgram.Main();

                // Assert
                var output = stringWriter?.ToString();
                Assert.IsNotNull(output);
                Assert.IsTrue(output.Contains("Invalid file path(s)."));
            }
        }
    }
}
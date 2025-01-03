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
    public class LoadPeopleFromExcelTests
    {
        private string? tempFilePath;

        [TestInitialize]
        public void TestInitialize()
        {
            // Create a temporary Excel file for testing
            tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");
                worksheet.Cell(1, 1).Value = "Name";
                worksheet.Cell(1, 2).Value = "Age";
                worksheet.Cell(1, 3).Value = "City";
                worksheet.Cell(2, 1).Value = "John Doe";
                worksheet.Cell(2, 2).Value = 30;
                worksheet.Cell(2, 3).Value = "New York";
                worksheet.Cell(3, 1).Value = "Jane Smith";
                worksheet.Cell(3, 2).Value = 25;
                worksheet.Cell(3, 3).Value = "Los Angeles";
                workbook.SaveAs(tempFilePath);
            }
        }

        [TestCleanup]
        public void TestCleanup()
        {
            // Delete the temporary file after testing
            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }
        }

        [TestMethod]
        public void TestLoadPeopleFromExcel_ValidData()
        {
            // Act
            var people = ExcelReaderProgram.LoadPeopleFromExcel(tempFilePath!, "Sheet1");

            // Assert
            Assert.AreEqual(2, people.Count);
            Assert.IsTrue(people.ContainsKey("john doe"));
            Assert.IsTrue(people.ContainsKey("jane smith"));
            Assert.AreEqual(1, people["john doe"].Count);
            Assert.AreEqual(1, people["jane smith"].Count);
            Assert.AreEqual("John Doe", people["john doe"][0].Name);
            Assert.AreEqual(30, people["john doe"][0].Age);
            Assert.AreEqual("New York", people["john doe"][0].City);
            Assert.AreEqual("Jane Smith", people["jane smith"][0].Name);
            Assert.AreEqual(25, people["jane smith"][0].Age);
            Assert.AreEqual("Los Angeles", people["jane smith"][0].City);
        }

        [TestMethod]
        public void TestLoadPeopleFromExcel_InvalidSheetName()
        {
            // Act
            var people = ExcelReaderProgram.LoadPeopleFromExcel(tempFilePath!, "InvalidSheet");

            // Assert
            Assert.AreEqual(0, people.Count);
        }

        [TestMethod]
        public void TestLoadPeopleFromExcel_InvalidFilePath()
        {
            // Act
            var people = ExcelReaderProgram.LoadPeopleFromExcel("invalid_path.xlsx", "Sheet1");

            // Assert
            Assert.AreEqual(0, people.Count);
        }
    }
}
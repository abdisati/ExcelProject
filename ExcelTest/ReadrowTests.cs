using System;
using System.IO;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using ExcelLibray;
using ExcelDataReaderApp;

namespace ExcelDataReaderTest
{
    [TestClass]
    public class ReadRowTests
    {
        private TextWriter? originalOut;
        private StringWriter? stringWriter;

        [TestInitialize]
        public void TestInitialize()
        {
            originalOut = Console.Out;
            stringWriter = new StringWriter();
            Console.SetOut(stringWriter);
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
        }

        private string CaptureConsoleOutput(Action action)
        {
            using (var sw = new StringWriter())
            {
                var originalOut = Console.Out;
                try
                {
                    Console.SetOut(sw);
                    action();
                    return sw.ToString().Trim();
                }
                finally
                {
                    Console.SetOut(originalOut);
                }
            }
        }

        [TestMethod]
        public void TestReadRow_ValidData_ReturnsCorrectPerson()
        {
            // Arrange
            var mockRow = CreateMockRow("John Doe", 30, "New York");

            // Act
            var person = ExcelReaderProgram.ReadRow(mockRow);

            // Assert
            Assert.AreEqual("John Doe", person.Name);
            Assert.AreEqual(30, person.Age);
            Assert.AreEqual("New York", person.City);
        }
        [TestMethod]
        public void TestReadRow_ErrorReadingName_ReturnsDefaultForName()
        {
            // Arrange
            var mockRow = CreateMockRow(1, 30, "New York");

            // Act
            var output = CaptureConsoleOutput(() =>
            {
                var person = ExcelReaderProgram.ReadRow(mockRow);

                // Assert
                Assert.AreEqual("Unknown", person.Name); // Default value for invalid name
                Assert.AreEqual(30, person.Age);
                Assert.AreEqual("New York", person.City);
            });

            Assert.IsTrue(output.Contains("Error reading 'Name'"), "Expected error message for invalid name not found.");
        }

        [TestMethod]
        public void TestReadRow_InvalidAge_ReturnsDefaultForInvalidAge()
        {
            // Arrange
            var mockRow = CreateMockRow("John Doe", "invalid_age", "New York");

            // Act
            var output = CaptureConsoleOutput(() =>
            {
                var person = ExcelReaderProgram.ReadRow(mockRow);

                // Assert
                Assert.AreEqual("John Doe", person.Name);
                Assert.AreEqual(-1, person.Age); // Default value for invalid age
                Assert.AreEqual("New York", person.City);
            });

            Assert.IsTrue(output.Contains("Invalid data type for 'Age'"), "Expected error message for invalid age not found.");
        }



        [TestMethod]
        public void TestReadRow_EmptyCity_ReturnsDefaultForEmptyCity()
        {
            // Arrange
            var mockRow = CreateMockRow("John Doe", 30, "");

            // Act
            var person = ExcelReaderProgram.ReadRow(mockRow);

            // Assert
            Assert.AreEqual("John Doe", person.Name);
            Assert.AreEqual(30, person.Age);
            Assert.AreEqual("Unknown", person.City); // Default value for empty city
        }

        [TestMethod]
        public void TestReadRow_ErrorReadingCity_ReturnsDefaultForCity()
        {
            // Arrange
            var mockRow = CreateMockRow("John Doe", 30, 33);

            // Act
            var output = CaptureConsoleOutput(() =>
            {
                var person = ExcelReaderProgram.ReadRow(mockRow);

                // Assert
                Assert.AreEqual("John Doe", person.Name);
                Assert.AreEqual(30, person.Age);
                Assert.AreEqual("Unknown", person.City); // Default value for invalid city
            });

            Assert.IsTrue(output.Contains("Error reading 'City'"), "Expected error message for invalid city not found.");
        }

        private static IXLRow CreateMockRow(object name, object age, object city)
        {
            var mockRow = new Mock<IXLRow>();
            var mockNameCell = new Mock<IXLCell>();
            var mockAgeCell = new Mock<IXLCell>();
            var mockCityCell = new Mock<IXLCell>();

            mockNameCell.Setup(cell => cell.GetValue<object>()).Returns(name);
            mockAgeCell.Setup(cell => cell.GetValue<object>()).Returns(age);
            mockCityCell.Setup(cell => cell.GetValue<object>()).Returns(city);

            mockRow.Setup(row => row.Cell(1)).Returns(mockNameCell.Object);
            mockRow.Setup(row => row.Cell(2)).Returns(mockAgeCell.Object);
            mockRow.Setup(row => row.Cell(3)).Returns(mockCityCell.Object);

            return mockRow.Object;
        }
    }
}
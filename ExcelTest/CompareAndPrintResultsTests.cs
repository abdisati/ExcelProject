using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelLibray;
using ExcelDataReaderApp;

namespace ExcelDataReaderTest
{
    [TestClass]
    public class CompareAndPrintResultsTests
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
        public void TestCompareAndPrintResults_ExactMatch()
        {
            // Arrange
            var peopleSheet1 = new Dictionary<string, List<Person>>
            {
                { "john doe", new List<Person> { new Person { Name = "John Doe", Age = 30, City = "New York" } } }
            };
            var peopleSheet2 = new Dictionary<string, List<Person>>
            {
                { "john doe", new List<Person> { new Person { Name = "John Doe", Age = 30, City = "New York" } } }
            };

            // Act
            var output = CaptureConsoleOutput(() => ExcelReaderProgram.CompareAndPrintResults(peopleSheet1, peopleSheet2));

            // Assert
            Assert.IsTrue(output.Contains("Exact match: Name: John Doe, Age: 30, City: New York"));
        }

        [TestMethod]
        public void TestCompareAndPrintResults_PartialMatch()
        {
            // Arrange
            var peopleSheet1 = new Dictionary<string, List<Person>>
            {
                { "john doe", new List<Person> { new Person { Name = "John Doe", Age = 30, City = "New York" } } }
            };
            var peopleSheet2 = new Dictionary<string, List<Person>>
            {
                { "john doe", new List<Person> { new Person { Name = "John Doe", Age = 31, City = "Los Angeles" } } }
            };

            // Act
            var output = CaptureConsoleOutput(() => ExcelReaderProgram.CompareAndPrintResults(peopleSheet1, peopleSheet2));

            // Assert
            Assert.IsTrue(output.Contains("Partial match for John Doe:"));
            Assert.IsTrue(output.Contains("Age differs: Sheet1=30, Sheet2=31"));
            Assert.IsTrue(output.Contains("City differs: Sheet1=New York, Sheet2=Los Angeles"));
        }

        [TestMethod]
        public void TestCompareAndPrintResults_OnlyInSheet1()
        {
            // Arrange
            var peopleSheet1 = new Dictionary<string, List<Person>>
            {
                { "john doe", new List<Person> { new Person { Name = "John Doe", Age = 30, City = "New York" } } }
            };
            var peopleSheet2 = new Dictionary<string, List<Person>>();

            // Act
            var output = CaptureConsoleOutput(() => ExcelReaderProgram.CompareAndPrintResults(peopleSheet1, peopleSheet2));

            // Assert
            Assert.IsTrue(output.Contains("Only in Sheet1: Name: John Doe, Age: 30, City: New York"));
        }

        [TestMethod]
        public void TestCompareAndPrintResults_OnlyInSheet2()
        {
            // Arrange
            var peopleSheet1 = new Dictionary<string, List<Person>>();
            var peopleSheet2 = new Dictionary<string, List<Person>>
            {
                { "john doe", new List<Person> { new Person { Name = "John Doe", Age = 30, City = "New York" } } }
            };

            // Act
            var output = CaptureConsoleOutput(() => ExcelReaderProgram.CompareAndPrintResults(peopleSheet1, peopleSheet2));

            // Assert
            Assert.IsTrue(output.Contains("Only in Sheet2: Name: John Doe, Age: 30, City: New York"));
        }
    }
}
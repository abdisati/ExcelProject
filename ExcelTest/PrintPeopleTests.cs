using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelLibray;
using ExcelDataReaderApp;

namespace ExcelDataReaderTest
{
    [TestClass]
    public class PrintPeopleTests
    {
        private TextWriter? originalOut;
        private StringWriter? stringWriter;

        [TestInitialize]
        public void TestInitialize()
        {
            originalOut = Console.Out;
            stringWriter = new StringWriter(); // Create a new StringWriter for each test
            Console.SetOut(stringWriter); // Redirect Console output to StringWriter
        }

        [TestCleanup]
        public void TestCleanup()
        {
            // Reset the console output to its original state
            if (originalOut != null)
            {
                Console.SetOut(originalOut);
            }
            stringWriter?.Dispose();
            stringWriter = null; // Ensure no residual state
        }

        private string CaptureConsoleOutput(Action action)
        {
            // Since the console output is already redirected in TestInitialize, 
            // this method simply executes the action without needing nested redirection.
            action();
            return stringWriter!.ToString().Trim(); // Capture output from the existing stringWriter
        }

        [TestMethod]
        public void TestPrintPeople_EmptyDictionary()
        {
            // Arrange
            var people = new Dictionary<string, Person>();

            // Act
            var output = CaptureConsoleOutput(() => ExcelReaderProgram.PrintPeople(people));

            // Assert
            Assert.AreEqual(string.Empty, output);
        }

        [TestMethod]
        public void TestPrintPeople()
        {
            // Arrange
            var people = new Dictionary<string, Person>
            {
                { "john doe-30-new york", new Person { Name = "John Doe", Age = 30, City = "New York" } },
                { "jane smith-25-los angeles", new Person { Name = "Jane Smith", Age = 25, City = "Los Angeles" } }
            };

            // Act
            var output = CaptureConsoleOutput(() => ExcelReaderProgram.PrintPeople(people));

            // Assert
            var expectedOutput = "Name: Jane Smith, Age: 25, City: Los Angeles\r\nName: John Doe, Age: 30, City: New York";
            Assert.AreEqual(expectedOutput, output);
        }
    }
}
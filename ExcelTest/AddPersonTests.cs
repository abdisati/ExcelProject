using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelLibray;
using ExcelDataReaderApp;

namespace ExcelDataReaderTest
{
    [TestClass]
    public class AddPersonTests
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
        public void TestAddPerson_ValidPerson()
        {
            // Arrange
            var people = new Dictionary<string, Person>();
            var person = new Person { Name = "John Doe", Age = 30, City = "New York" };

            // Act
            ExcelReaderProgram.AddPerson(people, person);

            // Assert
            Assert.AreEqual(1, people.Count);
            Assert.IsTrue(people.ContainsKey("john doe-30-new york"));
        }

        [TestMethod]
        public void TestAddPerson_DuplicatePerson()
        {
            // Arrange
            var people = new Dictionary<string, Person>();
            var person1 = new Person { Name = "John Doe", Age = 30, City = "New York" };
            var person2 = new Person { Name = "John Doe", Age = 30, City = "New York" };

            // Act
            var output = CaptureConsoleOutput(() =>
            {
                ExcelReaderProgram.AddPerson(people, person1);
                ExcelReaderProgram.AddPerson(people, person2);
            });

            // Assert
            Assert.AreEqual(1, people.Count); // Only one unique person should be added
            Assert.IsTrue(output.Contains("Duplicate entry detected: John Doe, 30, New York")); // Verify log message
        }

        [TestMethod]
        public void TestAddPerson_CaseInsensitive()
        {
            // Arrange
            var people = new Dictionary<string, Person>();
            var person1 = new Person { Name = "John Doe", Age = 30, City = "New York" };
            var person2 = new Person { Name = "john doe", Age = 30, City = "new york" };

            // Act
            ExcelReaderProgram.AddPerson(people, person1);
            ExcelReaderProgram.AddPerson(people, person2);

            // Assert
            Assert.AreEqual(1, people.Count);
        }
    }
}
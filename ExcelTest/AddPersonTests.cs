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
        public void TestAddPerson_NullPerson()
        {
            // Arrange
            var people = new Dictionary<string, Person>();
            var originalOut = Console.Out;

            using (var sw = new StringWriter())
            {
                Console.SetOut(sw);

                // Act
                ExcelReaderProgram.AddPerson(people, null);

                // Assert
                Assert.AreEqual(0, people.Count);
                var output = sw.ToString().Trim();
                Assert.IsTrue(output.Contains("Cannot add a null person to the dictionary."));

                // Reset the console output
                Console.SetOut(originalOut);
            }
        }

        [TestMethod]
        public void TestAddPerson_DuplicatePerson()
        {
            // Arrange
            var people = new Dictionary<string, Person>();
            var person1 = new Person { Name = "John Doe", Age = 30, City = "New York" };
            var person2 = new Person { Name = "John Doe", Age = 30, City = "New York" };

            var originalOut = Console.Out;
            using (var sw = new StringWriter())
            {
                Console.SetOut(sw); // Redirect Console.Out to StringWriter

                try
                {
                    // Act
                    ExcelReaderProgram.AddPerson(people, person1);
                    ExcelReaderProgram.AddPerson(people, person2);

                    // Assert
                    Assert.AreEqual(1, people.Count); // Only one unique person should be added
                    var output = sw.ToString().Trim();
                    Assert.IsTrue(output.Contains("Duplicate entry detected: John Doe, 30, New York")); // Verify log message
                }
                finally
                {
                    // Ensure Console.Out is reset even if an exception occurs
                    Console.SetOut(originalOut);
                }
            }
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
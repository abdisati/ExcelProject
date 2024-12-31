using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelLibray;
using ExcelDataReaderApp;

namespace ExcelDataReaderTest{

[TestClass]
public class PrintPeopleTests
{
    [TestMethod]
    public void TestPrintPeople()
    {
        // Arrange
        var people = new Dictionary<string, Person>
        {
            { "john doe-30-new york", new Person { Name = "John Doe", Age = 30, City = "New York" } },
            { "jane smith-25-los angeles", new Person { Name = "Jane Smith", Age = 25, City = "Los Angeles" } }
        };

        using (var sw = new StringWriter())
        {
            Console.SetOut(sw);

            // Act
            ExcelReaderProgram.PrintPeople(people);

            // Assert
            var output = sw.ToString().Trim();
            var expectedOutput = "Name: Jane Smith, Age: 25, City: Los Angeles\nName: John Doe, Age: 30, City: New York";
            Assert.AreEqual(expectedOutput, output);
        }
    }

    [TestMethod]
    public void TestPrintPeople_EmptyDictionary()
    {
        // Arrange
        var people = new Dictionary<string, Person>();

        using (var sw = new StringWriter())
        {
            Console.SetOut(sw);

            // Act
            ExcelReaderProgram.PrintPeople(people);

            // Assert
            var output = sw.ToString().Trim();
            Assert.AreEqual(string.Empty, output);
        }
    }
}}
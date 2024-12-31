using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelLibray;
using ExcelDataReaderApp;
namespace ExcelDataReaderTest{

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

        // Act
        ExcelReaderProgram.AddPerson(people, null);

        // Assert
        Assert.AreEqual(0, people.Count);
    }

    [TestMethod]
    public void TestAddPerson_DuplicatePerson()
    {
        // Arrange
        var people = new Dictionary<string, Person>();
        var person1 = new Person { Name = "John Doe", Age = 30, City = "New York" };
        var person2 = new Person { Name = "John Doe", Age = 30, City = "New York" };

        // Act
        ExcelReaderProgram.AddPerson(people, person1);
        ExcelReaderProgram.AddPerson(people, person2);

        // Assert
        Assert.AreEqual(1, people.Count);
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
}}
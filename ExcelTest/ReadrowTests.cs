using System;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using ExcelLibray;
using ExcelDataReaderApp;

namespace ExcelDataReaderTest{

[TestClass]
public class ReadRowTests
{
    [TestMethod]
    public void TestReadRow_ValidData()
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
    public void TestReadRow_InvalidAge()
    {
        // Arrange
        var mockRow = CreateMockRow("John Doe", "invalid_age", "New York");

        // Act
        var person = ExcelReaderProgram.ReadRow(mockRow);

        // Assert
        Assert.AreEqual("John Doe", person.Name);
        Assert.AreEqual(-1, person.Age); // Default value for invalid age
        Assert.AreEqual("New York", person.City);
    }

    [TestMethod]
    public void TestReadRow_EmptyCity()
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

    private static IXLRow CreateMockRow(string name, object age, string city)
    {
        var mockRow = new Mock<IXLRow>();
        var mockNameCell = new Mock<IXLCell>();
        var mockAgeCell = new Mock<IXLCell>();
        var mockCityCell = new Mock<IXLCell>();

        mockNameCell.Setup(cell => cell.GetValue<string>()).Returns(name);
        mockAgeCell.Setup(cell => cell.GetValue<object>()).Returns(age);
        mockCityCell.Setup(cell => cell.GetValue<string>()).Returns(city);

        mockRow.Setup(row => row.Cell(1)).Returns(mockNameCell.Object);
        mockRow.Setup(row => row.Cell(2)).Returns(mockAgeCell.Object);
        mockRow.Setup(row => row.Cell(3)).Returns(mockCityCell.Object);

        return mockRow.Object;
    }
}}
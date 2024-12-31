using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using ExcelLibray;
using ExcelDataReaderApp;

public interface IFileSystem
{
    bool Exists(string path);
}

[TestClass]
public class LoadPeopleFromExcelTests
{
    [TestMethod]
    public void TestLoadPeopleFromExcel_ValidFile()
    {
        // Arrange
        var mockWorkbook = new Mock<IXLWorkbook>();
        var mockWorksheet = new Mock<IXLWorksheet>();
        var mockRows = new List<IXLRow>
        {
            CreateMockRow("John Doe", 30, "New York"),
            CreateMockRow("Jane Smith", 25, "Los Angeles")
        };

        mockWorkbook.Setup(wb => wb.Worksheet(1)).Returns(mockWorksheet.Object);
        var mockRowsUsed = new Mock<IXLRows>();
        mockWorksheet.Setup(ws => ws.RowsUsed(It.IsAny<XLCellsUsedOptions>(), It.IsAny<Func<IXLRow, bool>>())).Returns(mockRowsUsed.Object);

        // Mock the file path
        var mockFilePath = "mockedFilePath.xlsx";
        var mockFile = new Mock<IFileSystem>();
        mockFile.Setup(f => f.Exists(mockFilePath)).Returns(true);

        // Act
        var people = ExcelReaderProgram.LoadPeopleFromExcel(mockFilePath);

        // Assert
        Assert.AreEqual(2, people.Count);
        Assert.IsTrue(people.ContainsKey("john doe-30-new york"));
        Assert.IsTrue(people.ContainsKey("jane smith-25-los angeles"));
    }

    [TestMethod]
    public void TestLoadPeopleFromExcel_InvalidFile()
    {
        // Arrange
        var mockFilePath = "invalidFilePath.xlsx";
        var mockFile = new Mock<IFileSystem>();
        mockFile.Setup(f => f.Exists(mockFilePath)).Returns(false);

        // Act
        var people = ExcelReaderProgram.LoadPeopleFromExcel(mockFilePath);

        // Assert
        Assert.AreEqual(0, people.Count);
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
}
using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using ExcelLibray;

[TestClass]
public class ProgramTest
{
    [TestMethod]
    public void TestExcelReading()
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
        mockRowsUsed.Setup(r => r.GetEnumerator()).Returns(mockRows.GetEnumerator());
        mockWorksheet.Setup(ws => ws.RowsUsed(It.IsAny<XLCellsUsedOptions>(), It.IsAny<Func<IXLRow, bool>>()))
            .Returns(mockRowsUsed.Object);
        mockWorksheet.Setup(ws => ws.RowsUsed(It.IsAny<Func<IXLRow, bool>>()))
            .Returns(mockRowsUsed.Object);
        mockWorksheet.Setup(ws => ws.RowsUsed(XLCellsUsedOptions.AllContents, null))
               .Returns(mockRowsUsed.Object);


        var people = new List<Person>();

        // Act
        using (var workbook = mockWorkbook.Object)
        {
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RowsUsed();

            foreach (var row in rows)
            {
                string name = row.Cell(1).GetValue<string>();
                int age = row.Cell(2).GetValue<int>();
                string city = row.Cell(3).GetValue<string>();

                people.Add(new Person { Name = name, Age = age, City = city });
            }
        }

        // Assert
        Assert.AreEqual(2, people.Count);
        Assert.AreEqual("John Doe", people[0].Name);
        Assert.AreEqual(30, people[0].Age);
        Assert.AreEqual("New York", people[0].City);
    }

    private static IXLRow CreateMockRow(string name, int age, string city)
    {
        var mockRow = new Mock<IXLRow>();
        mockRow.Setup(row => row.Cell(1).GetValue<string>()).Returns(name);
        mockRow.Setup(row => row.Cell(2).GetValue<int>()).Returns(age);
        mockRow.Setup(row => row.Cell(3).GetValue<string>()).Returns(city);
        return mockRow.Object;
    }
}
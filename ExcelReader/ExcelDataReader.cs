using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using ExcelLibray;

namespace ExcelDataReaderApp
{

    class ExcelReaderProgram
    {
        static void Main()
        {
            // Path to your Excel file
            string filePath = "C:\\Users\\v-abdideresa\\Desktop\\Data.xlsx"; ;

            // List to store rows as Person objects
            List<Person> people = new List<Person>();

            // Open the Excel file
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1); // Get the first sheet
                var rows = worksheet.RowsUsed();      // Get rows with data

                // Skip the header row
                foreach (var row in rows.Skip(1))
                {
                    // Read each cell
                    string name = row.Cell(1).GetValue<string>();
                    int age = row.Cell(2).GetValue<int>();
                    string city = row.Cell(3).GetValue<string>();

                    // Create a new Person object and add to the list
                    people.Add(new Person { Name = name, Age = age, City = city });
                }
            }

            // Print the contents of the list
            foreach (var person in people)
            {
                Console.WriteLine(person);
            }
        }
    }
}

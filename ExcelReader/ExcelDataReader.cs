using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using ExcelLibray;
using System.IO;

namespace ExcelDataReaderApp
{
    class ExcelReaderProgram
    {
        static void Main()
        {
            // Accept an input string from the user specifying the Excel file to open
            Console.WriteLine("Enter the path to the Excel file:");
            string? filePath = Console.ReadLine()?.Trim('"');

            // Call LoadPeopleFromExcel function to process the file
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                Console.WriteLine("Invalid file path.");
                return;
            }

            Dictionary<string, Person> people = LoadPeopleFromExcel(filePath);

            // Print the contents of the dictionary
            PrintPeople(people);
        }

        static Dictionary<string, Person> LoadPeopleFromExcel(string filePath)
        {
            Dictionary<string, Person> people = new Dictionary<string, Person>();

            // Open the Excel file
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1); // Get the first sheet
                var rows = worksheet.RowsUsed();      // Get rows with data

                // Skip the header row
                foreach (var row in rows.Skip(1))
                {
                    // Read each row and add to the dictionary
                    Person person = ReadRow(row);
                    AddPerson(people, person);
                }
            }

            return people;
        }

        static Person ReadRow(IXLRow row)
        {
            string name = row.Cell(1).GetValue<string>();
            int age;
            string city = row.Cell(3).GetValue<string>();

            try
            {
                age = row.Cell(2).GetValue<int>();
            }
            catch
            {
                age = -1; // Default value for invalid age
            }

            if (string.IsNullOrEmpty(city))
            {
                city = "Unknown"; // Default value for invalid city
            }

            return new Person { Name = name, Age = age, City = city };
        }

        static void AddPerson(Dictionary<string, Person> people, Person person)
        {
            string key = $"{person.Name}-{person.Age}-{person.City}";

            if (!people.ContainsKey(key))
            {
                people.Add(key, person);
            }
        }

        static void PrintPeople(Dictionary<string, Person> people)
        {
            foreach (var person in people.Values.OrderBy(p => p.Name))
            {
                Console.WriteLine(person);
            }
        }
    }
}

using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using ExcelLibray;
using System.IO;

namespace ExcelDataReaderApp
{
    public class ExcelReaderProgram
    {
        static void Main()
        {
            // Accept an input string from the user specifying the Excel file to open
            Console.WriteLine("Enter the path to the Excel file:");
            string? filePath = Console.ReadLine()?.Trim('"');

            // Check if the file path is valid
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                Console.WriteLine("Invalid file path.");
                return;
            }

            // Load the people from the Excel file

            Dictionary<string, Person> people = LoadPeopleFromExcel(filePath);

            // Print the contents of the dictionary
            PrintPeople(people);
        }

        public static Dictionary<string, Person> LoadPeopleFromExcel(string filePath)
        {
            Dictionary<string, Person> people = new Dictionary<string, Person>();

            try
            {
                // Open the Excel file
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(1); // Get the first sheet
                    var rows = worksheet.RowsUsed();      // Get rows with data

                    // Skip the header row
                    foreach (var row in rows.Skip(1))
                    {
                        try
                        {
                            // Read each row and add to the dictionary
                            Person person = ReadRow(row);
                            AddPerson(people, person);
                        }
                        catch (Exception ex)
                        {
                            LogError($"Error processing row: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogError($"Error loading Excel file: {ex.Message}");
            }

            return people;
        }

        public static Person ReadRow(IXLRow row)
        {
            string? name;
            int age;
            string? city;

            try
            {
                // Safely get the name and handle null/empty values
                name = row.Cell(1).GetValue<string>()?.Trim();
                if (string.IsNullOrEmpty(name))
                {
                    name = "Unknown"; // Default for invalid name
                }
            }
            catch (Exception ex)
            {
                LogError($"Error reading 'Name' from row: {ex.Message}");
                name = "Unknown";
            }

            try
            {
                // Safely parse age and handle invalid formats
                age = row.Cell(2).GetValue<int>();
            }
            catch (FormatException ex)
            {
                LogError($"Invalid format for 'Age' in row: {ex.Message}");
                age = -1; // Default value for invalid age
            }
            catch (InvalidCastException ex)
            {
                LogError($"Invalid data type for 'Age' in row: {ex.Message}");
                age = -1;
            }
            catch (Exception ex)
            {
                LogError($"Unexpected error reading 'Age' from row: {ex.Message}");
                age = -1;
            }

            try
            {
                // Safely get the city and handle null/empty values
                city = row.Cell(3).GetValue<string>()?.Trim();
                if (string.IsNullOrEmpty(city))
                {
                    city = "Unknown"; // Default for invalid city
                }
            }
            catch (Exception ex)
            {
                LogError($"Error reading 'City' from row: {ex.Message}");
                city = "Unknown";
            }

            return new Person { Name = name, Age = age, City = city };
        }

        // Log an error message
        private static void LogError(string message)
        {
            Console.WriteLine(message); // Print the error message
        }

        public static void AddPerson(Dictionary<string, Person> people, Person? person)
        {
            if (person == null)
            {
                LogError("Cannot add a null person to the dictionary.");
                return;
            }

            string key = $"{person.Name?.Trim().ToLower()}-{person.Age}-{person.City?.Trim().ToLower()}";

            // Check if the key already exists in the dictionary
            if (!people.ContainsKey(key))
            {
                people.Add(key, person);
            }
            else
            {
                LogError($"Duplicate entry detected: {person.Name}, {person.Age}, {person.City}");
            }
        }

        public static void PrintPeople(Dictionary<string, Person> people)
        {
            foreach (var person in people.Values.OrderBy(p => p.Name))
            {
                Console.WriteLine(person);
            }
        }
    }
}

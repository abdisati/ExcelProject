using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ExcelLibray;

namespace ExcelDataReaderApp
{
    public class ExcelReaderProgram
    {
        static void Main()
        {
            // Accept input strings from the user specifying the Excel files and sheets to open
            Console.WriteLine("Enter the path to the first Excel file:");
            string? filePath1 = Console.ReadLine()?.Trim('"');
            Console.WriteLine("Enter the sheet name for the first Excel file:");
            string? sheetName1 = Console.ReadLine()?.Trim();

            Console.WriteLine("Enter the path to the second Excel file:");
            string? filePath2 = Console.ReadLine()?.Trim('"');
            Console.WriteLine("Enter the sheet name for the second Excel file:");
            string? sheetName2 = Console.ReadLine()?.Trim();

            // Check if the file paths are valid
            if (string.IsNullOrEmpty(filePath1) || !File.Exists(filePath1) ||
                string.IsNullOrEmpty(filePath2) || !File.Exists(filePath2))
            {
                Console.WriteLine("Invalid file path(s).");
                return;
            }

            // Check if the sheet names are valid
            if (string.IsNullOrEmpty(sheetName1) || string.IsNullOrEmpty(sheetName2))
            {
                Console.WriteLine("Invalid sheet name(s).");
                return;
            }

            // Load the people from the Excel files
            Dictionary<string, List<Person>> peopleSheet1 = LoadPeopleFromExcel(filePath1, sheetName1);
            Dictionary<string, List<Person>> peopleSheet2 = LoadPeopleFromExcel(filePath2, sheetName2);

            // Compare the data from the two sheets
            CompareAndPrintResults(peopleSheet1, peopleSheet2);
        }

        public static Dictionary<string, List<Person>> LoadPeopleFromExcel(string filePath, string sheetName)
        {
            Dictionary<string, List<Person>> people = new Dictionary<string, List<Person>>();

            try
            {
                // Open the Excel file
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(sheetName); // Get the specified sheet
                    var rows = worksheet.RowsUsed();               // Get rows with data

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

        public static void AddPerson(Dictionary<string, List<Person>> people, Person person)
        {
            string key = (person.Name ?? "unknown").Trim().ToLower();

            if (!people.ContainsKey(key))
            {
                people[key] = new List<Person>();
            }

            people[key].Add(person);
        }

        public static void CompareAndPrintResults(Dictionary<string, List<Person>> peopleSheet1, Dictionary<string, List<Person>> peopleSheet2)
        {
            var allNames = new HashSet<string>(peopleSheet1.Keys.Concat(peopleSheet2.Keys));

            foreach (var name in allNames)
            {
                bool inSheet1 = peopleSheet1.TryGetValue(name, out var persons1);
                bool inSheet2 = peopleSheet2.TryGetValue(name, out var persons2);

                if (inSheet1 && inSheet2)
                {
                    foreach (var person1 in persons1!)
                    {
                        foreach (var person2 in persons2!)
                        {
                            if (person1.Name == person2.Name && person1.Age == person2.Age && person1.City == person2.City)
                            {
                                Console.WriteLine($"Exact match: {person1}");
                            }
                            else
                            {
                                Console.WriteLine($"Partial match for {person1.Name}:");
                                if (person1.Age != person2.Age)
                                {
                                    Console.WriteLine($"  Age differs: Sheet1={person1.Age}, Sheet2={person2.Age}");
                                }
                                if (person1.City != person2.City)
                                {
                                    Console.WriteLine($"  City differs: Sheet1={person1.City}, Sheet2={person2.City}");
                                }
                            }
                        }
                    }
                }
                else if (inSheet1)
                {
                    foreach (var person in persons1!)
                    {
                        Console.WriteLine($"Only in Sheet1: {person}");
                    }
                }
                else if (inSheet2)
                {
                    foreach (var person in persons2!)
                    {
                        Console.WriteLine($"Only in Sheet2: {person}");
                    }
                }
            }
        }
    }
}

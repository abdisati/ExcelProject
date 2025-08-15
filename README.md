# 📊 Excel Data Comparison Tool

**ExcelDataReaderApp** is a C# console application that compares data from two Excel sheets and outputs detailed differences.  
It helps detect **exact matches**, **partial matches**, and **records unique** to one sheet, making it ideal for data validation, migration checks, or audit purposes.

---

## ✨ Features

- **Compare two Excel sheets** row by row based on:
  - Name
  - Age
  - City
- **Detect:**
  - ✅ Exact matches (all fields match)
  - ⚠ Partial matches (same name but different age or city)
  - ❌ Records only found in one sheet
- **Error handling** for:
  - Missing or invalid file paths
  - Missing or invalid sheet names
  - Malformed or incomplete row data
- **Automatic defaults** for invalid data (`Unknown` for text, `-1` for numbers)

---

## 📂 Project Structure

ExcelProject/
├── ExcelReaderApp/ # Main console application
│ └── ExcelReaderProgram.cs # Core program logic
├── ExcelLibrary/ # Shared models and helpers (e.g., Person class)
├── ExcelTest/ # Unit tests (if applicable)
├── ExcelProject.sln # Solution file
└── README.md # Documentation

▶️ How to Run

Clone this repository:

git clone https://github.com/abdisati/ExcelProject.git
cd ExcelProject


Run the application:

dotnet run --project ExcelReaderApp


Follow the prompts:

Enter the full path to the first Excel file.

Enter the sheet name for the first file.

Enter the full path to the second Excel file.

Enter the sheet name for the second file.

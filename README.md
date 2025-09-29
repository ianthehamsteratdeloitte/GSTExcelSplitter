# GST Excel Splitter

A C# WinForms tool built with **Microsoft.Office.Interop.Excel** to split large Excel files into smaller batches (default batchSize: 3,600 rows) for GST Registration checks.

## Features
- Splits client Excel files into chunks of 3,600 rows.
- Uses **bulk range copy** for high performance (much faster than per-cell loops).
- Automatically creates new Excel files: `Batch_1.xlsx`, `Batch_2.xlsx`, etc.
- Handles datasets with tens of thousands of rows.

## Requirements
- Windows with Microsoft Excel installed.
- .NET Framework (4.7+ recommended).
- Visual Studio Community Edition.

## Usage
1. Place the client Excel file in a known folder (e.g., `C:\Temp`).
2. Update the file path in `Form1.cs`:
```csharp
   string sourceFile = @"C:\Temp\ClientData.xlsx";
```

3. Run the tool → it will output split files into the configured folder.

## Future Improvements

* Add file picker dialog instead of hardcoded path.
* Add progress bar and status logs.
* Option to auto-consolidate split files with Power Query.

---



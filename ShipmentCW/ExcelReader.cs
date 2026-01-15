using OfficeOpenXml;

namespace ShipmentCW;

/// <summary>
/// Provides functionality to read data from Excel files using EPPlus
/// </summary>
public class ExcelReader
{
    // Store the path to the Excel file
    private readonly string _filePath;

    /// <summary>
    /// Initializes a new instance of the ExcelReader class
    /// </summary>
    /// <param name="filePath">Full path to the Excel file to be read</param>
    public ExcelReader(string filePath)
    {
        _filePath = filePath;
    }

    /// <summary>
    /// Reads all data from the specified worksheet
    /// </summary>
    /// <param name="worksheetIndex">Worksheet index (1-based)</param>
    /// <returns>List of rows, where each row is a list of cell values</returns>
    public List<List<string>> ReadWorksheet(int worksheetIndex = 1)
    {
        // Initialize a list to store all rows of data
        var data = new List<List<string>>();

        // Verify the Excel file exists before attempting to open it
        if (!File.Exists(_filePath))
        {
            throw new FileNotFoundException($"Excel file not found: {_filePath}");
        }

        // Open the Excel file using EPPlus (automatically disposed via 'using')
        using var package = new ExcelPackage(new FileInfo(_filePath));

        // Validate that the requested worksheet index exists in the workbook
        if (package.Workbook.Worksheets.Count < worksheetIndex)
        {
            throw new ArgumentException($"Worksheet index {worksheetIndex} not found. File has {package.Workbook.Worksheets.Count} worksheet(s).");
        }

        // Get the worksheet (EPPlus uses 0-based indexing internally, so subtract 1)
        var worksheet = package.Workbook.Worksheets[worksheetIndex - 1];

        // Get the dimensions of the worksheet (number of rows and columns with data)
        // Using null-coalescing operator to handle empty worksheets
        var rowCount = worksheet.Dimension?.Rows ?? 0;
        var colCount = worksheet.Dimension?.Columns ?? 0;

        // Iterate through each row in the worksheet (EPPlus uses 1-based indexing for cells)
        for (int row = 1; row <= rowCount; row++)
        {
            var rowData = new List<string>();

            // Iterate through each column in the current row
            for (int col = 1; col <= colCount; col++)
            {
                // Get the cell value and convert to string (empty string if null)
                var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;
                rowData.Add(cellValue);
            }

            // Add the complete row to the data collection
            data.Add(rowData);
        }

        return data;
    }

    /// <summary>
    /// Reads all data from the specified worksheet by name
    /// </summary>
    /// <param name="worksheetName">Name of the worksheet</param>
    /// <returns>List of rows, where each row is a list of cell values</returns>
    public List<List<string>> ReadWorksheetByName(string worksheetName)
    {
        // Initialize a list to store all rows of data
        var data = new List<List<string>>();

        // Verify the Excel file exists before attempting to open it
        if (!File.Exists(_filePath))
        {
            throw new FileNotFoundException($"Excel file not found: {_filePath}");
        }

        // Open the Excel file using EPPlus
        using var package = new ExcelPackage(new FileInfo(_filePath));

        // Attempt to retrieve the worksheet by name
        var worksheet = package.Workbook.Worksheets[worksheetName];

        // Verify that the named worksheet exists
        if (worksheet == null)
        {
            throw new ArgumentException($"Worksheet '{worksheetName}' not found.");
        }

        // Get the dimensions of the worksheet (number of rows and columns with data)
        var rowCount = worksheet.Dimension?.Rows ?? 0;
        var colCount = worksheet.Dimension?.Columns ?? 0;

        // Iterate through each row in the worksheet
        for (int row = 1; row <= rowCount; row++)
        {
            var rowData = new List<string>();

            // Iterate through each column in the current row
            for (int col = 1; col <= colCount; col++)
            {
                // Get the cell value and convert to string (empty string if null)
                var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;
                rowData.Add(cellValue);
            }

            // Add the complete row to the data collection
            data.Add(rowData);
        }

        return data;
    }

    /// <summary>
    /// Gets the names of all worksheets in the Excel file
    /// </summary>
    /// <returns>List of worksheet names</returns>
    public List<string> GetWorksheetNames()
    {
        // Verify the Excel file exists before attempting to open it
        if (!File.Exists(_filePath))
        {
            throw new FileNotFoundException($"Excel file not found: {_filePath}");
        }

        // Open the Excel file and extract all worksheet names using LINQ
        using var package = new ExcelPackage(new FileInfo(_filePath));
        return package.Workbook.Worksheets.Select(ws => ws.Name).ToList();
    }

    /// <summary>
    /// Prints the data from a worksheet to the console
    /// </summary>
    /// <param name="worksheetIndex">Worksheet index (1-based)</param>
    public void PrintWorksheet(int worksheetIndex = 1)
    {
        // Read all data from the specified worksheet
        var data = ReadWorksheet(worksheetIndex);

        // Display header with worksheet information
        Console.WriteLine($"\n=== Worksheet {worksheetIndex} ===");
        Console.WriteLine($"Total rows: {data.Count}");

        // Print each row with cells separated by pipes
        foreach (var row in data)
        {
            Console.WriteLine(string.Join(" | ", row));
        }
    }
}

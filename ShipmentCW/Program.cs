using OfficeOpenXml;
using ShipmentCW;

class Program
{
    static void Main()
    {
        // Set EPPlus license for non-commercial personal use (EPPlus 8.x)
        // This is required when using EPPlus version 8.0 and above
        ExcelPackage.License.SetNonCommercialPersonal("ShipmentCW User");

        // Define the path to the Excel file to be read
        string excelFilePath = @"C:\temp\blTemplate.xlsx";

        // Display application header
        Console.WriteLine("ShipmentCW - Excel Reader");
        Console.WriteLine("=========================\n");

        try
        {
            // Initialize the Excel reader with the file path
            var excelReader = new ExcelReader(excelFilePath);

            // Retrieve and display all available worksheet names in the Excel file
            Console.WriteLine("Available worksheets:");
            var worksheetNames = excelReader.GetWorksheetNames();
            for (int i = 0; i < worksheetNames.Count; i++)
            {
                // Display worksheet index and name (1-based indexing for user-friendly display)
                Console.WriteLine($"  {i + 1}. {worksheetNames[i]}");
            }

            // Read all data from the first worksheet (index 1)
            Console.WriteLine("\nReading data from first worksheet...");
            var data = excelReader.ReadWorksheet(1);

            // Display the total number of rows found
            Console.WriteLine($"\nFound {data.Count} rows in the worksheet.\n");

            // Display a preview of the data (limited to first 10 rows)
            Console.WriteLine("Data preview:");
            Console.WriteLine(new string('-', 80));

            // Determine how many rows to display (max 10 or total rows if less)
            int rowsToDisplay = Math.Min(10, data.Count);
            for (int i = 0; i < rowsToDisplay; i++)
            {
                // Display each row with cells separated by pipes
                Console.WriteLine($"Row {i + 1}: {string.Join(" | ", data[i])}");
            }

            // If there are more than 10 rows, indicate how many are not shown
            if (data.Count > 10)
            {
                Console.WriteLine($"\n... and {data.Count - 10} more rows");
            }

            Console.WriteLine(new string('-', 80));
        }
        catch (FileNotFoundException ex)
        {
            // Handle case where the Excel file doesn't exist at the specified path
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine("Please make sure the Excel file exists at the specified path.");
        }
        catch (Exception ex)
        {
            // Handle any other exceptions that might occur during Excel processing
            Console.WriteLine($"Error reading Excel file: {ex.Message}");
        }

        // Wait for user input before closing the console window
        Console.WriteLine("\nPress any key to exit...");
        Console.ReadKey();
    }
}

using ClosedXML.Excel;
using System.Text.Json;

namespace ExcelToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            // Ask the user for the Excel file path
            Console.Write("Enter Excel file path: ");
            string filePath = Console.ReadLine();

            // Check if the file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine("File not found!");
                return;
            }

            // Open the Excel file
            XLWorkbook workbook = new XLWorkbook(filePath);

            // Get the first worksheet
            var sheet = workbook.Worksheet(1);

            // Get the used rows
            var rows = sheet.RowsUsed();

            // Get the first row as headers
            var firstRow = rows.First();

            List<string> headers = new List<string>();

            foreach (var cell in firstRow.CellsUsed())
            {
                headers.Add(cell.GetString());
            }

            // Create a list to store the data
            List<Dictionary<string, string>> data =
                new List<Dictionary<string, string>>();

            // Read each row after the first row
            foreach (var row in rows.Skip(1))
            {
                Dictionary<string, string> item =
                    new Dictionary<string, string>();

                for (int i = 0; i < headers.Count; i++)
                {
                    string header = headers[i];
                    string value = row.Cell(i + 1).GetString();

                    item.Add(header, value);
                }

                data.Add(item);
            }

            // Convert the data to JSON
            string json = JsonSerializer.Serialize(
                data,
                new JsonSerializerOptions
                {
                    WriteIndented = true
                }
            );

            // Create JSON file path
            string jsonPath = Path.ChangeExtension(filePath, ".json");

            // Save JSON file
            File.WriteAllText(jsonPath, json);

            // Show result
            Console.WriteLine();
            Console.WriteLine("Conversion completed!");
            Console.WriteLine("JSON file created at:");
            Console.WriteLine(jsonPath);

            // Close Excel file
            workbook.Dispose();
        }
    }
}
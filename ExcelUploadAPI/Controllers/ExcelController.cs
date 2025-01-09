using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelUploadAPI.Controllers
{
    [Route("api/excel")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        // POST api/excel/upload
        [HttpPost("upload")]
        public async Task<IActionResult> UploadExcelFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                Console.WriteLine("No file was uploaded.");
                return BadRequest("No file uploaded.");
            }

            try
            {
                Console.WriteLine("Processing uploaded file directly from stream...");

                using (var stream = file.OpenReadStream())
                {
                    // Use ClosedXML to process the Excel file
                    using (var workbook = new XLWorkbook(stream))
                    {
                        Console.WriteLine("Excel file opened successfully.");
                        var worksheet = workbook.Worksheet(1); // Get the first worksheet
                        Console.WriteLine($"Processing worksheet: {worksheet.Name}");

                        var dataTable = new DataTable();

                        // Read column headers (first row)
                        var firstRow = worksheet.Row(1);
                        foreach (var cell in firstRow.CellsUsed())
                        {
                            dataTable.Columns.Add(cell.Value.ToString());
                            Console.WriteLine($"Column added: {cell.Value}");
                        }

                        // Read the data (starting from the second row)
                        foreach (var row in worksheet.RowsUsed().Skip(1))
                        {
                            var dataRow = dataTable.NewRow();
                            int columnIndex = 0;
                            foreach (var cell in row.CellsUsed())
                            {
                                dataRow[columnIndex++] = cell.Value.ToString();
                            }
                            dataTable.Rows.Add(dataRow);
                        }

                        Console.WriteLine($"Total rows added: {dataTable.Rows.Count}");

                        // Implement the importer logic to save dataTable to the database
                        string connectionString = "YourChoice"; // Replace with your connection string
                        string tableName = "YourChoice"; // Replace with your table name
                        ImportToDatabase(dataTable, connectionString, tableName);
                    }
                }

                Console.WriteLine("File processed directly from stream and data imported successfully.");
                return Ok("File processed directly from stream and data imported successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Internal server error: {ex.Message}");
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        private void ImportToDatabase(DataTable dataTable, string connectionString, string tableName)
        {
            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    Console.WriteLine("Database connection opened.");

                    // Clear the table before importing new data
                    var clearTableQuery = $"DELETE FROM {tableName}";
                    using (var clearCommand = new SqlCommand(clearTableQuery, connection))
                    {
                        clearCommand.ExecuteNonQuery();
                        Console.WriteLine($"Table {tableName} cleared.");
                    }

                    // Check if the table exists; create it if it doesn't
                    var createTableQuery = GenerateCreateTableQuery(dataTable, tableName);
                    using (var createCommand = new SqlCommand(createTableQuery, connection))
                    {
                        createCommand.ExecuteNonQuery();
                        Console.WriteLine($"Table {tableName} created or verified.");
                    }

                    // Use SqlBulkCopy to insert data into the table
                    using (var bulkCopy = new SqlBulkCopy(connection))
                    {
                        bulkCopy.BulkCopyTimeout = 900;
                        bulkCopy.DestinationTableName = tableName;
                        bulkCopy.WriteToServer(dataTable);
                        Console.WriteLine($"Data imported to table {tableName} successfully.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during data import: {ex.Message}");
                throw;
            }
        }

        private string GenerateCreateTableQuery(DataTable dataTable, string tableName)
        {
            var query = $"IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{tableName}') BEGIN CREATE TABLE {tableName} (";

            foreach (DataColumn column in dataTable.Columns)
            {
                query += $"[{column.ColumnName}] NVARCHAR(MAX),";
                Console.WriteLine($"Generating column for SQL table: {column.ColumnName}");
            }

            query = query.TrimEnd(',') + "); END";
            return query;
        }
    }
}

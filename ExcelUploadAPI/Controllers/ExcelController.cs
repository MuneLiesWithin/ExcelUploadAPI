using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
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
        private readonly IConfiguration _configuration;

        public ExcelController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

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
                            dataTable.Columns.Add(cell.Value.ToString()); // Handle null values
                            Console.WriteLine($"Column added: {cell.Value}");
                        }

                        // Read the data (starting from the second row)
                        foreach (var row in worksheet.RowsUsed().Skip(1))
                        {
                            var dataRow = dataTable.NewRow();
                            int columnIndex = 0;
                            foreach (var cell in row.CellsUsed())
                            {
                                dataRow[columnIndex++] = cell.Value.ToString(); // Handle null values
                            }
                            dataTable.Rows.Add(dataRow);
                        }

                        Console.WriteLine($"Total rows added: {dataTable.Rows.Count}");

                        // Fetch connection string and table name from configuration
                        string? connectionString = _configuration.GetConnectionString("DefaultConnection");
                        string? tableName = _configuration["AppSettings:TableName"];

                        if (string.IsNullOrWhiteSpace(connectionString) || string.IsNullOrWhiteSpace(tableName))
                        {
                            Console.WriteLine("Invalid connection string or table name.");
                            return StatusCode(500, "Invalid configuration settings.");
                        }

                        // Import data into the database
                        await Task.Run(() => ImportToDatabase(dataTable, connectionString, tableName));
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

                    // Check if the table exists
                    var tableExistsQuery = $"IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{tableName}') SELECT 1 ELSE SELECT 0";
                    bool tableExists;

                    using (var tableExistsCommand = new SqlCommand(tableExistsQuery, connection))
                    {
                        tableExists = (int)tableExistsCommand.ExecuteScalar() == 1;
                    }

                    if (tableExists)
                    {
                        // Clear the table if it exists
                        var clearTableQuery = $"DELETE FROM {tableName}";
                        using (var clearCommand = new SqlCommand(clearTableQuery, connection))
                        {
                            clearCommand.ExecuteNonQuery();
                            Console.WriteLine($"Table {tableName} cleared.");
                        }
                    }
                    else
                    {
                        // Create the table if it does not exist
                        var createTableQuery = GenerateCreateTableQuery(dataTable, tableName);
                        using (var createCommand = new SqlCommand(createTableQuery, connection))
                        {
                            createCommand.ExecuteNonQuery();
                            Console.WriteLine($"Table {tableName} created.");
                        }
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
                // Trim the column name to avoid spaces
                var columnName = column.ColumnName.Trim();

                // Ensure the column name is valid for SQL
                if (string.IsNullOrWhiteSpace(columnName))
                {
                    throw new Exception("Column name cannot be empty or whitespace.");
                }

                query += $"[{columnName}] NVARCHAR(MAX),";
                Console.WriteLine($"Generating column for SQL table: {columnName}");
            }

            query = query.TrimEnd(',') + "); END";
            return query;
        }
    }
}

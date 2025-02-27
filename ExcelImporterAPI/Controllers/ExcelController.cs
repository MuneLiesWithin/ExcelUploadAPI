using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net;
using System.Web.Http;
using ClosedXML.Excel;
using System.Threading;
using System.Diagnostics;
using System.Web.Http.Results;

namespace ExcelImporterAPI.Controllers
{
    public class ExcelController : ApiController
    {
        // GET api/Excel
        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // POST api/excel/upload
        [HttpPost]
        public async Task<HttpResponseMessage> UploadExcelFile()
        {
            if (!Request.Content.IsMimeMultipartContent())
            {
                return Request.CreateResponse(HttpStatusCode.BadRequest, "No file uploaded.");
            }

            var provider = new MultipartMemoryStreamProvider();

            IEnumerable<HttpContent> parts = null;
            Task.Factory
                .StartNew(() => parts = Request.Content.ReadAsMultipartAsync(provider).Result.Contents,
                    CancellationToken.None,
                    TaskCreationOptions.LongRunning, // guarantees separate thread
                    TaskScheduler.Default)
                .Wait();

            var file = provider.Contents.FirstOrDefault();
            if (file == null || file.Headers.ContentLength == 0)
            {
                return Request.CreateResponse(HttpStatusCode.BadRequest, "No file uploaded.");
            }

            try
            {
                var stream = await file.ReadAsStreamAsync();
                using (var workbook = new XLWorkbook(stream))
                {
                    var worksheet = workbook.Worksheet(1);
                    var dataTable = new DataTable();

                    var firstRow = worksheet.Row(1);
                    var allowedColumns = new HashSet<string> { 
                        "Plaqueta", 
                        "Descrição do Bem", 
                        "PREDIO", 
                        "Data de Aquis", 
                        "Vlr Original C", 
                        "Residual Cont", 
                        "Cod", 
                        "Descrição da Conta" 
                    };

                    // Add only allowed columns to the DataTable
                    foreach (var cell in firstRow.CellsUsed())
                    {
                        var columnName = cell.Value.ToString().Replace("º", string.Empty).Trim();
                        if (allowedColumns.Contains(columnName))
                        {
                            dataTable.Columns.Add(columnName, typeof(string));
                        }
                    }

                    // Populate the DataTable with only allowed columns
                    foreach (var row in worksheet.RowsUsed().Skip(1))
                    {
                        var dataRow = dataTable.NewRow();
                        foreach (var cell in row.CellsUsed())
                        {
                            var columnName = firstRow.Cell(cell.Address.ColumnNumber).Value.ToString().Replace("º", string.Empty).Trim();
                            if (allowedColumns.Contains(columnName))
                            {
                                dataRow[columnName] = cell.Value.ToString();
                            }
                        }
                        dataTable.Rows.Add(dataRow);
                    }

                    string connectionString = ConfigurationManager.AppSettings["DefaultConnection"];
                    string tableName = ConfigurationManager.AppSettings["TableName"];

                    if (string.IsNullOrWhiteSpace(connectionString) || string.IsNullOrWhiteSpace(tableName))
                    {
                        return Request.CreateResponse(HttpStatusCode.InternalServerError, "Invalid configuration settings.");
                    }

                    ImportToDatabase(dataTable, connectionString, tableName);
                }

                return Request.CreateResponse(HttpStatusCode.OK, "File processed directly from stream and data imported successfully.");
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.InternalServerError, $"Internal server error: {ex.Message}");
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
                    Debug.WriteLine("Database connection opened.");

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
                        var clearTableQuery = $"TRUNCATE TABLE {tableName}";
                        using (var clearCommand = new SqlCommand(clearTableQuery, connection))
                        {
                            clearCommand.ExecuteNonQuery();
                            Console.WriteLine($"Table {tableName} cleared.");
                            Debug.WriteLine($"Table {tableName} cleared.");
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
                            Debug.WriteLine($"Table {tableName} cleared.");
                        }
                    }

                    // Use SqlBulkCopy to insert data into the table
                    using (var bulkCopy = new SqlBulkCopy(connection))
                    {
                        bulkCopy.BulkCopyTimeout = 900;
                        bulkCopy.DestinationTableName = tableName;
                        foreach (DataColumn col in dataTable.Columns)
                        {
                            bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);
                        }
                        bulkCopy.WriteToServer(dataTable);
                        Console.WriteLine($"Data imported to table {tableName} successfully.");
                        Debug.WriteLine($"Data imported to table {tableName} successfully.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during data import: {ex.Message}");
                Debug.WriteLine($"Error during data import: {ex.Message}");
            }
        }

        private string GenerateCreateTableQuery(DataTable dataTable, string tableName)
        {
            var query = $"IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{tableName}') BEGIN CREATE TABLE {tableName} (";

            foreach (DataColumn column in dataTable.Columns)
            {
                var columnName = column.ColumnName.Replace("º", string.Empty).Trim();
                query += $"[{columnName}] NVARCHAR(300),";
                Console.WriteLine($"Generating column for SQL table: {columnName}");
            }

            query = query.TrimEnd(',') + "); END";
            return query;
        }
    }
}
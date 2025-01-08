using OfficeOpenXml;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.IO;
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
            // Set the license context for EPPlus (necessary for version 5+)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            try
            {
                // Read the file into a memory stream
                using (var stream = new MemoryStream())
                {
                    await file.CopyToAsync(stream);

                    // Use EPPlus to process the Excel file
                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets[0]; // Get the first worksheet
                        var columnCount = worksheet.Dimension.End.Column; // Get the number of columns

                        // Get the column headers
                        var columns = new string[columnCount];
                        for (int col = 1; col <= columnCount; col++)
                        {
                            columns[col - 1] = worksheet.Cells[1, col].Text; // Read the first row as headers
                        }

                        // Save the uploaded file to the specified directory in the correct format
                        var siteDirectory = @"C:\Your\Choice\Bro\"; // Path to your site directory
                        var filePath = Path.Combine(siteDirectory, "Excel2.xlsx");

                        // Ensure the directory exists
                        if (!Directory.Exists(siteDirectory))
                        {
                            Directory.CreateDirectory(siteDirectory);
                        }

                        // Save the file to disk in .xlsx format
                        using (var fileStream = new FileStream(filePath, FileMode.Create))
                        {
                            package.SaveAs(fileStream); // Save as .xlsx using EPPlus
                        }

                        return Ok("File uploaded and saved successfully.");
                    }
                }
            }
            catch (System.Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }
    }
}

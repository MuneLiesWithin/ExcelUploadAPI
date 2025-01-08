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
                var siteDirectory = @"C:\Users\Administrator\Desktop\Rômulo\Github\HMVPortalVerona\Excel"; // Path to your site directory
                var filePath = Path.Combine(siteDirectory, "Excel2.xlsx");

                // Ensure the directory exists
                if (!Directory.Exists(siteDirectory))
                {
                    Directory.CreateDirectory(siteDirectory);
                }

                // Save the raw file directly to disk
                //using (var stream = new FileStream(filePath, FileMode.Create))
                //{
                //    await file.CopyToAsync(stream);
                //}

                var stream = file.OpenReadStream();

                // Validate the uploaded file with EPPlus
                //using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))

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
                    //Importer
                }



                return Ok("File uploaded and saved successfully.");
            }
            catch (System.Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }
    }
}

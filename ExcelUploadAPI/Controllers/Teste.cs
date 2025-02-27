using Microsoft.AspNetCore.Mvc;

namespace ExcelUploadAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class Teste : Controller
    {
        private static readonly string[] Summaries = new[] 
        { "1","2","3","Testando"
        };
    }
}

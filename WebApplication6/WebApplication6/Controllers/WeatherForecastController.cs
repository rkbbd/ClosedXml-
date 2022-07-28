using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.Buffers.Text;

namespace WebApplication6.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
        "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
    };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet(Name = "DownloadExcel")]  //<a href="https://localhost:7135/WeatherForecast"></a>
        public ActionResult DownloadExcel()
        {
            var wbCorgiBabiesTemplate = new XLWorkbook();
            var wsCoriBabiesAmendementTemplate = wbCorgiBabiesTemplate.Worksheets.Add(" Work Sheet Corgi baby Template");
            wsCoriBabiesAmendementTemplate.Cell("A1").Value = "Corgi Parent";
            wsCoriBabiesAmendementTemplate.Cell("B1").Value = "Corgi Child";

            wsCoriBabiesAmendementTemplate.Cell("A2").Value = "Petunia";
            wsCoriBabiesAmendementTemplate.Cell("B2").Value = "Khaleesi";
            wbCorgiBabiesTemplate.SaveAs("new.xlsx");
            var ms = new MemoryStream();
           wbCorgiBabiesTemplate.SaveAs(ms);

            ms.Position = 0;
            var fileName = "CorgiBabies.xlsx";
            return File(ms, "application/octet-stream", fileName);
        }
    }
}
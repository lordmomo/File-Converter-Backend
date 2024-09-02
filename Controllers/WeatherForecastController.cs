using Microsoft.AspNetCore.Mvc;
using FileConversion.Entity;
using FileConversion.DocumentConversion;
namespace FileConversion.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : Controller
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

        [HttpGet(Name = "GetWeatherForecast")]
        public IEnumerable<WeatherForecast> Get()
        {

            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateOnly.FromDateTime(DateTime.Now.AddDays(index)),
                TemperatureC = Random.Shared.Next(-20, 55),
                Summary = Summaries[Random.Shared.Next(Summaries.Length)]
            })
            .ToArray();
        }

        [HttpPost]
        [Route("/createPdf")]
        public ActionResult Post()
        {
            ExampleStart.HelloWorld();
            Console.WriteLine("done");
            //return Ok("Success");
            return Json("Success");
        }


        [HttpPost]
        [Route("person")]
        public IActionResult PostPerson(Pesron person)
        {
            try
            {
                Console.WriteLine(person);
                ExampleStart.Greeting(person);
                return Ok("success");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error occurred while processing person data.");
                return StatusCode(500, new { message = "Internal server error occurred." });
            }
        }

        [HttpPost]
        [Route("compareDocs")]
        public IActionResult PostCompareDocument(CompareDocumentsDto compareDocumentsDto)
        {
            try
            {
                var output = DocumentComparer.CompareDocument(compareDocumentsDto);
                return Ok(output);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error occurred while processing person data.");
                return StatusCode(500, new { message = "Internal server error occurred." });
            }
        }

    }

}

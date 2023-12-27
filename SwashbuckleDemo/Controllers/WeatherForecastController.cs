using Microsoft.AspNetCore.Mvc;

namespace RestFullMinimalApi.Controllers;

[ApiController]
[Route("[controller]")]
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

    /// <summary>
    /// Retrieves WeatherForecast
    /// </summary>
    /// <remarks>Awesomeness!</remarks>
    /// <response code="200">Retrieved</response>
    /// <response code="404">Not found</response>
    /// <response code="500">Oops! Can't lookup your request right now</response>
    [HttpGet(Name = "GetWeatherForecast")]
    public IEnumerable<WeatherForecast> Get()
    {
        return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateTime.Now.AddDays(index),
                TemperatureC = Random.Shared.Next(-20, 55),
                Summary = Summaries[Random.Shared.Next(Summaries.Length)]
            })
            .ToArray();
    }

    /// <summary>
    /// Retrieves WeatherForecast as Pdf
    /// </summary>
    /// <remarks>Awesomeness!</remarks>
    /// <response code="200">Retrieved</response>
    /// <response code="404">Not found</response>
    /// <response code="500">Oops! Can't lookup your request right now</response>
    [HttpGet("download", Name = "DownloadWeatherForecast")]
    public IActionResult GetWeatherExcel()
    {
        var results = Enumerable.Range(1, 5).Select(index => new WeatherForecast
        {
            Date = DateTime.Now.AddDays(index),
            TemperatureC = Random.Shared.Next(-20, 55),
            Summary = Summaries[Random.Shared.Next(Summaries.Length)]
        }).ToArray();

        var html = GetHtml(results);
        var Renderer = new ChromePdfRenderer();
        var PDF = Renderer.RenderHtmlAsPdf(html);
        var fileName = "WeatherReport.pdf";
        PDF.SaveAs(fileName);

        var stream = new FileStream(fileName, FileMode.Open);

        // Save the excel file
        return new FileStreamResult(stream, "application/octet-stream") { FileDownloadName = fileName };
    }

    private static string GetHtml(WeatherForecast[] weatherForecasts)
    {
        string header = $@"
<html>
<head><title>WeatherForecast</title></head>
<body>
<h1>WeatherForecast</h1>
    ";
        var footer = @"
</body>
</html>";

        var htmlContent = header;

        foreach (var weather in weatherForecasts)
        {
            htmlContent += $@"
    <h2>{weather.Date}</h2>
    <p>Summary: {weather.Summary}</p>
    <p>Temperature in Celcius: {weather.TemperatureC}</p>
    <p>Temperature in Farenheit: {weather.TemperatureF}</p>
";
        }

        htmlContent += footer;

        return htmlContent;
    }
}
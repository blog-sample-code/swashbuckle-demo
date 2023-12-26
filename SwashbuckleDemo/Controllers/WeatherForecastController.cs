using IronXL;
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
    /// Retrieves WeatherForecast as Excel
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
        // Create new Excel WorkBook document.
        WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
        workBook.Metadata.Author = "IronXL";

        // Add a blank WorkSheet
        WorkSheet workSheet = workBook.CreateWorkSheet("main_sheet");

        // Add data and styles to the new worksheet
        workSheet["A1"].Value = "Date";
        workSheet["A1"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Thick;
        workSheet["B1"].Value = "TemperatureC";
        workSheet["B1"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Thick;
        workSheet["C1"].Value = "TemperatureF";
        workSheet["C1"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Thick;
        workSheet["D1"].Value = "Summary";
        workSheet["D1"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Thick;


        var row = 1;
        foreach (var item in results)
        {
            row++;
            workSheet[$"A{row}"].Value = item.Date.ToShortDateString();
            workSheet[$"B{row}"].Value = item.TemperatureC;
            workSheet[$"C{row}"].Value = item.TemperatureF;
            workSheet[$"D{row}"].Value = item.Summary;

        }


        // Save the excel file
        return new FileStreamResult(workBook.ToStream(), "application/octet-stream") { FileDownloadName = "weather.xlsx" };
    }
}
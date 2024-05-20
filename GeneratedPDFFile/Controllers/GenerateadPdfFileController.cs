using GeneratedPDFFile.Services.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace GeneratedPDFFile.Controllers;

[ApiController]
[Route("api/[controller]")]
public class GenerateadPdfFileController : ControllerBase
{
    private readonly IGenerateadPdfFileService _generateadPdfFileService;

    public GenerateadPdfFileController(IGenerateadPdfFileService generateadPdfFileService)
    {
        _generateadPdfFileService = generateadPdfFileService;
    }



    [HttpPost("generateadPdfFile")]
    public async ValueTask<IActionResult> GenerateadPdf(string htmlToString)
    {
        await this._generateadPdfFileService.InsertHtmlToWord(htmlToString);
        return Ok();
    }



    [HttpPost("htmlcode")]
    public async ValueTask<IActionResult> GenerateadFile(string html)
    {
         this._generateadPdfFileService.ConvertHtmlToDocx(html);
        return Ok();
    }
}
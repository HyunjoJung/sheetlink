using Microsoft.AspNetCore.Mvc;
using ExcelLinkExtractorWeb.Services;

namespace ExcelLinkExtractorWeb.Controllers;

[ApiController]
[Route("api/[controller]")]
public class FileController : ControllerBase
{
    private readonly LinkExtractorService _extractorService;

    public FileController(LinkExtractorService extractorService)
    {
        _extractorService = extractorService;
    }

    [HttpPost("extract")]
    public async Task<IActionResult> ExtractLinks(IFormFile file, [FromForm] string columnName = "Title")
    {
        if (file == null || file.Length == 0)
            return BadRequest(new { error = "파일을 선택해주세요." });

        if (file.Length > 10 * 1024 * 1024)
            return BadRequest(new { error = "파일 크기는 10MB 이하여야 합니다." });

        var extension = Path.GetExtension(file.FileName).ToLower();
        if (extension != ".xlsx" && extension != ".xls")
            return BadRequest(new { error = "xlsx 또는 xls 파일만 지원합니다." });

        using var stream = new MemoryStream();
        await file.CopyToAsync(stream);
        stream.Position = 0;

        var result = await _extractorService.ExtractLinksAsync(stream, columnName);

        if (!string.IsNullOrEmpty(result.ErrorMessage))
            return BadRequest(new { error = result.ErrorMessage });

        return Ok(new
        {
            totalRows = result.TotalRows,
            linksFound = result.LinksFound,
            links = result.Links.Take(10).Select(l => new { l.Row, l.Title, l.Url }),
            outputFileBase64 = Convert.ToBase64String(result.OutputFile!)
        });
    }

    [HttpGet("template")]
    public IActionResult DownloadTemplate()
    {
        var bytes = _extractorService.CreateTemplate();
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "link_extract_template.xlsx");
    }

    [HttpPost("merge")]
    public IActionResult MergeLinks([FromBody] MergeRequest request)
    {
        if (request.Titles == null || request.Urls == null)
            return BadRequest(new { error = "Titles and URLs are required." });

        if (request.Titles.Count != request.Urls.Count)
            return BadRequest(new { error = "Title and URL counts must match." });

        var bytes = _extractorService.CreateMergedFile(request.Titles, request.Urls);
        var base64 = Convert.ToBase64String(bytes);
        return Ok(new { base64 });
    }

    public class MergeRequest
    {
        public List<string> Titles { get; set; } = new();
        public List<string> Urls { get; set; } = new();
    }
}

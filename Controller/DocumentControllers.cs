using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

[ApiController]
[Route("api/[controller]")]
public class DocumentController : ControllerBase
{
    private readonly IWebHostEnvironment _env;

    public DocumentController(IWebHostEnvironment env)
    {
        _env = env;
    }

    [HttpPost("generate-pdf")]
    public async Task<IActionResult> GeneratePdf([FromBody] DocumentRequest request)
    {
        try
        {
            string templatePath = Path.Combine(_env.ContentRootPath, "templates", request.TemplateName);
            string outputPath = Path.Combine(_env.ContentRootPath, "output", $"{Guid.NewGuid()}.docx");
            string pdfPath = Path.ChangeExtension(outputPath, ".pdf");

            if (!System.IO.File.Exists(templatePath))
                return BadRequest($"Template not found: {request.TemplateName}");

            // Copy template to working location
            System.IO.File.Copy(templatePath, outputPath, true);

            // Create service and replace variables
            var service = new DocumentTemplateService(outputPath, outputPath);
            service.ReplaceTemplateVariables(request.Variables);

            // Convert to PDF
            service.ConvertToPdf(pdfPath);

            // Return PDF
            var fileBytes = await System.IO.File.ReadAllBytesAsync(pdfPath);
            return File(fileBytes, "application/pdf", $"{Path.GetFileNameWithoutExtension(request.TemplateName)}.pdf");
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { error = ex.Message });
        }
    }
}

public class DocumentRequest
{
    public string TemplateName { get; set; }
    public Dictionary<string, string> Variables { get; set; }
}

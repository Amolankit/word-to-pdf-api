using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;
using System.IO;

public class DocumentTemplateService
{
    private readonly string _templatePath;
    private readonly string _outputPath;

    public DocumentTemplateService(string templatePath, string outputPath)
    {
        _templatePath = templatePath;
        _outputPath = outputPath;
    }

    /// <summary>
    /// Replace placeholders in Word document with values
    /// </summary>
    public void ReplaceTemplateVariables(Dictionary<string, string> variables)
    {
        using (WordprocessingDocument doc = WordprocessingDocument.Open(_templatePath, true))
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;
            
            // Replace in main document body
            ReplaceTextInDocument(mainPart.Document.Body, variables);
            
            // Replace in headers
            foreach (var headerPart in mainPart.HeaderParts)
            {
                ReplaceTextInDocument(headerPart.Header, variables);
            }
            
            // Replace in footers
            foreach (var footerPart in mainPart.FooterParts)
            {
                ReplaceTextInDocument(footerPart.Footer, variables);
            }

            doc.Save();
        }
    }

    /// <summary>
    /// Replace images in the document
    /// </summary>
    public void ReplaceImage(string placeholder, string imagePath)
    {
        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image not found: {imagePath}");

        using (WordprocessingDocument doc = WordprocessingDocument.Open(_templatePath, true))
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;
            
            // Find and replace drawing elements
            var drawings = mainPart.Document.Body.Descendants<Drawing>().ToList();
            
            // Logic to replace specific drawing by placeholder
            // This is complex; consider using bookmarks instead
            
            doc.Save();
        }
    }

    /// <summary>
    /// Convert Word document to PDF using LibreOffice
    /// </summary>
    public void ConvertToPdf(string outputPdfPath)
    {
        try
        {
            // Using LibreOffice command-line conversion
            var processInfo = new ProcessStartInfo
            {
                FileName = GetLibreOfficePath(),
                Arguments = $"--headless --convert-to pdf --outdir \"{Path.GetDirectoryName(outputPdfPath)}\" \"{_templatePath}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };

            using (var process = Process.Start(processInfo))
            {
                process.WaitForExit(30000); // 30 second timeout
                
                if (process.ExitCode != 0)
                    throw new Exception($"LibreOffice conversion failed with exit code: {process.ExitCode}");
            }

            // Move the generated PDF to the desired location
            string generatedPdf = Path.Combine(
                Path.GetDirectoryName(outputPdfPath),
                Path.GetFileNameWithoutExtension(_templatePath) + ".pdf"
            );
            
            if (File.Exists(generatedPdf))
            {
                File.Move(generatedPdf, outputPdfPath, true);
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"PDF conversion failed: {ex.Message}", ex);
        }
    }

    private void ReplaceTextInDocument(OpenXmlElement element, Dictionary<string, string> variables)
    {
        foreach (var paragraph in element.Descendants<Paragraph>())
        {
            foreach (var run in paragraph.Descendants<Run>())
            {
                foreach (var text in run.Descendants<Text>())
                {
                    foreach (var kvp in variables)
                    {
                        if (text.Text.Contains(kvp.Key))
                        {
                            text.Text = text.Text.Replace(kvp.Key, kvp.Value);
                        }
                    }
                }
            }
        }
    }

    private string GetLibreOfficePath()
    {
        // Windows paths
        var possiblePaths = new[]
        {
            @"C:\Program Files\LibreOffice\program\soffice.exe",
            @"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            "/usr/bin/libreoffice", // Linux
            "/Applications/LibreOffice.app/Contents/MacOS/libreoffice" // macOS
        };

        foreach (var path in possiblePaths)
        {
            if (File.Exists(path))
                return path;
        }

        throw new Exception("LibreOffice not found. Please install LibreOffice.");
    }
}

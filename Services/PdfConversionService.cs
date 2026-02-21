using System.Diagnostics;

namespace WordToPdfApi.Services;

public class PdfConversionService : IPdfConversionService
{
    private readonly ILogger<PdfConversionService> _logger;

    public PdfConversionService(ILogger<PdfConversionService> logger)
    {
        _logger = logger;
    }

    public async Task ConvertToPdfAsync(string inputDocPath, string outputPdfPath)
    {
        try
        {
            if (!File.Exists(inputDocPath))
                throw new FileNotFoundException($"Input document not found: {inputDocPath}");

            string libreOfficePath = GetLibreOfficePath();

            var processInfo = new ProcessStartInfo
            {
                FileName = libreOfficePath,
                Arguments = $"--headless --convert-to pdf --outdir \"{Path.GetDirectoryName(outputPdfPath)}\" \"{inputDocPath}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using (var process = Process.Start(processInfo))
            {
                if (process == null)
                    throw new Exception("Failed to start LibreOffice process");

                bool exited = process.WaitForExit(60000); // 60 second timeout

                if (!exited)
                {
                    process.Kill();
                    throw new Exception("LibreOffice conversion timed out");
                }

                if (process.ExitCode != 0)
                {
                    string error = process.StandardError.ReadToEnd();
                    throw new Exception($"LibreOffice conversion failed with exit code {process.ExitCode}: {error}");
                }

                _logger.LogInformation("LibreOffice conversion completed successfully");
            }

            // Find and move the generated PDF
            string outputDir = Path.GetDirectoryName(outputPdfPath) ?? Environment.CurrentDirectory;
            string generatedPdfName = Path.GetFileNameWithoutExtension(inputDocPath) + ".pdf";
            string generatedPdfPath = Path.Combine(outputDir, generatedPdfName);

            if (File.Exists(generatedPdfPath))
            {
                // Wait a bit for file to be completely written
                await Task.Delay(500);
                File.Move(generatedPdfPath, outputPdfPath, true);
                _logger.LogInformation($"PDF moved to final location: {outputPdfPath}");
            }
            else
            {
                throw new Exception($"Generated PDF not found at: {generatedPdfPath}");
            }

            await Task.CompletedTask;
        }
        catch (Exception ex)
        {
            _logger.LogError($"PDF conversion error: {ex.Message}");
            throw;
        }
    }

    public async Task<bool> IsLibreOfficeAvailableAsync()
    {
        try
        {
            string path = GetLibreOfficePath();
            return await Task.FromResult(File.Exists(path));
        }
        catch
        {
            return await Task.FromResult(false);
        }
    }

    public string GetLibreOfficePath()
    {
        var possiblePaths = new List<string>();

        // Windows paths
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            possiblePaths.AddRange(new[]
            {
                @"C:\Program Files\LibreOffice\program\soffice.exe",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.com"
            });
        }
        // Linux paths
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            possiblePaths.AddRange(new[]
            {
                "/usr/bin/libreoffice",
                "/usr/bin/soffice",
                "/snap/bin/libreoffice"
            });
        }
        // macOS paths
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            possiblePaths.AddRange(new[]
            {
                "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                "/opt/homebrew/bin/libreoffice"
            });
        }

        foreach (var path in possiblePaths)
        {
            if (File.Exists(path))
            {
                _logger.LogInformation($"Found LibreOffice at: {path}");
                return path;
            }
        }

        throw new FileNotFoundException("LibreOffice not found. Please install LibreOffice from https://www.libreoffice.org/");
    }
}

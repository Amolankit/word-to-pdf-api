using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO.Packaging;

namespace WordToPdfApi.Services;

public class DocumentService : IDocumentService
{
    private readonly ILogger<DocumentService> _logger;

    public DocumentService(ILogger<DocumentService> logger)
    {
        _logger = logger;
    }

    public async Task ReplaceTemplateVariablesAsync(string documentPath, Dictionary<string, string> variables)
    {
        try
        {
            if (!File.Exists(documentPath))
                throw new FileNotFoundException($"Document not found: {documentPath}");

            using (WordprocessingDocument doc = WordprocessingDocument.Open(documentPath, true))
            {
                MainDocumentPart mainPart = doc.MainDocumentPart ?? 
                    throw new InvalidOperationException("Document does not contain main document part");

                // Replace in main document body
                ReplaceTextInElement(mainPart.Document.Body, variables);

                // Replace in headers
                foreach (var headerPart in mainPart.HeaderParts)
                {
                    ReplaceTextInElement(headerPart.Header, variables);
                }

                // Replace in footers
                foreach (var footerPart in mainPart.FooterParts)
                {
                    ReplaceTextInElement(footerPart.Footer, variables);
                }

                doc.Save();
                _logger.LogInformation($"Successfully replaced variables in document: {documentPath}");
            }

            await Task.CompletedTask;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error replacing template variables: {ex.Message}");
            throw;
        }
    }

    public async Task ReplaceBookmarkTextAsync(string documentPath, Dictionary<string, string> bookmarkValues)
    {
        try
        {
            if (!File.Exists(documentPath))
                throw new FileNotFoundException($"Document not found: {documentPath}");

            using (WordprocessingDocument doc = WordprocessingDocument.Open(documentPath, true))
            {
                MainDocumentPart mainPart = doc.MainDocumentPart ?? 
                    throw new InvalidOperationException("Document does not contain main document part");

                var bookmarks = mainPart.Document.Body.Descendants<BookmarkStart>().ToList();

                foreach (var bookmark in bookmarks)
                {
                    string bookmarkName = bookmark.Name?.Value ?? string.Empty;

                    if (bookmarkValues.TryGetValue(bookmarkName, out string? replacementText))
                    {
                        var bookmarkEnd = mainPart.Document.Body.Descendants<BookmarkEnd>()
                            .FirstOrDefault(be => be.Id?.Value == bookmark.Id?.Value);

                        if (bookmarkEnd != null)
                        {
                            // Get all nodes between bookmark start and end
                            var parent = bookmark.Parent;
                            var nodesBetween = GetNodesBetween(bookmark, bookmarkEnd);

                            // Remove old content
                            foreach (var node in nodesBetween)
                            {
                                node.Remove();
                            }

                            // Insert new text run
                            var newRun = new Run(new Text(replacementText ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve });
                            bookmark.Parent?.InsertAfter(newRun, bookmark);

                            _logger.LogInformation($"Replaced bookmark: {bookmarkName}");
                        }
                    }
                }

                doc.Save();
                _logger.LogInformation($"Successfully replaced bookmarks in document: {documentPath}");
            }

            await Task.CompletedTask;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error replacing bookmark text: {ex.Message}");
            throw;
        }
    }

    public async Task ReplaceImageAsync(string documentPath, string bookmarkName, string imagePath)
    {
        try
        {
            if (!File.Exists(documentPath))
                throw new FileNotFoundException($"Document not found: {documentPath}");

            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image not found: {imagePath}");

            using (WordprocessingDocument doc = WordprocessingDocument.Open(documentPath, true))
            {
                MainDocumentPart mainPart = doc.MainDocumentPart ?? 
                    throw new InvalidOperationException("Document does not contain main document part");

                // Find bookmark
                var bookmarkStart = mainPart.Document.Body.Descendants<BookmarkStart>()
                    .FirstOrDefault(b => b.Name?.Value == bookmarkName);

                if (bookmarkStart != null)
                {
                    var bookmarkEnd = mainPart.Document.Body.Descendants<BookmarkEnd>()
                        .FirstOrDefault(be => be.Id?.Value == bookmarkStart.Id?.Value);

                    if (bookmarkEnd != null)
                    {
                        // Remove content between bookmarks
                        var nodesBetween = GetNodesBetween(bookmarkStart, bookmarkEnd);
                        foreach (var node in nodesBetween)
                        {
                            node.Remove();
                        }

                        // Add image
                        var imagePart = mainPart.AddImagePart(GetImagePartType(imagePath));
                        using (FileStream fs = new FileStream(imagePath, FileMode.Open))
                        {
                            imagePart.FeedData(fs);
                        }

                        var imageElement = CreateImageElement(mainPart, imagePart, imagePath);
                        var paragraph = new Paragraph(new ParagraphProperties(), new Run(imageElement));
                        bookmarkStart.Parent?.InsertAfter(paragraph, bookmarkStart);

                        _logger.LogInformation($"Replaced image in bookmark: {bookmarkName}");
                    }
                }

                doc.Save();
                _logger.LogInformation($"Successfully replaced image in document: {documentPath}");
            }

            await Task.CompletedTask;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error replacing image: {ex.Message}");
            throw;
        }
    }

    public async Task<List<string>> GetBookmarksAsync(string documentPath)
    {
        try
        {
            if (!File.Exists(documentPath))
                throw new FileNotFoundException($"Document not found: {documentPath}");

            var bookmarks = new List<string>();

            using (WordprocessingDocument doc = WordprocessingDocument.Open(documentPath, false))
            {
                MainDocumentPart mainPart = doc.MainDocumentPart ?? 
                    throw new InvalidOperationException("Document does not contain main document part");

                var bookmarkStarts = mainPart.Document.Body.Descendants<BookmarkStart>();
                foreach (var bookmark in bookmarkStarts)
                {
                    if (!string.IsNullOrEmpty(bookmark.Name?.Value))
                    {
                        bookmarks.Add(bookmark.Name.Value);
                    }
                }
            }

            return await Task.FromResult(bookmarks);
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error getting bookmarks: {ex.Message}");
            throw;
        }
    }

    private void ReplaceTextInElement(OpenXmlElement element, Dictionary<string, string> variables)
    {
        var paragraphs = element.Descendants<Paragraph>().ToList();

        foreach (var paragraph in paragraphs)
        {
            var runs = paragraph.Descendants<Run>().ToList();

            foreach (var run in runs)
            {
                var texts = run.Descendants<Text>().ToList();

                foreach (var text in texts)
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

    private List<OpenXmlElement> GetNodesBetween(OpenXmlElement start, OpenXmlElement end)
    {
        var nodes = new List<OpenXmlElement>();
        var current = start.NextSibling();

        while (current != null && current != end)
        {
            var next = current.NextSibling();
            nodes.Add(current);
            current = next;
        }

        return nodes;
    }

    private ImagePartType GetImagePartType(string imagePath)
    {
        string extension = Path.GetExtension(imagePath).ToLower();
        return extension switch
        {
            ".jpg" or ".jpeg" => ImagePartType.Jpeg,
            ".png" => ImagePartType.Png,
            ".gif" => ImagePartType.Gif,
            ".bmp" => ImagePartType.Bmp,
            _ => throw new NotSupportedException($"Image format not supported: {extension}")
        };
    }

    private Drawing CreateImageElement(MainDocumentPart mainPart, ImagePart imagePart, string imagePath)
    {
        var imageId = mainPart.GetIdOfPart(imagePart);
        var fileInfo = new FileInfo(imagePath);
        
        // Default dimensions in EMU (English Metric Units)
        long width = 2000000;  // 2 inches
        long height = 2000000; // 2 inches

        var drawing = new Drawing(
            new Wp.Anchor(
                new Wp.SimplePosition { X = 0, Y = 0 },
                new Wp.PositionH { RelativeHeight = 251658240U, AlignH = new Wp.AlignH { Text = "center" } },
                new Wp.PositionV { RelativeHeight = 251658240U, AlignV = new Wp.AlignV { Text = "center" } },
                new Wp.Extent { Cx = width, Cy = height },
                new Wp.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new Wp.WrapNone(),
                new Wp.DocProperties { Id = 1U, Name = "Image 1" },
                new A.GraphicData(
                    new Pic.Picture(
                        new Pic.NonVisualPictureProperties(
                            new Pic.NonVisualDrawingProperties { Id = 0U, Name = "Image 1" },
                            new Pic.NonVisualPictureDrawingProperties()),
                        new Pic.BlipFill(
                            new A.Blip { Embed = imageId },
                            new Pic.Stretch(new A.FillRect())),
                        new Pic.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0, Y = 0 },
                                new A.Extents { Cx = width, Cy = height }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Prst = "rect" })))
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 114300U, DistanceFromRight = 114300U, SimplePos = false, RelativeHeight = 251658240U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true });

        return drawing;
    }
}

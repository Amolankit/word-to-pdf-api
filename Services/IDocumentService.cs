namespace WordToPdfApi.Services;

public interface IDocumentService
{
    /// <summary>
    /// Replace template variables in a Word document using placeholders
    /// </summary>
    Task ReplaceTemplateVariablesAsync(string documentPath, Dictionary<string, string> variables);

    /// <summary>
    /// Replace text in bookmarks
    /// </summary>
    Task ReplaceBookmarkTextAsync(string documentPath, Dictionary<string, string> bookmarkValues);

    /// <summary>
    /// Replace images in the document by bookmark
    /// </summary>
    Task ReplaceImageAsync(string documentPath, string bookmarkName, string imagePath);

    /// <summary>
    /// Get all bookmarks from a document
    /// </summary>
    Task<List<string>> GetBookmarksAsync(string documentPath);
}
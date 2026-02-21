using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class BookmarkReplacer
{
    public static void ReplaceBookmarkText(string documentPath, Dictionary<string, string> bookmarkValues)
    {
        using (WordprocessingDocument doc = WordprocessingDocument.Open(documentPath, true))
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;
            var bookmarks = mainPart.Document.Body.Descendants<BookmarkStart>().ToList();

            foreach (var bookmark in bookmarks)
            {
                string bookmarkName = bookmark.Name.Value;

                if (bookmarkValues.TryGetValue(bookmarkName, out string replacementText))
                {
                    // Find the bookmark end
                    var bookmarkEnd = mainPart.Document.Body.Descendants<BookmarkEnd>()
                        .FirstOrDefault(be => be.Id.Value == bookmark.Id.Value);

                    if (bookmarkEnd != null)
                    {
                        // Remove old content and insert new
                        var nodesBetween = GetNodesBetween(bookmark, bookmarkEnd);
                        foreach (var node in nodesBetween)
                        {
                            node.Remove();
                        }

                        // Insert new text run
                        bookmark.Parent.InsertAfter(new Run(new Text(replacementText)), bookmark);
                    }
                }
            }

            doc.Save();
        }
    }

    private static List<OpenXmlElement> GetNodesBetween(OpenXmlElement start, OpenXmlElement end)
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
}

namespace WordToPdfApi.Models;

public class GeneratePdfRequest
{
    /// <summary>
    /// Template file name (should be in /templates folder)
    /// </summary>
    public string TemplateName { get; set; } = string.Empty;

    /// <summary>
    /// Dictionary of variables to replace in the template
    /// </summary>
    public Dictionary<string, string> Variables { get; set; } = new();
}

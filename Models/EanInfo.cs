namespace LauraAssetBuildReview.Models;

/// <summary>
/// Information about an EAN found in a PowerPoint slide, including its text context.
/// </summary>
public class EanInfo
{
    /// <summary>
    /// The EAN value (normalized).
    /// </summary>
    public string Ean { get; set; } = string.Empty;

    /// <summary>
    /// The text context where this EAN was found (e.g., the cell text or surrounding text).
    /// </summary>
    public string TextContext { get; set; } = string.Empty;
}

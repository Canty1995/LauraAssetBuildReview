namespace LauraAssetBuildReview.Models;

/// <summary>
/// Result of reading EANs from a PowerPoint file, including per-slide statistics.
/// </summary>
public class PowerPointReadResult
{
    /// <summary>
    /// Set of all unique EANs found across all slides.
    /// </summary>
    public HashSet<string> Eans { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Dictionary mapping slide number (1-based) to the count of EANs found on that slide.
    /// </summary>
    public Dictionary<int, int> EanCountsPerSlide { get; set; } = new();

    /// <summary>
    /// Dictionary mapping slide number (1-based) to the list of EANs found on that slide with their text context.
    /// </summary>
    public Dictionary<int, List<EanInfo>> EansPerSlide { get; set; } = new();
}

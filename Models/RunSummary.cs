namespace LauraAssetBuildReview.Models;

public class RunSummary
{
    public int TotalEansProcessed { get; set; }
    public int MatchesInReferenceA { get; set; }
    public int MatchesInReferenceB { get; set; }
    public int NoMatches { get; set; }
    public TimeSpan ProcessingDuration { get; set; }
    public List<string> ValidationErrors { get; set; } = new();
    public List<string> Warnings { get; set; } = new();
}

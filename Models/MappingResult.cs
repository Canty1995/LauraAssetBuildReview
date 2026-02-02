namespace LauraAssetBuildReview.Models;

public class MappingResult
{
    public string? ReferenceAOption { get; set; }
    public string? ReferenceBOption { get; set; }
    public string? NoMatchOption { get; set; }
    public bool IsValid { get; set; }
    public List<string> ValidationErrors { get; set; } = new();
}

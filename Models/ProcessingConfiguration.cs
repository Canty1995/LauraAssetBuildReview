using System.Collections.Generic;

namespace LauraAssetBuildReview.Models;

/// <summary>
/// Configuration model for flexible processing options.
/// Replaces hardcoded values with user-configurable settings.
/// </summary>
public class ProcessingConfiguration
{
    // Main file settings
    public string EanColumn { get; set; } = "C";
    public string StatusColumn { get; set; } = "G";
    public string DropdownColumn { get; set; } = "G"; // Column where dropdown validation is located
    public int StartRow { get; set; } = 3;
    public int WorksheetIndex { get; set; } = 1; // 1-based index
    public string? WorksheetName { get; set; } // If specified, use name instead of index

    // EAN validation settings
    public int MinEanDigits { get; set; } = 14;
    public int MaxEanDigits { get; set; } = 14;
    public bool AllowNonNumericEans { get; set; } = false;

    // Reference file settings
    public List<ReferenceFileConfig> ReferenceFiles { get; set; } = new();

    // Dropdown/mapping settings
    public int? ExpectedDropdownCount { get; set; } = null; // null = any count
    public Dictionary<string, string> ManualDropdownMappings { get; set; } = new(); // File path -> dropdown option
    public bool AutoMapFilenames { get; set; } = true; // If false, use manual mappings only

    // Comparison settings
    public string ComparisonColumn { get; set; } = "G";
    public int ComparisonStartRow { get; set; } = 3;

    // Helper methods
    public int GetEanColumnNumber() => ColumnLetterToNumber(EanColumn);
    public int GetStatusColumnNumber() => ColumnLetterToNumber(StatusColumn);
    public int GetDropdownColumnNumber() => ColumnLetterToNumber(DropdownColumn);
    public int GetComparisonColumnNumber() => ColumnLetterToNumber(ComparisonColumn);

    public static int ColumnLetterToNumber(string columnLetter)
    {
        if (string.IsNullOrWhiteSpace(columnLetter))
            return 1;

        columnLetter = columnLetter.ToUpperInvariant().Trim();
        int result = 0;
        foreach (char c in columnLetter)
        {
            result = result * 26 + (c - 'A' + 1);
        }
        return result;
    }

    public static string ColumnNumberToLetter(int columnNumber)
    {
        string result = string.Empty;
        while (columnNumber > 0)
        {
            columnNumber--;
            result = (char)('A' + columnNumber % 26) + result;
            columnNumber /= 26;
        }
        return result;
    }
}

public class ReferenceFileConfig
{
    public string FilePath { get; set; } = string.Empty;
    public string FileType { get; set; } = "Excel"; // "Excel" or "PowerPoint"
    public string EanColumn { get; set; } = "C"; // For Excel files
    public int StartRow { get; set; } = 1; // For Excel files
    public int WorksheetIndex { get; set; } = 1; // For Excel files
    public string? WorksheetName { get; set; } // For Excel files
    public List<int> SelectedSlides { get; set; } = new(); // For PowerPoint files (1-based slide indices)
    public string? MappedDropdownOption { get; set; } // Manual mapping
    public int Priority { get; set; } = 0; // Lower number = higher priority
}

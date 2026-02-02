using System.IO;
using System.Text.RegularExpressions;
using LauraAssetBuildReview.Models;

namespace LauraAssetBuildReview.Services;

public class MatchingService
{
    /// <summary>
    /// Maps reference filenames to dropdown options using case-insensitive matching.
    /// Extracts key identifiers (like KING codes) from filenames and matches them to dropdown options.
    /// </summary>
    [Obsolete("Use MapFilenamesToDropdownsFlexible for flexible file count")]
    public MappingResult MapFilenamesToDropdowns(string referenceAPath, string referenceBPath, List<string> dropdownOptions)
    {
        var result = new MappingResult();

        // Validate dropdown has exactly 3 options
        if (dropdownOptions.Count != 3)
        {
            result.IsValid = false;
            result.ValidationErrors.Add($"Expected exactly 3 dropdown options, but found {dropdownOptions.Count}.");
            return result;
        }

        var refAName = Path.GetFileNameWithoutExtension(referenceAPath);
        var refBName = Path.GetFileNameWithoutExtension(referenceBPath);

        // Try to match filenames to dropdown options
        string? matchedA = null;
        string? matchedB = null;

        // Extract key identifiers from filenames (e.g., "KING01058", "KING01042")
        var refAIdentifiers = ExtractIdentifiers(refAName);
        var refBIdentifiers = ExtractIdentifiers(refBName);

        // First pass: Try exact case-insensitive matches
        foreach (var option in dropdownOptions)
        {
            var optionLower = option.ToLowerInvariant();
            var refALower = refAName.ToLowerInvariant();
            var refBLower = refBName.ToLowerInvariant();

            if (optionLower == refALower && matchedA == null)
            {
                matchedA = option;
            }
            if (optionLower == refBLower && matchedB == null)
            {
                matchedB = option;
            }
        }

        // Second pass: Try containment matches (if exact match not found)
        if (matchedA == null)
        {
            foreach (var option in dropdownOptions)
            {
                var optionLower = option.ToLowerInvariant();
                var refALower = refAName.ToLowerInvariant();

                if ((optionLower.Contains(refALower) || refALower.Contains(optionLower)) && option != matchedB)
                {
                    matchedA = option;
                    break;
                }
            }
        }

        if (matchedB == null)
        {
            foreach (var option in dropdownOptions)
            {
                var optionLower = option.ToLowerInvariant();
                var refBLower = refBName.ToLowerInvariant();

                if ((optionLower.Contains(refBLower) || refBLower.Contains(optionLower)) && option != matchedA)
                {
                    matchedB = option;
                    break;
                }
            }
        }

        // Third pass: Match based on extracted identifiers (e.g., KING codes)
        if (matchedA == null && refAIdentifiers.Count > 0)
        {
            foreach (var option in dropdownOptions)
            {
                if (option == matchedB) continue;
                
                var optionIdentifiers = ExtractIdentifiers(option);
                // Check if any identifier from filename appears in the option
                if (refAIdentifiers.Any(id => optionIdentifiers.Contains(id, StringComparer.OrdinalIgnoreCase)))
                {
                    matchedA = option;
                    break;
                }
            }
        }

        if (matchedB == null && refBIdentifiers.Count > 0)
        {
            foreach (var option in dropdownOptions)
            {
                if (option == matchedA) continue;
                
                var optionIdentifiers = ExtractIdentifiers(option);
                // Check if any identifier from filename appears in the option
                if (refBIdentifiers.Any(id => optionIdentifiers.Contains(id, StringComparer.OrdinalIgnoreCase)))
                {
                    matchedB = option;
                    break;
                }
            }
        }

        // Validate that both filenames were matched
        if (matchedA == null || matchedB == null)
        {
            result.IsValid = false;
            if (matchedA == null)
            {
                result.ValidationErrors.Add($"Could not match reference file A filename '{refAName}' to any dropdown option.");
            }
            if (matchedB == null)
            {
                result.ValidationErrors.Add($"Could not match reference file B filename '{refBName}' to any dropdown option.");
            }
            return result;
        }

        // Find the remaining option (no-match option)
        var noMatchOption = dropdownOptions.FirstOrDefault(o => o != matchedA && o != matchedB);
        if (noMatchOption == null)
        {
            result.IsValid = false;
            result.ValidationErrors.Add("Could not identify the no-match dropdown option.");
            return result;
        }

        result.ReferenceAOption = matchedA;
        result.ReferenceBOption = matchedB;
        result.NoMatchOption = noMatchOption;
        result.IsValid = true;

        return result;
    }

    /// <summary>
    /// Maps multiple reference filenames to dropdown options flexibly.
    /// Supports any number of reference files and dropdown options.
    /// </summary>
    public Dictionary<string, string> MapFilenamesToDropdownsFlexible(
        List<string> referenceFilePaths,
        List<string> dropdownOptions,
        Dictionary<string, string>? manualMappings = null)
    {
        var mappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        
        // Start with manual mappings if provided
        if (manualMappings != null)
        {
            foreach (var mapping in manualMappings)
            {
                mappings[mapping.Key] = mapping.Value;
            }
        }

        // For each reference file, try to auto-map if not manually mapped
        foreach (var filePath in referenceFilePaths)
        {
            var fileName = Path.GetFileNameWithoutExtension(filePath);
            
            // Skip if already manually mapped
            if (mappings.ContainsKey(filePath) || mappings.ContainsKey(fileName))
                continue;

            // Try to find a matching dropdown option
            var matchedOption = FindMatchingDropdownOption(fileName, dropdownOptions, mappings.Values);
            if (matchedOption != null)
            {
                mappings[filePath] = matchedOption;
            }
        }

        return mappings;
    }

    /// <summary>
    /// Finds a dropdown option that matches the given filename.
    /// </summary>
    private string? FindMatchingDropdownOption(string fileName, List<string> dropdownOptions, ICollection<string> alreadyMapped)
    {
        var fileNameLower = fileName.ToLowerInvariant();
        var fileNameIdentifiers = ExtractIdentifiers(fileName);

        // First pass: Exact match
        foreach (var option in dropdownOptions)
        {
            if (alreadyMapped.Contains(option, StringComparer.OrdinalIgnoreCase))
                continue;

            if (option.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                return option;
        }

        // Second pass: Containment match
        foreach (var option in dropdownOptions)
        {
            if (alreadyMapped.Contains(option, StringComparer.OrdinalIgnoreCase))
                continue;

            var optionLower = option.ToLowerInvariant();
            if (optionLower.Contains(fileNameLower) || fileNameLower.Contains(optionLower))
                return option;
        }

        // Third pass: Identifier match
        foreach (var option in dropdownOptions)
        {
            if (alreadyMapped.Contains(option, StringComparer.OrdinalIgnoreCase))
                continue;

            var optionIdentifiers = ExtractIdentifiers(option);
            if (fileNameIdentifiers.Any(id => optionIdentifiers.Contains(id, StringComparer.OrdinalIgnoreCase)))
                return option;
        }

        return null;
    }

    /// <summary>
    /// Extracts key identifiers from a string (e.g., "KING01058", "KING01042").
    /// Looks for patterns like uppercase letters followed by numbers, or common code patterns.
    /// </summary>
    private HashSet<string> ExtractIdentifiers(string text)
    {
        var identifiers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        
        if (string.IsNullOrWhiteSpace(text))
            return identifiers;

        // Pattern 1: Letters followed by numbers (e.g., "KING01058", "KING01042")
        // Matches: 1-10 letters (case-insensitive), followed by 1-10 digits
        var pattern1 = @"[A-Za-z]{1,10}\d{1,10}";
        var matches1 = Regex.Matches(text, pattern1, RegexOptions.IgnoreCase);
        foreach (Match match in matches1)
        {
            identifiers.Add(match.Value);
        }

        // Pattern 2: Any alphanumeric sequence that looks like a code (at least 3 chars, contains both letters and numbers)
        // This catches codes like "KING01058" even if they're part of a larger string
        var pattern2 = @"[A-Za-z]+\d+|\d+[A-Za-z]+";
        var matches2 = Regex.Matches(text, pattern2);
        foreach (Match match in matches2)
        {
            if (match.Value.Length >= 3)
            {
                identifiers.Add(match.Value);
            }
        }

        return identifiers;
    }

    /// <summary>
    /// Matches EANs from the main file against reference files and returns status values for each row.
    /// Priority: Reference A → Reference B → No Match
    /// </summary>
    [Obsolete("Use MatchEans with flexible reference files")]
    public Dictionary<int, string> MatchEans(
        Dictionary<int, string> mainEans,
        HashSet<string> refAEans,
        HashSet<string> refBEans,
        MappingResult mapping)
    {
        var rowStatuses = new Dictionary<int, string>();

        foreach (var kvp in mainEans)
        {
            var row = kvp.Key;
            var ean = kvp.Value;

            string status;
            if (refAEans.Contains(ean))
            {
                status = mapping.ReferenceAOption ?? mapping.NoMatchOption ?? string.Empty;
            }
            else if (refBEans.Contains(ean))
            {
                status = mapping.ReferenceBOption ?? mapping.NoMatchOption ?? string.Empty;
            }
            else
            {
                status = mapping.NoMatchOption ?? string.Empty;
            }

            rowStatuses[row] = status;
        }

        return rowStatuses;
    }

    /// <summary>
    /// Matches EANs from the main file against multiple reference files with priorities.
    /// Returns status values for each row based on which reference file contains the EAN.
    /// </summary>
    /// <param name="mainEans">Dictionary of row number to EAN from main file</param>
    /// <param name="referenceFileMatches">List of tuples: (EAN set, dropdown option, priority). Lower priority = checked first.</param>
    /// <param name="noMatchOption">Option to use when EAN is not found in any reference file</param>
    public Dictionary<int, string> MatchEansFlexible(
        Dictionary<int, string> mainEans,
        List<(HashSet<string> EanSet, string DropdownOption, int Priority)> referenceFileMatches,
        string noMatchOption)
    {
        var rowStatuses = new Dictionary<int, string>();

        // Sort reference files by priority (lower number = higher priority)
        var sortedReferences = referenceFileMatches.OrderBy(r => r.Priority).ToList();

        foreach (var kvp in mainEans)
        {
            var row = kvp.Key;
            var ean = kvp.Value;

            string status = noMatchOption;
            
            // Check each reference file in priority order
            foreach (var (eanSet, dropdownOption, _) in sortedReferences)
            {
                if (eanSet.Contains(ean))
                {
                    status = dropdownOption;
                    break; // Found a match, stop checking
                }
            }

            rowStatuses[row] = status;
        }

        return rowStatuses;
    }
}

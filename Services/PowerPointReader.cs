using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.IO;
using System.Text.RegularExpressions;

namespace LauraAssetBuildReview.Services;

/// <summary>
/// Service for reading EAN values from PowerPoint (.PPTX) files.
/// Extracts EANs from tables in specified slides.
/// </summary>
public class PowerPointReader
{
    /// <summary>
    /// Reads EAN values from tables in specified slides of a PowerPoint file.
    /// </summary>
    /// <param name="filePath">Path to the PowerPoint file</param>
    /// <param name="selectedSlides">List of slide indices (1-based) to read from. If empty, reads from all slides.</param>
    /// <param name="minDigits">Minimum number of digits for a valid EAN</param>
    /// <param name="maxDigits">Maximum number of digits for a valid EAN</param>
    /// <param name="allowNonNumeric">Whether to allow non-numeric EANs</param>
    /// <returns>HashSet of normalized EAN values</returns>
    public HashSet<string> ReadEansFromPowerPoint(
        string filePath,
        List<int> selectedSlides,
        int minDigits = 8,
        int maxDigits = 14,
        bool allowNonNumeric = false)
    {
        var eans = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"PowerPoint file not found: {filePath}");
        }

        using (var presentationDocument = PresentationDocument.Open(filePath, false))
        {
            var presentationPart = presentationDocument.PresentationPart;
            if (presentationPart == null)
            {
                throw new InvalidOperationException("PowerPoint file does not have a presentation part.");
            }

            var presentation = presentationPart.Presentation;
            var slideIdList = presentation?.SlideIdList;

            if (slideIdList == null)
            {
                return eans; // No slides found
            }

            var slides = slideIdList.Elements<SlideId>().ToList();
            var totalSlides = slides.Count;

            // Determine which slides to process
            var slidesToProcess = new List<int>();
            if (selectedSlides == null || selectedSlides.Count == 0)
            {
                // Process all slides
                slidesToProcess = Enumerable.Range(1, totalSlides).ToList();
            }
            else
            {
                // Process only selected slides (filter valid slide numbers)
                slidesToProcess = selectedSlides
                    .Where(s => s >= 1 && s <= totalSlides)
                    .Distinct()
                    .OrderBy(s => s)
                    .ToList();
            }

            foreach (var slideIndex in slidesToProcess)
            {
                // slideIndex is 1-based, but list is 0-based
                var slideId = slides[slideIndex - 1];
                var relationshipId = slideId.RelationshipId;
                if (relationshipId == null || string.IsNullOrEmpty(relationshipId))
                    continue;
                    
                var slidePart = (SlidePart?)presentationPart.GetPartById(relationshipId!);

                if (slidePart == null)
                    continue;

                var slide = slidePart.Slide;
                if (slide == null)
                    continue;

                // Extract EANs from all tables in this slide
                var slideEans = ExtractEansFromSlide(slide, minDigits, maxDigits, allowNonNumeric);
                foreach (var ean in slideEans)
                {
                    eans.Add(ean);
                }
            }
        }

        return eans;
    }

    /// <summary>
    /// Extracts EAN values from all tables in a slide.
    /// </summary>
    private HashSet<string> ExtractEansFromSlide(Slide slide, int minDigits, int maxDigits, bool allowNonNumeric)
    {
        var eans = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Find all tables in the slide
        var tables = slide.Descendants<DocumentFormat.OpenXml.Drawing.Table>().ToList();

        foreach (var table in tables)
        {
            // Get all text from table cells
            var cells = table.Descendants<DocumentFormat.OpenXml.Drawing.TableCell>().ToList();

            foreach (var cell in cells)
            {
                // Get text from the cell
                var textElements = cell.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();
                var cellText = string.Join(" ", textElements.Select(t => t.Text ?? string.Empty));

                if (string.IsNullOrWhiteSpace(cellText))
                    continue;

                // Extract potential EANs from the cell text
                var cellEans = ExtractEansFromText(cellText, minDigits, maxDigits, allowNonNumeric);
                foreach (var ean in cellEans)
                {
                    eans.Add(ean);
                }
            }
        }

        // Also check for EANs in regular text shapes (in case they're not in tables)
        var textShapes = slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();
        foreach (var textElement in textShapes)
        {
            var text = textElement.Text;
            if (!string.IsNullOrWhiteSpace(text))
            {
                var textEans = ExtractEansFromText(text, minDigits, maxDigits, allowNonNumeric);
                foreach (var ean in textEans)
                {
                    eans.Add(ean);
                }
            }
        }

        return eans;
    }

    /// <summary>
    /// Extracts EAN values from text using pattern matching.
    /// </summary>
    private HashSet<string> ExtractEansFromText(string text, int minDigits, int maxDigits, bool allowNonNumeric)
    {
        var eans = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (string.IsNullOrWhiteSpace(text))
            return eans;

        // Normalize the text - remove common formatting
        var normalizedText = text.Trim();

        // Pattern 1: Look for sequences of digits with the right length
        // This matches EANs that are clearly separated (e.g., "1234567890123" or "EAN: 1234567890123")
        var digitPattern = $@"\b\d{{{minDigits},{maxDigits}}}\b";
        var digitMatches = Regex.Matches(normalizedText, digitPattern);
        foreach (Match match in digitMatches)
        {
            var potentialEan = match.Value;
            if (IsValidEan(potentialEan, minDigits, maxDigits, allowNonNumeric))
            {
                eans.Add(potentialEan);
            }
        }

        // Pattern 2: Look for EANs that might have spaces or dashes (e.g., "123-456-789-0123")
        var formattedPattern = $@"\b[\d\s-]{{{minDigits},{maxDigits + 10}}}\b";
        var formattedMatches = Regex.Matches(normalizedText, formattedPattern);
        foreach (Match match in formattedMatches)
        {
            var potentialEan = NormalizeEan(match.Value);
            if (IsValidEan(potentialEan, minDigits, maxDigits, allowNonNumeric))
            {
                eans.Add(potentialEan);
            }
        }

        // Pattern 3: If the entire cell/text is a potential EAN
        var normalized = NormalizeEan(normalizedText);
        if (IsValidEan(normalized, minDigits, maxDigits, allowNonNumeric))
        {
            eans.Add(normalized);
        }

        return eans;
    }

    /// <summary>
    /// Normalizes an EAN value by removing formatting characters.
    /// </summary>
    private string NormalizeEan(string ean)
    {
        if (string.IsNullOrWhiteSpace(ean))
            return string.Empty;

        // Remove spaces, dashes, dots, and other formatting
        var cleaned = Regex.Replace(ean, @"[\s\-\.]", "");
        return cleaned.Trim();
    }

    /// <summary>
    /// Validates if a string is a valid EAN based on the criteria.
    /// </summary>
    private bool IsValidEan(string value, int minDigits, int maxDigits, bool allowNonNumeric)
    {
        if (string.IsNullOrWhiteSpace(value))
            return false;

        var cleaned = NormalizeEan(value);

        // Check if it's all digits and has the right length
        if (cleaned.All(char.IsDigit))
        {
            return cleaned.Length >= minDigits && cleaned.Length <= maxDigits;
        }

        // If non-numeric is allowed, check minimum length
        if (allowNonNumeric && cleaned.Length >= minDigits)
        {
            return true;
        }

        return false;
    }

    /// <summary>
    /// Gets the total number of slides in a PowerPoint file.
    /// </summary>
    public int GetSlideCount(string filePath)
    {
        if (!File.Exists(filePath))
        {
            return 0;
        }

        try
        {
            using (var presentationDocument = PresentationDocument.Open(filePath, false))
            {
                var presentationPart = presentationDocument.PresentationPart;
                if (presentationPart == null)
                {
                    return 0;
                }

                var presentation = presentationPart.Presentation;
                var slideIdList = presentation?.SlideIdList;

                if (slideIdList == null)
                {
                    return 0;
                }

                return slideIdList.Elements<SlideId>().Count();
            }
        }
        catch
        {
            return 0;
        }
    }
}

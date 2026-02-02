using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using LauraAssetBuildReview.Models;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace LauraAssetBuildReview.Services;

public class ExcelReader
{
    /// <summary>
    /// Reads EAN values from column C starting at the specified row.
    /// Returns a dictionary mapping row index (1-based) to normalized EAN string.
    /// Only reads values that look like EANs (numeric, typically 8-14 digits).
    /// </summary>
    [Obsolete("Use ReadEansFromColumn with column number parameter")]
    public Dictionary<int, string> ReadEansFromColumnC(IXLWorksheet worksheet, int startRow)
    {
        return ReadEansFromColumn(worksheet, 3, startRow, 8, 14, false);
    }

    /// <summary>
    /// Reads EAN values from a specified column starting at the specified row.
    /// Returns a dictionary mapping row index (1-based) to normalized EAN string.
    /// </summary>
    public Dictionary<int, string> ReadEansFromColumn(IXLWorksheet worksheet, int columnNumber, int startRow, int minDigits = 8, int maxDigits = 14, bool allowNonNumeric = false)
    {
        var eans = new Dictionary<int, string>();
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? startRow - 1;

        for (int row = startRow; row <= lastRow; row++)
        {
            var cell = worksheet.Cell(row, columnNumber);
            if (cell.IsEmpty())
                continue;

            // Get EAN value - handle both text and numeric formats
            string eanValue;
            if (cell.DataType == XLDataType.Number)
            {
                // If it's a number, get it as a string to preserve precision
                // Use G format to avoid scientific notation for large numbers
                var numValue = cell.GetDouble();
                eanValue = numValue.ToString("G", System.Globalization.CultureInfo.InvariantCulture);
                // Remove decimal point and trailing zeros if it was formatted as decimal
                if (eanValue.Contains('.'))
                {
                    eanValue = eanValue.TrimEnd('0').TrimEnd('.');
                }
            }
            else
            {
                eanValue = cell.GetString();
            }
            
            var normalizedEan = NormalizeEan(eanValue);
            
            // Only include if it looks like a valid EAN based on configuration
            if (!string.IsNullOrWhiteSpace(normalizedEan) && IsValidEan(normalizedEan, minDigits, maxDigits, allowNonNumeric))
            {
                eans[row] = normalizedEan;
            }
        }

        return eans;
    }

    /// <summary>
    /// Checks if a string looks like a valid EAN (European Article Number).
    /// EANs are typically 8, 12, 13, or 14 digits.
    /// </summary>
    private bool IsValidEan(string value, int minDigits = 8, int maxDigits = 14, bool allowNonNumeric = false)
    {
        if (string.IsNullOrWhiteSpace(value))
            return false;

        // Remove any formatting characters
        var cleaned = value.Replace("-", "").Replace(" ", "").Replace(".", "").Trim();
        
        // Check if it's all digits and has a valid EAN length
        if (cleaned.All(char.IsDigit))
        {
            var length = cleaned.Length;
            return length >= minDigits && length <= maxDigits;
        }
        
        // If non-numeric is allowed, check if it has minimum length
        if (allowNonNumeric && cleaned.Length >= minDigits)
        {
            return true;
        }
        
        return false;
    }

    /// <summary>
    /// Reads all EAN values from column C in the worksheet.
    /// Scans from row 1 to the last used row.
    /// Only reads values that look like EANs (numeric, typically 8-14 digits).
    /// </summary>
    [Obsolete("Use ReadAllEansFromColumn with column number parameter")]
    public HashSet<string> ReadAllEansFromColumnC(IXLWorksheet worksheet)
    {
        return ReadAllEansFromColumn(worksheet, 3, 1, 8, 14, false);
    }

    /// <summary>
    /// Reads all EAN values from a specified column in the worksheet.
    /// Scans from startRow to the last used row.
    /// </summary>
    public HashSet<string> ReadAllEansFromColumn(IXLWorksheet worksheet, int columnNumber, int startRow = 1, int minDigits = 8, int maxDigits = 14, bool allowNonNumeric = false)
    {
        var eans = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;

        for (int row = startRow; row <= lastRow; row++)
        {
            var cell = worksheet.Cell(row, columnNumber);
            if (cell.IsEmpty())
                continue;

            // Get EAN value - handle both text and numeric formats
            string eanValue;
            if (cell.DataType == XLDataType.Number)
            {
                // If it's a number, get it as a string to preserve precision
                // Use G format to avoid scientific notation for large numbers
                var numValue = cell.GetDouble();
                eanValue = numValue.ToString("G", System.Globalization.CultureInfo.InvariantCulture);
                // Remove decimal point and trailing zeros if it was formatted as decimal
                if (eanValue.Contains('.'))
                {
                    eanValue = eanValue.TrimEnd('0').TrimEnd('.');
                }
            }
            else
            {
                eanValue = cell.GetString();
            }
            
            var normalizedEan = NormalizeEan(eanValue);
            
            // Only include if it looks like a valid EAN based on configuration
            if (!string.IsNullOrWhiteSpace(normalizedEan) && IsValidEan(normalizedEan, minDigits, maxDigits, allowNonNumeric))
            {
                eans.Add(normalizedEan);
            }
        }

        return eans;
    }

    /// <summary>
    /// Reads dropdown validation options from a specified column.
    /// Uses DocumentFormat.OpenXml directly since ClosedXML's DataValidations collection may be empty.
    /// </summary>
    public List<string> ReadDropdownOptions(IXLWorksheet worksheet, int column = 7, string? filePath = null)
    {
        var options = new List<string>();
        var columnLetter = ProcessingConfiguration.ColumnNumberToLetter(column);
        var columnLetterUpper = columnLetter.ToUpperInvariant();

        // Try to get file path from workbook if not provided
        if (string.IsNullOrEmpty(filePath))
        {
            try
            {
                var workbook = worksheet.Workbook;
                var workbookType = workbook?.GetType();
                var fileProperty = workbookType?.GetProperty("File", BindingFlags.NonPublic | BindingFlags.Instance) ??
                                  workbookType?.GetProperty("_file", BindingFlags.NonPublic | BindingFlags.Instance);
                
                if (fileProperty != null)
                {
                    var fileObj = fileProperty.GetValue(workbook);
                    if (fileObj != null)
                    {
                        var pathProperty = fileObj.GetType().GetProperty("FullName") ?? 
                                         fileObj.GetType().GetProperty("Name");
                        if (pathProperty != null)
                        {
                            filePath = pathProperty.GetValue(fileObj) as string;
                        }
                    }
                }
            }
            catch
            {
                // If we can't get file path, return empty list
            }
        }

        // Use DocumentFormat.OpenXml directly (most reliable method)
        if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
        {
            try
            {
                using (var doc = SpreadsheetDocument.Open(filePath, false))
                {
                    var workbookPart = doc.WorkbookPart;
                    if (workbookPart != null)
                    {
                        var worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                        if (worksheetPart != null)
                        {
                            var openXmlWorksheet = worksheetPart.Worksheet;
                            var dataValidations = openXmlWorksheet.Descendants<DataValidations>().FirstOrDefault();
                            
                            if (dataValidations != null)
                            {
                                foreach (var dv in dataValidations.Elements<DataValidation>())
                                {
                                    if (dv.Type?.Value == DataValidationValues.List)
                                    {
                                        var formula1 = dv.Formula1;
                                        if (formula1 != null && formula1.Text != null)
                                        {
                                            var formula1Text = formula1.Text;
                                            
                                            // Check if this applies to our column
                                            var sqref = dv.SequenceOfReferences;
                                            bool appliesToColumn = false;
                                            
                                            if (sqref != null)
                                            {
                                                foreach (var refValue in sqref)
                                                {
                                                    if (refValue?.Value != null)
                                                    {
                                                        var refValueStr = refValue.Value.ToUpperInvariant();
                                                        
                                                        // Check if the reference contains our column letter
                                                        // Examples: "F1:F100", "F:F", "$F$1:$F$100", "F3:F1000", "F1", "F"
                                                        // We need to ensure it's actually for this column, not part of another (e.g., "AF" when looking for "F")
                                                        
                                                        // Simple approach: look for column letter followed by digit or colon or end of string
                                                        // This handles: F1, F:F, F1:F100, $F$1:$F$100, etc.
                                                        var escapedColumn = System.Text.RegularExpressions.Regex.Escape(columnLetterUpper);
                                                        
                                                        // Pattern matches:
                                                        // - Column letter at start or after $ or : or space
                                                        // - Followed by optional $, then optional digits
                                                        // - Or followed by : and same column
                                                        // Examples that match for "F": "F1", "F:F", "F1:F100", "$F$1:$F$100", " F3:F1000"
                                                        // Examples that DON'T match for "F": "AF1" (A before F), "BF2" (B before F)
                                                        var pattern = $@"(^|[\$:\s]){escapedColumn}(\$?\d+|:\$?{escapedColumn}\$?\d*|$)";
                                                        
                                                        if (System.Text.RegularExpressions.Regex.IsMatch(refValueStr, pattern))
                                                        {
                                                            appliesToColumn = true;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                            
                                            // Fallback: If SequenceOfReferences is empty or doesn't match, but we have a comma-separated list,
                                            // try a simpler check - just see if the column letter appears anywhere in the reference
                                            // This handles cases where Excel's SequenceOfReferences format might be unexpected
                                            if (!appliesToColumn && sqref != null)
                                            {
                                                foreach (var refValue in sqref)
                                                {
                                                    if (refValue?.Value != null)
                                                    {
                                                        var refValueStr = refValue.Value.ToUpperInvariant();
                                                        // Simple check: does the reference string contain our column letter?
                                                        // We'll be more permissive here as a fallback
                                                        if (refValueStr.Contains(columnLetterUpper))
                                                        {
                                                            // Additional safety: make sure it's not clearly for another column
                                                            // e.g., if looking for "F", don't match "AF" or "BF"
                                                            // But do match "F1", "F:F", "$F$", etc.
                                                            var beforeChar = refValueStr.IndexOf(columnLetterUpper);
                                                            if (beforeChar > 0)
                                                            {
                                                                var charBefore = refValueStr[beforeChar - 1];
                                                                // If the character before is a letter, it's part of another column (e.g., "AF")
                                                                if (char.IsLetter(charBefore))
                                                                {
                                                                    continue; // Skip this one
                                                                }
                                                            }
                                                            appliesToColumn = true;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                            
                                            // Only process if it applies to our column
                                            if (appliesToColumn)
                                            {
                                                // Parse the list source (comma-separated or range reference)
                                                if (!formula1Text.Contains(':'))
                                                {
                                                    // Comma-separated list
                                                    ParseCommaSeparatedList(formula1Text, options);
                                                    if (options.Count > 0)
                                                    {
                                                        return options; // Success!
                                                    }
                                                }
                                                else
                                                {
                                                    // Range reference - try to read values from the range
                                                    // This handles cases where dropdown references a cell range
                                                    try
                                                    {
                                                        var rangeValues = ReadDropdownFromRange(worksheet, formula1Text);
                                                        if (rangeValues.Count > 0)
                                                        {
                                                            options.AddRange(rangeValues);
                                                            return options;
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        // If we can't read from range, continue to next validation
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                // If file-based approach fails, return empty list
            }
        }

        return options;
    }

    /// <summary>
    /// Reads dropdown options from a cell range reference (e.g., "Sheet1!$A$1:$A$5").
    /// </summary>
    private List<string> ReadDropdownFromRange(IXLWorksheet worksheet, string rangeFormula)
    {
        var values = new List<string>();
        
        try
        {
            // Parse range like "Sheet1!$A$1:$A$5" or "$A$1:$A$5" or "A1:A5"
            var range = worksheet.Workbook.Range(rangeFormula);
            if (range != null)
            {
                foreach (var cell in range.Cells())
                {
                    if (!cell.IsEmpty())
                    {
                        var value = cell.GetString();
                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            values.Add(value.Trim());
                        }
                    }
                }
            }
        }
        catch
        {
            // If we can't parse the range, return empty list
        }
        
        return values;
    }

    private void ParseCommaSeparatedList(string listValue, List<string> options)
    {
        // Split by comma, handling potential whitespace
        // Example: "Recieved - KING01042,Missing - Need to request,Recieved - KING01058"
        var separators = new[] { ',' };
        var values = listValue.Split(separators, StringSplitOptions.RemoveEmptyEntries);
        foreach (var value in values)
        {
            var trimmed = value.Trim();
            
            // Remove leading and trailing quotes if present
            while (trimmed.StartsWith("\"") || trimmed.StartsWith("'"))
            {
                trimmed = trimmed.Substring(1);
            }
            while (trimmed.EndsWith("\"") || trimmed.EndsWith("'"))
            {
                trimmed = trimmed.Substring(0, trimmed.Length - 1);
            }
            
            trimmed = trimmed.Trim();
            
            if (!string.IsNullOrEmpty(trimmed))
            {
                options.Add(trimmed);
            }
        }
    }

    /// <summary>
    /// Normalizes an EAN value: trims whitespace and preserves leading zeros.
    /// Handles numeric formats and ensures consistent comparison.
    /// </summary>
    public string NormalizeEan(string ean)
    {
        if (string.IsNullOrWhiteSpace(ean))
            return string.Empty;

        var trimmed = ean.Trim();
        
        // Remove any non-digit characters except for potential separators
        // EANs are typically numeric, but might have formatting
        // Remove common formatting characters but preserve the numeric value
        var cleaned = trimmed.Replace("-", "").Replace(" ", "").Replace(".", "");
        
        // If it's all digits, return as-is (preserves leading zeros)
        if (cleaned.All(char.IsDigit))
        {
            return cleaned;
        }
        
        // If it contains non-digits, return trimmed (might be a code like "KING01058")
        return trimmed;
    }

    /// <summary>
    /// Gets the last used row number in the worksheet.
    /// </summary>
    public int GetLastUsedRow(IXLWorksheet worksheet)
    {
        return worksheet.LastRowUsed()?.RowNumber() ?? 0;
    }
}

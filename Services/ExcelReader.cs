using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
    public Dictionary<int, string> ReadEansFromColumnC(IXLWorksheet worksheet, int startRow)
    {
        var eans = new Dictionary<int, string>();
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? startRow - 1;

        for (int row = startRow; row <= lastRow; row++)
        {
            var cell = worksheet.Cell(row, 3); // Column C
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
            
            // Only include if it looks like an EAN (numeric, typically 8-14 digits)
            // EANs are usually 8, 12, 13, or 14 digits
            if (!string.IsNullOrWhiteSpace(normalizedEan) && IsValidEan(normalizedEan))
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
    private bool IsValidEan(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return false;

        // Remove any formatting characters
        var cleaned = value.Replace("-", "").Replace(" ", "").Replace(".", "").Trim();
        
        // Check if it's all digits and has a valid EAN length (8, 12, 13, or 14 digits)
        if (cleaned.All(char.IsDigit))
        {
            var length = cleaned.Length;
            return length >= 8 && length <= 14; // EAN-8, EAN-12, EAN-13, or EAN-14
        }
        
        return false;
    }

    /// <summary>
    /// Reads all EAN values from column C in the worksheet.
    /// Scans from row 1 to the last used row.
    /// Only reads values that look like EANs (numeric, typically 8-14 digits).
    /// </summary>
    public HashSet<string> ReadAllEansFromColumnC(IXLWorksheet worksheet)
    {
        var eans = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;

        for (int row = 1; row <= lastRow; row++)
        {
            var cell = worksheet.Cell(row, 3); // Column C
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
            
            // Only include if it looks like an EAN (numeric, typically 8-14 digits)
            if (!string.IsNullOrWhiteSpace(normalizedEan) && IsValidEan(normalizedEan))
            {
                eans.Add(normalizedEan);
            }
        }

        return eans;
    }

    /// <summary>
    /// Reads dropdown validation options from column G.
    /// Uses DocumentFormat.OpenXml directly since ClosedXML's DataValidations collection may be empty.
    /// </summary>
    public List<string> ReadDropdownOptions(IXLWorksheet worksheet, int column = 7, string? filePath = null) // Column G
    {
        var options = new List<string>();

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
                                                    if (refValue?.Value != null && refValue.Value.Contains("G"))
                                                    {
                                                        appliesToColumn = true;
                                                        break;
                                                    }
                                                }
                                            }
                                            
                                            // If it applies to our column, or if it's a comma-separated list (not a range reference)
                                            if (appliesToColumn || (!formula1Text.Contains(':') && formula1Text.Contains(',')))
                                            {
                                                // Parse the list source (comma-separated)
                                                if (!formula1Text.Contains(':'))
                                                {
                                                    ParseCommaSeparatedList(formula1Text, options);
                                                    if (options.Count > 0)
                                                    {
                                                        return options; // Success!
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

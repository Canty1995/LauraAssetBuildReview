using ClosedXML.Excel;
using LauraAssetBuildReview.Models;

namespace LauraAssetBuildReview.Services;

public class FileComparisonResult
{
    public int TotalRowsCompared { get; set; }
    public int MatchingRows { get; set; }
    public int MismatchingRows { get; set; }
    public int MissingInFile1 { get; set; }
    public int MissingInFile2 { get; set; }
    public List<RowComparison> Mismatches { get; set; } = new();
    public bool IsIdentical { get; set; }
}

public class RowComparison
{
    public int Row { get; set; }
    public string? File1Value { get; set; }
    public string? File2Value { get; set; }
}

public class FileComparisonService
{
    private readonly ExcelReader _excelReader;

    public FileComparisonService(ExcelReader excelReader)
    {
        _excelReader = excelReader;
    }

    /// <summary>
    /// Compares column G values between two Excel files.
    /// </summary>
    /// <param name="file1Path">Path to the first file (e.g., manually populated)</param>
    /// <param name="file2Path">Path to the second file (e.g., program-generated)</param>
    /// <param name="startRow">Starting row to compare (default: 3)</param>
    /// <returns>Comparison result with details of matches and mismatches</returns>
    [Obsolete("Use CompareColumn with column number parameter")]
    public FileComparisonResult CompareColumnG(string file1Path, string file2Path, int startRow = 3)
    {
        return CompareColumn(file1Path, file2Path, startRow, 7);
    }

    /// <summary>
    /// Compares values in a specified column between two Excel files.
    /// </summary>
    /// <param name="file1Path">Path to the first file (e.g., manually populated)</param>
    /// <param name="file2Path">Path to the second file (e.g., program-generated)</param>
    /// <param name="startRow">Starting row to compare</param>
    /// <param name="columnNumber">Column number (1-based) to compare</param>
    /// <returns>Comparison result with details of matches and mismatches</returns>
    public FileComparisonResult CompareColumn(string file1Path, string file2Path, int startRow, int columnNumber)
    {
        var result = new FileComparisonResult();

        // Read column values from both files
        Dictionary<int, string> file1Values;
        Dictionary<int, string> file2Values;

        using (var workbook1 = new XLWorkbook(file1Path))
        {
            var worksheet1 = workbook1.Worksheet(1);
            if (worksheet1 == null)
            {
                throw new InvalidOperationException($"File 1 does not have a first worksheet: {file1Path}");
            }
            file1Values = ReadColumn(worksheet1, startRow, columnNumber);
        }

        using (var workbook2 = new XLWorkbook(file2Path))
        {
            var worksheet2 = workbook2.Worksheet(1);
            if (worksheet2 == null)
            {
                throw new InvalidOperationException($"File 2 does not have a first worksheet: {file2Path}");
            }
            file2Values = ReadColumn(worksheet2, startRow, columnNumber);
        }

        // Get all unique row numbers
        var allRows = new HashSet<int>();
        foreach (var row in file1Values.Keys) allRows.Add(row);
        foreach (var row in file2Values.Keys) allRows.Add(row);

        result.TotalRowsCompared = allRows.Count;

        // Compare values
        foreach (var row in allRows.OrderBy(r => r))
        {
            var hasFile1 = file1Values.TryGetValue(row, out var value1);
            var hasFile2 = file2Values.TryGetValue(row, out var value2);

            if (!hasFile1)
            {
                result.MissingInFile1++;
                result.Mismatches.Add(new RowComparison
                {
                    Row = row,
                    File1Value = null,
                    File2Value = value2
                });
            }
            else if (!hasFile2)
            {
                result.MissingInFile2++;
                result.Mismatches.Add(new RowComparison
                {
                    Row = row,
                    File1Value = value1,
                    File2Value = null
                });
            }
            else
            {
                // Both have values - compare them
                var normalized1 = NormalizeValue(value1);
                var normalized2 = NormalizeValue(value2);

                if (string.Equals(normalized1, normalized2, StringComparison.OrdinalIgnoreCase))
                {
                    result.MatchingRows++;
                }
                else
                {
                    result.MismatchingRows++;
                    result.Mismatches.Add(new RowComparison
                    {
                        Row = row,
                        File1Value = value1,
                        File2Value = value2
                    });
                }
            }
        }

        result.IsIdentical = result.MismatchingRows == 0 && result.MissingInFile1 == 0 && result.MissingInFile2 == 0;

        return result;
    }

    /// <summary>
    /// Reads all values from column G starting at the specified row.
    /// </summary>
    [Obsolete("Use ReadColumn with column number parameter")]
    private Dictionary<int, string> ReadColumnG(IXLWorksheet worksheet, int startRow)
    {
        return ReadColumn(worksheet, startRow, 7);
    }

    /// <summary>
    /// Reads all values from a specified column starting at the specified row.
    /// </summary>
    private Dictionary<int, string> ReadColumn(IXLWorksheet worksheet, int startRow, int columnNumber)
    {
        var values = new Dictionary<int, string>();
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? startRow - 1;

        for (int row = startRow; row <= lastRow; row++)
        {
            var cell = worksheet.Cell(row, columnNumber);
            if (!cell.IsEmpty())
            {
                var value = cell.GetString();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    // Store the raw value (normalization happens during comparison)
                    values[row] = value.Trim();
                }
            }
        }

        return values;
    }

    /// <summary>
    /// Normalizes a value for comparison (trim, remove quotes, case-insensitive).
    /// </summary>
    private string NormalizeValue(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var normalized = value.Trim();
        
        // Remove leading and trailing quotes (both single and double)
        // Handle cases like '"Recieved - KING01042' or 'Recieved - KING01058""'
        while (normalized.StartsWith("\"") || normalized.StartsWith("'"))
        {
            normalized = normalized.Substring(1);
        }
        while (normalized.EndsWith("\"") || normalized.EndsWith("'"))
        {
            normalized = normalized.Substring(0, normalized.Length - 1);
        }
        
        return normalized.Trim();
    }
}

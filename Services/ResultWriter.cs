using ClosedXML.Excel;
using System.IO;

namespace LauraAssetBuildReview.Services;

/// <summary>
/// Service for creating result Excel files with EANs and their reference file matches.
/// </summary>
public class ResultWriter
{
    /// <summary>
    /// Creates a new Excel file listing all EANs and which reference files contain them.
    /// </summary>
    /// <param name="eans">List of all EANs found in the main file</param>
    /// <param name="referenceFileMatches">Dictionary mapping EAN to list of reference file names that contain it</param>
    /// <param name="outputFileName">Name for the output file (without extension)</param>
    /// <returns>Path to the created Excel file</returns>
    public string CreateResultFile(
        List<string> eans,
        Dictionary<string, List<string>> referenceFileMatches,
        string outputFileName = "EAN_Results")
    {
        // Get the application root directory
        var appRoot = AppDomain.CurrentDomain.BaseDirectory;
        var processedFolder = Path.Combine(appRoot, "processed");

        // Create folder if it doesn't exist
        if (!Directory.Exists(processedFolder))
        {
            Directory.CreateDirectory(processedFolder);
        }

        // Create output file path
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var outputPath = Path.Combine(processedFolder, $"{outputFileName}_{timestamp}.xlsx");

        // Create new workbook
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("EAN Results");

        // Add headers
        worksheet.Cell(1, 1).Value = "EAN";
        worksheet.Cell(1, 2).Value = "Found In Files";
        
        // Style headers
        var headerRange = worksheet.Range(1, 1, 1, 2);
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        // Add data
        int row = 2;
        foreach (var ean in eans.OrderBy(e => e))
        {
            worksheet.Cell(row, 1).Value = ean;
            
            // Get list of files that contain this EAN
            if (referenceFileMatches.TryGetValue(ean, out var files) && files.Count > 0)
            {
                worksheet.Cell(row, 2).Value = string.Join(", ", files);
            }
            else
            {
                worksheet.Cell(row, 2).Value = "Not Found";
            }
            
            row++;
        }

        // Auto-fit columns
        worksheet.Column(1).Width = 20;
        worksheet.Column(2).Width = 50;

        // Save the workbook
        workbook.SaveAs(outputPath);

        return outputPath;
    }
}

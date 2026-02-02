using ClosedXML.Excel;
using LauraAssetBuildReview.Models;
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
    /// <param name="eanCountsPerSlide">Optional dictionary mapping slide number to EAN count for summary sheet</param>
    /// <param name="eansPerSlide">Optional dictionary mapping slide number to list of EANs with text context</param>
    /// <returns>Path to the created Excel file</returns>
    public string CreateResultFile(
        List<string> eans,
        Dictionary<string, List<string>> referenceFileMatches,
        string outputFileName = "EAN_Results",
        Dictionary<int, int>? eanCountsPerSlide = null,
        Dictionary<int, List<EanInfo>>? eansPerSlide = null)
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

        // Add summary sheet with per-slide EAN counts and details if provided
        if ((eanCountsPerSlide != null && eanCountsPerSlide.Count > 0) || 
            (eansPerSlide != null && eansPerSlide.Count > 0))
        {
            var summaryWorksheet = workbook.Worksheets.Add("Slide Summary");
            
            // Add headers
            summaryWorksheet.Cell(1, 1).Value = "Slide Number";
            summaryWorksheet.Cell(1, 2).Value = "EAN Number";
            summaryWorksheet.Cell(1, 3).Value = "Text Context";
            
            // Style headers
            var summaryHeaderRange = summaryWorksheet.Range(1, 1, 1, 3);
            summaryHeaderRange.Style.Font.Bold = true;
            summaryHeaderRange.Style.Fill.BackgroundColor = XLColor.LightGray;
            summaryHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            summaryHeaderRange.Style.Alignment.WrapText = true;
            
            // Add data - one row per EAN for clarity
            int summaryRow = 2;
            int totalEans = 0;
            
            // Use eansPerSlide if available, otherwise fall back to eanCountsPerSlide
            var slidesToProcess = eansPerSlide?.Keys.OrderBy(x => x).ToList() ?? 
                                 eanCountsPerSlide?.Keys.OrderBy(x => x).ToList() ?? 
                                 new List<int>();
            
            foreach (var slideNumber in slidesToProcess)
            {
                if (eansPerSlide != null && eansPerSlide.TryGetValue(slideNumber, out var slideEans) && slideEans.Count > 0)
                {
                    var eanCount = slideEans.Count;
                    totalEans += eanCount;
                    
                    // Create one row per EAN with its corresponding text
                    foreach (var eanInfo in slideEans)
                    {
                        summaryWorksheet.Cell(summaryRow, 1).Value = slideNumber;
                        summaryWorksheet.Cell(summaryRow, 2).Value = eanInfo.Ean;
                        
                        // Truncate very long contexts to keep the sheet readable
                        var contextText = eanInfo.TextContext;
                        if (!string.IsNullOrWhiteSpace(contextText))
                        {
                            if (contextText.Length > 500)
                            {
                                contextText = contextText.Substring(0, 500) + "...";
                            }
                            summaryWorksheet.Cell(summaryRow, 3).Value = contextText;
                        }
                        else
                        {
                            summaryWorksheet.Cell(summaryRow, 3).Value = "(No text context)";
                        }
                        
                        // Enable text wrapping for context column
                        summaryWorksheet.Cell(summaryRow, 3).Style.Alignment.WrapText = true;
                        summaryRow++;
                    }
                }
                else if (eanCountsPerSlide != null && eanCountsPerSlide.TryGetValue(slideNumber, out var count))
                {
                    // If we only have counts but not details, show a summary row
                    summaryWorksheet.Cell(summaryRow, 1).Value = slideNumber;
                    summaryWorksheet.Cell(summaryRow, 2).Value = $"{count} EAN(s)";
                    summaryWorksheet.Cell(summaryRow, 3).Value = "N/A - Details not available";
                    totalEans += count;
                    summaryRow++;
                }
            }
            
            // Add total row
            summaryWorksheet.Cell(summaryRow, 1).Value = "Total";
            summaryWorksheet.Cell(summaryRow, 1).Style.Font.Bold = true;
            summaryWorksheet.Cell(summaryRow, 2).Value = totalEans;
            summaryWorksheet.Cell(summaryRow, 2).Style.Font.Bold = true;
            summaryWorksheet.Cell(summaryRow, 3).Value = "EANs found";
            
            // Auto-fit columns
            summaryWorksheet.Column(1).Width = 15;
            summaryWorksheet.Column(2).Width = 20;
            summaryWorksheet.Column(3).Width = 80;
            
            // Set row heights for better readability
            for (int rowIndex = 2; rowIndex < summaryRow; rowIndex++)
            {
                summaryWorksheet.Row(rowIndex).Height = 40; // Allow room for wrapped text
            }
        }

        // Save the workbook
        workbook.SaveAs(outputPath);

        return outputPath;
    }
}

using ClosedXML.Excel;
using System.IO;

namespace LauraAssetBuildReview.Services;

public class ExcelWriter
{
    /// <summary>
    /// Writes status values to column G for the specified rows.
    /// </summary>
    /// <param name="worksheet">The worksheet to write to</param>
    /// <param name="rowStatuses">Dictionary mapping row number (1-based) to status value</param>
    [Obsolete("Use WriteStatuses with column number parameter")]
    public void WriteStatuses(IXLWorksheet worksheet, Dictionary<int, string> rowStatuses)
    {
        WriteStatuses(worksheet, rowStatuses, 7);
    }

    /// <summary>
    /// Writes status values to a specified column for the specified rows.
    /// </summary>
    /// <param name="worksheet">The worksheet to write to</param>
    /// <param name="rowStatuses">Dictionary mapping row number (1-based) to status value</param>
    /// <param name="columnNumber">Column number (1-based) to write to</param>
    public void WriteStatuses(IXLWorksheet worksheet, Dictionary<int, string> rowStatuses, int columnNumber)
    {
        foreach (var kvp in rowStatuses)
        {
            var row = kvp.Key;
            var status = kvp.Value;
            
            // Clean the status value - remove any leading/trailing quotes
            var cleanedStatus = CleanStatusValue(status);
            
            // Set the cell value directly (not as a formula or string with quotes)
            worksheet.Cell(row, columnNumber).Value = cleanedStatus;
        }
    }

    /// <summary>
    /// Cleans a status value by removing unwanted quotes.
    /// </summary>
    private string CleanStatusValue(string status)
    {
        if (string.IsNullOrWhiteSpace(status))
            return string.Empty;

        var cleaned = status.Trim();
        
        // Remove leading and trailing quotes
        while (cleaned.StartsWith("\"") || cleaned.StartsWith("'"))
        {
            cleaned = cleaned.Substring(1);
        }
        while (cleaned.EndsWith("\"") || cleaned.EndsWith("'"))
        {
            cleaned = cleaned.Substring(0, cleaned.Length - 1);
        }
        
        return cleaned.Trim();
    }

    /// <summary>
    /// Safely overwrites the original file by writing to a temporary file first,
    /// then replacing the original. This prevents corruption if the process crashes.
    /// </summary>
    /// <param name="workbook">The workbook to save</param>
    /// <param name="originalPath">Path to the original file</param>
    [Obsolete("Use SaveToProcessedFolder instead")]
    public void SafeOverwrite(XLWorkbook workbook, string originalPath)
    {
        SaveToProcessedFolder(workbook, originalPath);
    }

    /// <summary>
    /// Saves the processed file to a "processed" folder and moves the original to an "original" folder.
    /// Both folders are created in the application root directory.
    /// </summary>
    /// <param name="workbook">The workbook to save</param>
    /// <param name="originalPath">Path to the original file</param>
    /// <returns>Path to the saved processed file</returns>
    public string SaveToProcessedFolder(XLWorkbook workbook, string originalPath)
    {
        // Get the application root directory (where the executable is running from)
        var appRoot = AppDomain.CurrentDomain.BaseDirectory;
        var processedFolder = Path.Combine(appRoot, "processed");
        var originalFolder = Path.Combine(appRoot, "original");

        // Create folders if they don't exist
        if (!Directory.Exists(processedFolder))
        {
            Directory.CreateDirectory(processedFolder);
        }
        if (!Directory.Exists(originalFolder))
        {
            Directory.CreateDirectory(originalFolder);
        }

        var fileName = Path.GetFileName(originalPath);
        var processedFilePath = Path.Combine(processedFolder, fileName);
        var originalFilePath = Path.Combine(originalFolder, fileName);

        // Handle duplicate filenames by adding a timestamp
        if (File.Exists(processedFilePath))
        {
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(originalPath);
            var extension = Path.GetExtension(originalPath);
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            processedFilePath = Path.Combine(processedFolder, $"{fileNameWithoutExt}_{timestamp}{extension}");
        }

        // Handle duplicate filenames in original folder
        if (File.Exists(originalFilePath))
        {
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(originalPath);
            var extension = Path.GetExtension(originalPath);
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            originalFilePath = Path.Combine(originalFolder, $"{fileNameWithoutExt}_{timestamp}{extension}");
        }

        try
        {
            // Save the processed workbook to the processed folder
            workbook.SaveAs(processedFilePath);

            // Verify the processed file was created
            if (!File.Exists(processedFilePath))
            {
                throw new IOException($"Processed file was not created: {processedFilePath}");
            }

            var processedFileInfo = new FileInfo(processedFilePath);
            if (processedFileInfo.Length == 0)
            {
                throw new IOException($"Processed file is empty: {processedFilePath}");
            }

            // Wait a bit to ensure file system has flushed
            System.Threading.Thread.Sleep(100);

            // Move the original file to the original folder
            if (File.Exists(originalPath))
            {
                // Try to remove read-only attribute if present
                var fileInfo = new FileInfo(originalPath);
                if (fileInfo.IsReadOnly)
                {
                    fileInfo.IsReadOnly = false;
                }

                // Move the original file to the original folder
                File.Move(originalPath, originalFilePath);
            }
            else
            {
                // If original file doesn't exist (shouldn't happen, but handle gracefully)
                throw new FileNotFoundException($"Original file not found: {originalPath}");
            }

            return processedFilePath;
        }
        catch (Exception ex)
        {
            // Clean up processed file if it exists
            if (File.Exists(processedFilePath))
            {
                try
                {
                    File.Delete(processedFilePath);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }

            // Re-throw with more context
            throw new IOException($"Failed to save processed file: {ex.Message}", ex);
        }
    }
}

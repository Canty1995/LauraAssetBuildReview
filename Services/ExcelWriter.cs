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
    public void WriteStatuses(IXLWorksheet worksheet, Dictionary<int, string> rowStatuses)
    {
        foreach (var kvp in rowStatuses)
        {
            var row = kvp.Key;
            var status = kvp.Value;
            
            // Clean the status value - remove any leading/trailing quotes
            var cleanedStatus = CleanStatusValue(status);
            
            // Set the cell value directly (not as a formula or string with quotes)
            worksheet.Cell(row, 7).Value = cleanedStatus;
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
    public void SafeOverwrite(XLWorkbook workbook, string originalPath)
    {
        var directory = Path.GetDirectoryName(originalPath) ?? string.Empty;
        var fileName = Path.GetFileNameWithoutExtension(originalPath);
        var extension = Path.GetExtension(originalPath);
        var tempPath = Path.Combine(directory, $"{fileName}_temp_{Guid.NewGuid()}{extension}");

        try
        {
            // Save to temporary file
            // Use default save options to preserve all workbook parts including styles
            workbook.SaveAs(tempPath);

            // Ensure the file is fully written
            // The SaveAs method should handle flushing, but we'll verify the file exists
            if (!File.Exists(tempPath))
            {
                throw new IOException($"Temporary file was not created: {tempPath}");
            }

            // Verify the temp file is not empty
            var tempFileInfo = new FileInfo(tempPath);
            if (tempFileInfo.Length == 0)
            {
                throw new IOException($"Temporary file is empty: {tempPath}");
            }

            // Wait a bit to ensure file system has flushed
            System.Threading.Thread.Sleep(100);

            // Replace original file
            // Delete the original first to avoid issues with file locks
            if (File.Exists(originalPath))
            {
                // Try to remove read-only attribute if present
                var fileInfo = new FileInfo(originalPath);
                if (fileInfo.IsReadOnly)
                {
                    fileInfo.IsReadOnly = false;
                }
                
                // Delete the original file
                File.Delete(originalPath);
                
                // Wait a bit after deletion
                System.Threading.Thread.Sleep(100);
            }

            // Move temp file to original location (more atomic than copy)
            File.Move(tempPath, originalPath);
            
            // Verify the final file exists and is not empty
            var finalFileInfo = new FileInfo(originalPath);
            if (!finalFileInfo.Exists || finalFileInfo.Length == 0)
            {
                throw new IOException($"Final file verification failed: {originalPath}");
            }
        }
        catch (Exception ex)
        {
            // Clean up temp file if it exists
            if (File.Exists(tempPath))
            {
                try
                {
                    File.Delete(tempPath);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
            
            // Re-throw with more context
            throw new IOException($"Failed to save workbook to {originalPath}: {ex.Message}", ex);
        }
    }
}

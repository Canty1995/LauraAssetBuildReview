using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ClosedXML.Excel;
using LauraAssetBuildReview.Models;
using LauraAssetBuildReview.Services;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;

namespace LauraAssetBuildReview.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private readonly ExcelReader _excelReader;
    private readonly ExcelWriter _excelWriter;
    private readonly MatchingService _matchingService;
    private readonly LoggingService _loggingService;
    private readonly FileComparisonService _fileComparisonService;

    [ObservableProperty]
    private string _mainFilePath = string.Empty;

    [ObservableProperty]
    private string _referenceAPath = string.Empty;

    [ObservableProperty]
    private string _referenceBPath = string.Empty;

    [ObservableProperty]
    private string _comparisonFile1Path = string.Empty;

    [ObservableProperty]
    private string _comparisonFile2Path = string.Empty;

    [ObservableProperty]
    private bool _isProcessing;

    [ObservableProperty]
    private string _statusMessage = "Ready";

    public ObservableCollection<string> LogMessages { get; } = new();

    public MainViewModel()
    {
        _excelReader = new ExcelReader();
        _excelWriter = new ExcelWriter();
        _matchingService = new MatchingService();
        _loggingService = new LoggingService();
        _fileComparisonService = new FileComparisonService(_excelReader);
    }

    [RelayCommand]
    private void BrowseMainFile()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
            Title = "Select Main Excel File"
        };

        if (dialog.ShowDialog() == true)
        {
            MainFilePath = dialog.FileName;
            _loggingService.Log($"Selected main file: {MainFilePath}");
            UpdateLogDisplay();
        }
    }

    [RelayCommand]
    private void BrowseReferenceA()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
            Title = "Select Reference File A"
        };

        if (dialog.ShowDialog() == true)
        {
            ReferenceAPath = dialog.FileName;
            _loggingService.Log($"Selected reference file A: {ReferenceAPath}");
            UpdateLogDisplay();
        }
    }

    [RelayCommand]
    private void BrowseReferenceB()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
            Title = "Select Reference File B"
        };

        if (dialog.ShowDialog() == true)
        {
            ReferenceBPath = dialog.FileName;
            _loggingService.Log($"Selected reference file B: {ReferenceBPath}");
            UpdateLogDisplay();
        }
    }

    [RelayCommand]
    private void BrowseComparisonFile1()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
            Title = "Select First File for Comparison (e.g., Manually Populated)"
        };

        if (dialog.ShowDialog() == true)
        {
            ComparisonFile1Path = dialog.FileName;
            _loggingService.Log($"Selected comparison file 1: {ComparisonFile1Path}");
            UpdateLogDisplay();
        }
    }

    [RelayCommand]
    private void BrowseComparisonFile2()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
            Title = "Select Second File for Comparison (e.g., Program-Generated)"
        };

        if (dialog.ShowDialog() == true)
        {
            ComparisonFile2Path = dialog.FileName;
            _loggingService.Log($"Selected comparison file 2: {ComparisonFile2Path}");
            UpdateLogDisplay();
        }
    }

    [RelayCommand]
    private async Task CompareFiles()
    {
        if (IsProcessing)
            return;

        IsProcessing = true;
        StatusMessage = "Comparing files...";
        _loggingService.Clear();
        LogMessages.Clear();

        try
        {
            // Validate inputs
            if (string.IsNullOrWhiteSpace(ComparisonFile1Path))
            {
                _loggingService.Log("Error: Please select the first comparison file.", LogLevel.Error);
                StatusMessage = "Comparison failed";
                UpdateLogDisplay();
                return;
            }

            if (string.IsNullOrWhiteSpace(ComparisonFile2Path))
            {
                _loggingService.Log("Error: Please select the second comparison file.", LogLevel.Error);
                StatusMessage = "Comparison failed";
                UpdateLogDisplay();
                return;
            }

            if (!File.Exists(ComparisonFile1Path))
            {
                _loggingService.Log($"Error: File 1 does not exist: {ComparisonFile1Path}", LogLevel.Error);
                StatusMessage = "Comparison failed";
                UpdateLogDisplay();
                return;
            }

            if (!File.Exists(ComparisonFile2Path))
            {
                _loggingService.Log($"Error: File 2 does not exist: {ComparisonFile2Path}", LogLevel.Error);
                StatusMessage = "Comparison failed";
                UpdateLogDisplay();
                return;
            }

            var startTime = DateTime.Now;
            _loggingService.Log("=== File Comparison Started ===");
            _loggingService.Log($"File 1 (Manual): {ComparisonFile1Path}");
            _loggingService.Log($"File 2 (Program): {ComparisonFile2Path}");
            UpdateLogDisplay();

            // Run comparison
            await Task.Run(() =>
            {
                var result = _fileComparisonService.CompareColumnG(ComparisonFile1Path, ComparisonFile2Path, 3);

                _loggingService.Log($"=== Comparison Results ===");
                _loggingService.Log($"Total rows compared: {result.TotalRowsCompared}");
                _loggingService.Log($"Matching rows: {result.MatchingRows}");
                _loggingService.Log($"Mismatching rows: {result.MismatchingRows}");
                _loggingService.Log($"Missing in File 1: {result.MissingInFile1}");
                _loggingService.Log($"Missing in File 2: {result.MissingInFile2}");
                _loggingService.Log($"");
                
                if (result.IsIdentical)
                {
                    _loggingService.Log("✓ FILES ARE IDENTICAL - All values match!", LogLevel.Info);
                }
                else
                {
                    _loggingService.Log($"✗ FILES DIFFER - Found {result.Mismatches.Count} differences", LogLevel.Warning);
                    _loggingService.Log($"");
                    _loggingService.Log("Mismatches:");
                    
                    // Show first 20 mismatches
                    var mismatchesToShow = result.Mismatches.Take(20).ToList();
                    foreach (var mismatch in mismatchesToShow)
                    {
                        var file1Val = mismatch.File1Value ?? "(empty)";
                        var file2Val = mismatch.File2Value ?? "(empty)";
                        _loggingService.Log($"  Row {mismatch.Row}: File1='{file1Val}' | File2='{file2Val}'");
                    }
                    
                    if (result.Mismatches.Count > 20)
                    {
                        _loggingService.Log($"  ... and {result.Mismatches.Count - 20} more mismatches");
                    }
                }

                var duration = DateTime.Now - startTime;
                _loggingService.Log($"");
                _loggingService.Log($"=== Comparison completed in {duration.TotalSeconds:F2} seconds ===");
            });

            StatusMessage = "Comparison completed";
            UpdateLogDisplay();
        }
        catch (Exception ex)
        {
            var errorMsg = $"An error occurred during comparison: {ex.Message}";
            _loggingService.Log(errorMsg, LogLevel.Error);
            _loggingService.Log($"Stack trace: {ex.StackTrace}", LogLevel.Error);
            StatusMessage = "Comparison error";
            UpdateLogDisplay();

            MessageBox.Show(
                $"An error occurred while comparing files:\n\n{errorMsg}\n\nCheck the log for details.",
                "Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            IsProcessing = false;
        }
    }

    [RelayCommand]
    private async Task Run()
    {
        if (IsProcessing)
            return;

        IsProcessing = true;
        StatusMessage = "Processing...";
        _loggingService.Clear();
        LogMessages.Clear();

        try
        {
            var startTime = DateTime.Now;
            _loggingService.Log("=== EAN Matching Process Started ===");
            _loggingService.Log($"Main file: {MainFilePath}");
            _loggingService.Log($"Reference A: {ReferenceAPath}");
            _loggingService.Log($"Reference B: {ReferenceBPath}");
            UpdateLogDisplay();

            // Validate inputs
            var validationResult = ValidateInputs();
            if (!validationResult.IsValid)
            {
                foreach (var error in validationResult.Errors)
                {
                    _loggingService.Log(error, LogLevel.Error);
                    StatusMessage = "Validation failed";
                }
                UpdateLogDisplay();
                return;
            }

            // Process files
            await Task.Run(() => ProcessFiles());

            var duration = DateTime.Now - startTime;
            _loggingService.Log($"=== Process completed in {duration.TotalSeconds:F2} seconds ===");
            StatusMessage = "Completed successfully";
            UpdateLogDisplay();

            // Flush logs to file
            _loggingService.FlushToFile(MainFilePath);
        }
        catch (Exception ex)
        {
            var errorMsg = $"An error occurred: {ex.Message}";
            _loggingService.Log(errorMsg, LogLevel.Error);
            _loggingService.Log($"Stack trace: {ex.StackTrace}", LogLevel.Error);
            StatusMessage = "Error occurred";
            UpdateLogDisplay();

            MessageBox.Show(
                $"An error occurred while processing:\n\n{errorMsg}\n\nCheck the log for details.",
                "Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);

            // Still try to flush logs
            try
            {
                _loggingService.FlushToFile(MainFilePath);
            }
            catch { }
        }
        finally
        {
            IsProcessing = false;
        }
    }

    private ValidationResult ValidateInputs()
    {
        var result = new ValidationResult { IsValid = true };

        // Validate file paths
        if (string.IsNullOrWhiteSpace(MainFilePath))
        {
            result.IsValid = false;
            result.Errors.Add("Main file path is required.");
        }
        else if (!File.Exists(MainFilePath))
        {
            result.IsValid = false;
            result.Errors.Add($"Main file does not exist: {MainFilePath}");
        }
        else if (!MainFilePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            result.IsValid = false;
            result.Errors.Add("Main file must be an .xlsx file.");
        }

        if (string.IsNullOrWhiteSpace(ReferenceAPath))
        {
            result.IsValid = false;
            result.Errors.Add("Reference file A path is required.");
        }
        else if (!File.Exists(ReferenceAPath))
        {
            result.IsValid = false;
            result.Errors.Add($"Reference file A does not exist: {ReferenceAPath}");
        }
        else if (!ReferenceAPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            result.IsValid = false;
            result.Errors.Add("Reference file A must be an .xlsx file.");
        }

        if (string.IsNullOrWhiteSpace(ReferenceBPath))
        {
            result.IsValid = false;
            result.Errors.Add("Reference file B path is required.");
        }
        else if (!File.Exists(ReferenceBPath))
        {
            result.IsValid = false;
            result.Errors.Add($"Reference file B does not exist: {ReferenceBPath}");
        }
        else if (!ReferenceBPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            result.IsValid = false;
            result.Errors.Add("Reference file B must be an .xlsx file.");
        }

        return result;
    }

    private void ProcessFiles()
    {
        // Load main file
        _loggingService.Log("Loading main file...");
        UpdateLogDisplay();
        using var mainWorkbook = new XLWorkbook(MainFilePath);
        var mainWorksheet = mainWorkbook.Worksheet(1);
        if (mainWorksheet == null)
        {
            throw new InvalidOperationException("Main file does not have a first worksheet.");
        }

        // Read EANs from main file (starting at row 3)
        _loggingService.Log("Reading EANs from main file (column C, starting at row 3)...");
        _loggingService.Log("Note: Only numeric EANs (8-14 digits) will be read. Product names and other text will be skipped.");
        UpdateLogDisplay();
        var mainEans = _excelReader.ReadEansFromColumnC(mainWorksheet, 3);
        _loggingService.Log($"Found {mainEans.Count} valid EANs in main file.");
        
        if (mainEans.Count == 0)
        {
            _loggingService.Log("WARNING: No valid EANs found. Checking if EANs might be in a different column or starting row...", LogLevel.Warning);
            // Log what we found in the first few rows to help diagnose
            for (int row = 3; row <= Math.Min(10, mainWorksheet.LastRowUsed()?.RowNumber() ?? 10); row++)
            {
                var cell = mainWorksheet.Cell(row, 3);
                if (!cell.IsEmpty())
                {
                    var value = cell.GetString();
                    _loggingService.Log($"  Row {row}, Column C: '{value}' (Type: {cell.DataType})");
                }
            }
        }

        if (mainEans.Count == 0)
        {
            throw new InvalidOperationException("No EANs found in main file column C starting from row 3.");
        }

        // Read dropdown options from main file
        _loggingService.Log("Reading dropdown options from main file (column G)...");
        UpdateLogDisplay();
        
        var dropdownOptions = _excelReader.ReadDropdownOptions(mainWorksheet, 7, MainFilePath);
        
        if (dropdownOptions.Count > 0)
        {
            _loggingService.Log($"Found {dropdownOptions.Count} dropdown options: {string.Join(", ", dropdownOptions)}");
        }
        else
        {
            _loggingService.Log("No dropdown options found via data validation API.", LogLevel.Warning);
        }

        if (dropdownOptions.Count == 0)
        {
            // Try to provide more helpful error message
            var sampleCell = mainWorksheet.Cell(3, 7); // Check row 3, column G
            var hasValue = !sampleCell.IsEmpty();
            var sampleValue = hasValue ? sampleCell.GetString() : "(empty)";
            
            _loggingService.Log($"Column G sample (row 3): {sampleValue}", LogLevel.Warning);
            _loggingService.Log("Attempting to read all unique values from column G as fallback...", LogLevel.Warning);
            UpdateLogDisplay();
            
            // Last resort: read all unique non-empty values from column G
            var allValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var lastRow = mainWorksheet.LastRowUsed()?.RowNumber() ?? 1000;
            for (int row = 1; row <= Math.Min(1000, lastRow); row++)
            {
                var cell = mainWorksheet.Cell(row, 7);
                if (!cell.IsEmpty())
                {
                    var value = cell.GetString().Trim();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        allValues.Add(value);
                    }
                }
            }
            
            if (allValues.Count > 0)
            {
                dropdownOptions = allValues.OrderBy(v => v).ToList();
                _loggingService.Log($"Using {dropdownOptions.Count} unique values from column G as dropdown options: {string.Join(", ", dropdownOptions)}", LogLevel.Warning);
            }
            else
            {
                throw new InvalidOperationException(
                    "No dropdown options found in main file column G. " +
                    "Please ensure column G has data validation with dropdown options, " +
                    "or has at least some values that can be used as options.");
            }
        }

        // Load reference files
        _loggingService.Log("Loading reference file A...");
        UpdateLogDisplay();
        using var refAWorkbook = new XLWorkbook(ReferenceAPath);
        var refAWorksheet = refAWorkbook.Worksheet(1);
        if (refAWorksheet == null)
        {
            throw new InvalidOperationException("Reference file A does not have a first worksheet.");
        }

        _loggingService.Log("Loading reference file B...");
        UpdateLogDisplay();
        using var refBWorkbook = new XLWorkbook(ReferenceBPath);
        var refBWorksheet = refBWorkbook.Worksheet(1);
        if (refBWorksheet == null)
        {
            throw new InvalidOperationException("Reference file B does not have a first worksheet.");
        }

        // Read EANs from reference files
        _loggingService.Log("Reading EANs from reference file A (column C)...");
        UpdateLogDisplay();
        var refAEans = _excelReader.ReadAllEansFromColumnC(refAWorksheet);
        _loggingService.Log($"Found {refAEans.Count} EANs in reference file A.");

        _loggingService.Log("Reading EANs from reference file B (column C)...");
        UpdateLogDisplay();
        var refBEans = _excelReader.ReadAllEansFromColumnC(refBWorksheet);
        _loggingService.Log($"Found {refBEans.Count} EANs in reference file B.");

        // Map filenames to dropdown options
        _loggingService.Log("Mapping filenames to dropdown options...");
        UpdateLogDisplay();
        var mapping = _matchingService.MapFilenamesToDropdowns(ReferenceAPath, ReferenceBPath, dropdownOptions);
        
        if (!mapping.IsValid)
        {
            var errorMsg = "Failed to map filenames to dropdown options:\n" + string.Join("\n", mapping.ValidationErrors);
            throw new InvalidOperationException(errorMsg);
        }

        _loggingService.Log($"Mapped reference A to: {mapping.ReferenceAOption}");
        _loggingService.Log($"Mapped reference B to: {mapping.ReferenceBOption}");
        _loggingService.Log($"No-match option: {mapping.NoMatchOption}");
        UpdateLogDisplay();

        // Debug: Log sample EANs for comparison
        if (mainEans.Count > 0)
        {
            var sampleMainEans = mainEans.Take(5).Select(kvp => kvp.Value).ToList();
            _loggingService.Log($"Sample main file EANs (first 5): {string.Join(", ", sampleMainEans)}");
        }
        if (refAEans.Count > 0)
        {
            var sampleRefAEans = refAEans.Take(5).ToList();
            _loggingService.Log($"Sample reference A EANs (first 5): {string.Join(", ", sampleRefAEans)}");
        }
        if (refBEans.Count > 0)
        {
            var sampleRefBEans = refBEans.Take(5).ToList();
            _loggingService.Log($"Sample reference B EANs (first 5): {string.Join(", ", sampleRefBEans)}");
        }
        UpdateLogDisplay();

        // Match EANs
        _loggingService.Log("Matching EANs...");
        UpdateLogDisplay();
        var rowStatuses = _matchingService.MatchEans(mainEans, refAEans, refBEans, mapping);

        // Count matches
        var summary = new RunSummary
        {
            TotalEansProcessed = mainEans.Count
        };

        foreach (var status in rowStatuses.Values)
        {
            if (status == mapping.ReferenceAOption)
                summary.MatchesInReferenceA++;
            else if (status == mapping.ReferenceBOption)
                summary.MatchesInReferenceB++;
            else
                summary.NoMatches++;
        }

        _loggingService.Log($"Processing complete:");
        _loggingService.Log($"  Total EANs processed: {summary.TotalEansProcessed}");
        _loggingService.Log($"  Matches in Reference A: {summary.MatchesInReferenceA}");
        _loggingService.Log($"  Matches in Reference B: {summary.MatchesInReferenceB}");
        _loggingService.Log($"  No matches: {summary.NoMatches}");
        UpdateLogDisplay();

        // Write statuses to main file
        _loggingService.Log("Writing statuses to main file...");
        UpdateLogDisplay();
        _excelWriter.WriteStatuses(mainWorksheet, rowStatuses);

        // Save main file safely
        _loggingService.Log("Saving main file...");
        UpdateLogDisplay();
        _excelWriter.SafeOverwrite(mainWorkbook, MainFilePath);
        _loggingService.Log("Main file saved successfully.");
        UpdateLogDisplay();
    }

    private void UpdateLogDisplay()
    {
        Application.Current.Dispatcher.Invoke(() =>
        {
            var newLogs = _loggingService.GetAndClearLogs();
            foreach (var log in newLogs)
            {
                LogMessages.Add(log);
            }
        });
    }

    private class ValidationResult
    {
        public bool IsValid { get; set; }
        public List<string> Errors { get; set; } = new();
    }
}

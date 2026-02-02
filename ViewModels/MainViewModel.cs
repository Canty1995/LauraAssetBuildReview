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
    private readonly PowerPointReader _powerPointReader;
    private readonly ResultWriter _resultWriter;

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

    // Configuration properties
    [ObservableProperty]
    private ProcessingConfiguration _config;

    [ObservableProperty]
    private bool _useWorksheetIndex = true;

    [ObservableProperty]
    private bool _showLegacyReferenceFiles = false;

    [ObservableProperty]
    private string _expectedDropdownCountText = string.Empty;

    public ObservableCollection<string> LogMessages { get; } = new();
    public ObservableCollection<ReferenceFileViewModel> ReferenceFileConfigs { get; } = new();
    public ObservableCollection<ManualMappingViewModel> ManualMappings { get; } = new();

    private readonly ConfigurationService _configService;

    public MainViewModel()
    {
        _excelReader = new ExcelReader();
        _excelWriter = new ExcelWriter();
        _matchingService = new MatchingService();
        _loggingService = new LoggingService();
        _fileComparisonService = new FileComparisonService(_excelReader);
        _powerPointReader = new PowerPointReader();
        _resultWriter = new ResultWriter();
        _configService = new ConfigurationService();
        
        // Initialize with default configuration
        Config = _configService.GetDefaultConfiguration();
        
        // Initialize with 2 reference files for backward compatibility
        ReferenceFileConfigs.Add(new ReferenceFileViewModel("Reference File A", 1));
        ReferenceFileConfigs.Add(new ReferenceFileViewModel("Reference File B", 2));
        
        // Try to load saved configuration
        var savedConfig = _configService.LoadConfiguration();
        if (savedConfig != null)
        {
            Config = savedConfig;
            UpdateReferenceFilesFromConfig();
        }
    }

    [RelayCommand]
    private void BrowseMainFile()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "PowerPoint Files (*.pptx)|*.pptx|Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
            Title = "Select Main File (PowerPoint or Excel)"
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
                var result = _fileComparisonService.CompareColumn(ComparisonFile1Path, ComparisonFile2Path, 3, 7);

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

        // Validate main file path
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
        else if (!MainFilePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) && 
                 !MainFilePath.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
        {
            result.IsValid = false;
            result.Errors.Add("Main file must be an .xlsx or .pptx file.");
        }

        // Validate reference files - check both new system and legacy
        var hasValidReferenceFiles = false;
        
        // Check new flexible system
        foreach (var refFileVm in ReferenceFileConfigs)
        {
            var filePath = refFileVm.FilePath;
            if (string.IsNullOrWhiteSpace(filePath))
            {
                // Try legacy paths for backward compatibility
                if (refFileVm.Config.Priority == 1 && !string.IsNullOrWhiteSpace(ReferenceAPath))
                {
                    filePath = ReferenceAPath;
                }
                else if (refFileVm.Config.Priority == 2 && !string.IsNullOrWhiteSpace(ReferenceBPath))
                {
                    filePath = ReferenceBPath;
                }
            }

            if (!string.IsNullOrWhiteSpace(filePath))
            {
                if (!File.Exists(filePath))
                {
                    result.IsValid = false;
                    result.Errors.Add($"{refFileVm.DisplayName} does not exist: {filePath}");
                }
                else if (!filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) && 
                         !filePath.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
                {
                    result.IsValid = false;
                    result.Errors.Add($"{refFileVm.DisplayName} must be an .xlsx or .pptx file.");
                }
                else
                {
                    hasValidReferenceFiles = true;
                }
            }
        }

        // Legacy validation for backward compatibility (only if no valid files in new system)
        if (!hasValidReferenceFiles)
        {
            if (string.IsNullOrWhiteSpace(ReferenceAPath))
            {
                result.IsValid = false;
                result.Errors.Add("At least one reference file is required. Please configure reference files in the Configuration section.");
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
                result.Errors.Add("At least two reference files are required. Please configure reference files in the Configuration section.");
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
        }

        return result;
    }

    private void ProcessFiles()
    {
        // Sync configuration from UI
        SyncConfigFromUI();

        // Determine main file type
        var mainFileExtension = Path.GetExtension(MainFilePath).ToLowerInvariant();
        var isPowerPoint = mainFileExtension == ".pptx";

        if (isPowerPoint)
        {
            // PowerPoint workflow: Extract EANs and create new Excel result file
            ProcessPowerPointFile();
        }
        else
        {
            // Excel workflow: Original functionality - read dropdowns, match, write back to file
            ProcessExcelFile();
        }
    }

    private void ProcessPowerPointFile()
    {
        // Extract EANs from PowerPoint
        _loggingService.Log("Extracting EANs from PowerPoint file...");
        _loggingService.Log("Note: EANs must be exactly 14 digits.");
        UpdateLogDisplay();
        
        // Get selected slides from config (if any)
        var selectedSlides = new List<int>(); // For now, process all slides
        // TODO: Add UI to select slides for main PowerPoint file
        
        var mainEansResult = _powerPointReader.ReadEansFromPowerPointWithStats(
            MainFilePath,
            selectedSlides,
            Config.MinEanDigits,
            Config.MaxEanDigits,
            Config.AllowNonNumericEans);
        
        var mainEansList = mainEansResult.Eans.ToList();
        _loggingService.Log($"Found {mainEansList.Count} unique EANs in PowerPoint file.");
        
        // Log per-slide EAN counts
        _loggingService.Log("EAN counts per slide:");
        foreach (var kvp in mainEansResult.EanCountsPerSlide.OrderBy(x => x.Key))
        {
            _loggingService.Log($"  Slide {kvp.Key}: {kvp.Value} EAN(s)");
        }
        UpdateLogDisplay();

        if (mainEansList.Count == 0)
        {
            throw new InvalidOperationException("No EANs found in PowerPoint file.");
        }

        // Load reference files and create mapping of EAN to file names
        var eanToFiles = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        
        // Initialize all EANs with empty list
        foreach (var ean in mainEansList)
        {
            eanToFiles[ean] = new List<string>();
        }

        // Process reference files from configuration
        var referenceFiles = GetReferenceFilesFromConfig();

        foreach (var (refFilePath, displayName) in referenceFiles)
        {
            var fileName = Path.GetFileNameWithoutExtension(refFilePath);
            _loggingService.Log($"Loading {displayName}: {fileName}...");
            UpdateLogDisplay();

            var extension = Path.GetExtension(refFilePath).ToLowerInvariant();
            HashSet<string> refEans;

            if (extension == ".pptx")
            {
                // PowerPoint reference file
                var refFileVm = ReferenceFileConfigs.FirstOrDefault(r => r.FilePath == refFilePath);
                var refSelectedSlides = refFileVm?.Config.SelectedSlides ?? new List<int>();
                
                var refEansResult = _powerPointReader.ReadEansFromPowerPointWithStats(
                    refFilePath,
                    refSelectedSlides,
                    Config.MinEanDigits,
                    Config.MaxEanDigits,
                    Config.AllowNonNumericEans);
                
                refEans = refEansResult.Eans;
                
                // Log per-slide counts for reference files
                _loggingService.Log($"EAN counts per slide in {fileName}:");
                foreach (var kvp in refEansResult.EanCountsPerSlide.OrderBy(x => x.Key))
                {
                    _loggingService.Log($"  Slide {kvp.Key}: {kvp.Value} EAN(s)");
                }
                UpdateLogDisplay();
            }
            else
            {
                // Excel reference file
                using var refWorkbook = new XLWorkbook(refFilePath);
                var refWorksheet = refWorkbook.Worksheet(1);
                if (refWorksheet == null)
                {
                    _loggingService.Log($"WARNING: {fileName} does not have a first worksheet. Skipping.", LogLevel.Warning);
                    continue;
                }

                var refFileVm = ReferenceFileConfigs.FirstOrDefault(r => r.FilePath == refFilePath);
                var eanColumn = refFileVm != null 
                    ? ProcessingConfiguration.ColumnLetterToNumber(refFileVm.Config.EanColumn)
                    : 3;
                var startRow = refFileVm?.Config.StartRow ?? 1;

                refEans = _excelReader.ReadAllEansFromColumn(
                    refWorksheet, 
                    eanColumn, 
                    startRow, 
                    Config.MinEanDigits, 
                    Config.MaxEanDigits, 
                    Config.AllowNonNumericEans);
            }

            _loggingService.Log($"Found {refEans.Count} EANs in {fileName}.");

            // Match EANs and add file name to the list
            foreach (var ean in mainEansList)
            {
                if (refEans.Contains(ean))
                {
                    if (!eanToFiles[ean].Contains(fileName))
                    {
                        eanToFiles[ean].Add(fileName);
                    }
                }
            }
        }

        // Create result Excel file
        _loggingService.Log("Creating result Excel file...");
        UpdateLogDisplay();
        
        var outputFileName = Path.GetFileNameWithoutExtension(MainFilePath);
        var resultPath = _resultWriter.CreateResultFile(
            mainEansList, 
            eanToFiles, 
            outputFileName, 
            mainEansResult.EanCountsPerSlide,
            mainEansResult.EansPerSlide);
        
        _loggingService.Log($"Result file created: {resultPath}");
        
        // Log summary
        var foundCount = eanToFiles.Values.Count(files => files.Count > 0);
        var notFoundCount = mainEansList.Count - foundCount;
        
        _loggingService.Log($"Processing complete:");
        _loggingService.Log($"  Total EANs processed: {mainEansList.Count}");
        _loggingService.Log($"  EANs found in reference files: {foundCount}");
        _loggingService.Log($"  EANs not found: {notFoundCount}");
        UpdateLogDisplay();
    }

    private void ProcessExcelFile()
    {
        // Load main file
        _loggingService.Log("Loading main file...");
        UpdateLogDisplay();
        using var mainWorkbook = new XLWorkbook(MainFilePath);
        
        // Get worksheet based on configuration
        IXLWorksheet mainWorksheet;
        if (!string.IsNullOrWhiteSpace(Config.WorksheetName))
        {
            mainWorksheet = mainWorkbook.Worksheet(Config.WorksheetName);
            if (mainWorksheet == null)
            {
                throw new InvalidOperationException($"Main file does not have a worksheet named '{Config.WorksheetName}'.");
            }
        }
        else
        {
            mainWorksheet = mainWorkbook.Worksheet(Config.WorksheetIndex);
            if (mainWorksheet == null)
            {
                throw new InvalidOperationException($"Main file does not have a worksheet at index {Config.WorksheetIndex}.");
            }
        }

        // Read EANs from main file using configuration
        var eanColumn = Config.GetEanColumnNumber();
        _loggingService.Log($"Reading EANs from main file (column {Config.EanColumn}, starting at row {Config.StartRow})...");
        _loggingService.Log($"EAN validation: {Config.MinEanDigits}-{Config.MaxEanDigits} digits, Allow non-numeric: {Config.AllowNonNumericEans}");
        UpdateLogDisplay();
        var mainEans = _excelReader.ReadEansFromColumn(
            mainWorksheet, 
            eanColumn, 
            Config.StartRow, 
            Config.MinEanDigits, 
            Config.MaxEanDigits, 
            Config.AllowNonNumericEans);
        _loggingService.Log($"Found {mainEans.Count} valid EANs in main file.");
        
        if (mainEans.Count == 0)
        {
            _loggingService.Log($"WARNING: No valid EANs found. Checking if EANs might be in a different column or starting row...", LogLevel.Warning);
            // Log what we found in the first few rows to help diagnose
            for (int row = Config.StartRow; row <= Math.Min(Config.StartRow + 7, mainWorksheet.LastRowUsed()?.RowNumber() ?? Config.StartRow + 7); row++)
            {
                var cell = mainWorksheet.Cell(row, eanColumn);
                if (!cell.IsEmpty())
                {
                    var value = cell.GetString();
                    _loggingService.Log($"  Row {row}, Column {Config.EanColumn}: '{value}' (Type: {cell.DataType})");
                }
            }
        }

        if (mainEans.Count == 0)
        {
            throw new InvalidOperationException($"No EANs found in main file column {Config.EanColumn} starting from row {Config.StartRow}.");
        }

        // Read dropdown options from main file
        var dropdownColumn = Config.GetDropdownColumnNumber();
        _loggingService.Log($"Reading dropdown options from main file (column {Config.DropdownColumn})...");
        UpdateLogDisplay();
        
        var dropdownOptions = _excelReader.ReadDropdownOptions(mainWorksheet, dropdownColumn, MainFilePath);
        
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
            var sampleCell = mainWorksheet.Cell(Config.StartRow, dropdownColumn);
            var hasValue = !sampleCell.IsEmpty();
            var sampleValue = hasValue ? sampleCell.GetString() : "(empty)";
            
            _loggingService.Log($"Column {Config.DropdownColumn} sample (row {Config.StartRow}): {sampleValue}", LogLevel.Warning);
            _loggingService.Log("Attempting to read all unique values from dropdown column as fallback...", LogLevel.Warning);
            UpdateLogDisplay();
            
            // Last resort: read all unique non-empty values from dropdown column
            var allValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var lastRow = mainWorksheet.LastRowUsed()?.RowNumber() ?? 1000;
            for (int row = 1; row <= Math.Min(1000, lastRow); row++)
            {
                var cell = mainWorksheet.Cell(row, dropdownColumn);
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
                _loggingService.Log($"Using {dropdownOptions.Count} unique values from column {Config.DropdownColumn} as dropdown options: {string.Join(", ", dropdownOptions)}", LogLevel.Warning);
            }
            else
            {
                throw new InvalidOperationException(
                    $"No dropdown options found in main file column {Config.DropdownColumn}. " +
                    "Please ensure the dropdown column has data validation with dropdown options, " +
                    "or has at least some values that can be used as options.");
            }
        }

        // Validate dropdown count if specified
        if (Config.ExpectedDropdownCount.HasValue && dropdownOptions.Count != Config.ExpectedDropdownCount.Value)
        {
            _loggingService.Log($"WARNING: Expected {Config.ExpectedDropdownCount.Value} dropdown options, but found {dropdownOptions.Count}.", LogLevel.Warning);
        }

        // Load reference files using configuration
        var referenceFileMatches = new List<(HashSet<string> EanSet, string DropdownOption, int Priority)>();
        var referenceFilePaths = new List<string>();

        foreach (var refFileVm in ReferenceFileConfigs.OrderBy(r => r.Config.Priority))
        {
            if (string.IsNullOrWhiteSpace(refFileVm.FilePath))
            {
                // Try legacy paths for backward compatibility
                if (refFileVm.Config.Priority == 1 && !string.IsNullOrWhiteSpace(ReferenceAPath))
                {
                    refFileVm.FilePath = ReferenceAPath;
                    refFileVm.Config.FilePath = ReferenceAPath;
                }
                else if (refFileVm.Config.Priority == 2 && !string.IsNullOrWhiteSpace(ReferenceBPath))
                {
                    refFileVm.FilePath = ReferenceBPath;
                    refFileVm.Config.FilePath = ReferenceBPath;
                }
                else
                {
                    _loggingService.Log($"WARNING: Reference file {refFileVm.DisplayName} has no file path. Skipping.", LogLevel.Warning);
                    continue;
                }
            }

            if (!File.Exists(refFileVm.FilePath))
            {
                _loggingService.Log($"WARNING: Reference file {refFileVm.DisplayName} does not exist: {refFileVm.FilePath}. Skipping.", LogLevel.Warning);
                continue;
            }

            _loggingService.Log($"Loading {refFileVm.DisplayName}...");
            UpdateLogDisplay();

            HashSet<string> refEans;

            // Check file type and process accordingly
            var fileType = refFileVm.Config.FileType ?? "Excel";
            var extension = Path.GetExtension(refFileVm.FilePath).ToLowerInvariant();
            
            if (extension == ".pptx" || fileType == "PowerPoint")
            {
                // Process PowerPoint file
                var selectedSlides = refFileVm.Config.SelectedSlides ?? new List<int>();
                if (selectedSlides.Count == 0)
                {
                    _loggingService.Log($"Reading EANs from {refFileVm.DisplayName} (all slides)...");
                }
                else
                {
                    _loggingService.Log($"Reading EANs from {refFileVm.DisplayName} (slides: {string.Join(", ", selectedSlides)})...");
                }
                UpdateLogDisplay();

                var refEansResult = _powerPointReader.ReadEansFromPowerPointWithStats(
                    refFileVm.FilePath,
                    selectedSlides,
                    Config.MinEanDigits,
                    Config.MaxEanDigits,
                    Config.AllowNonNumericEans);

                refEans = refEansResult.Eans;
                _loggingService.Log($"Found {refEans.Count} EANs in {refFileVm.DisplayName}.");
                
                // Log per-slide counts for reference files
                _loggingService.Log($"EAN counts per slide in {refFileVm.DisplayName}:");
                foreach (var kvp in refEansResult.EanCountsPerSlide.OrderBy(x => x.Key))
                {
                    _loggingService.Log($"  Slide {kvp.Key}: {kvp.Value} EAN(s)");
                }
                UpdateLogDisplay();
            }
            else
            {
                // Process Excel file
                using var refWorkbook = new XLWorkbook(refFileVm.FilePath);
                IXLWorksheet refWorksheet;
                
                if (!string.IsNullOrWhiteSpace(refFileVm.Config.WorksheetName))
                {
                    refWorksheet = refWorkbook.Worksheet(refFileVm.Config.WorksheetName);
                    if (refWorksheet == null)
                    {
                        throw new InvalidOperationException($"{refFileVm.DisplayName} does not have a worksheet named '{refFileVm.Config.WorksheetName}'.");
                    }
                }
                else
                {
                    refWorksheet = refWorkbook.Worksheet(refFileVm.Config.WorksheetIndex);
                    if (refWorksheet == null)
                    {
                        throw new InvalidOperationException($"{refFileVm.DisplayName} does not have a worksheet at index {refFileVm.Config.WorksheetIndex}.");
                    }
                }

                var refEanColumn = ProcessingConfiguration.ColumnLetterToNumber(refFileVm.Config.EanColumn);
                _loggingService.Log($"Reading EANs from {refFileVm.DisplayName} (column {refFileVm.Config.EanColumn}, starting at row {refFileVm.Config.StartRow})...");
                UpdateLogDisplay();
                
                refEans = _excelReader.ReadAllEansFromColumn(
                    refWorksheet, 
                    refEanColumn, 
                    refFileVm.Config.StartRow, 
                    Config.MinEanDigits, 
                    Config.MaxEanDigits, 
                    Config.AllowNonNumericEans);
                
                _loggingService.Log($"Found {refEans.Count} EANs in {refFileVm.DisplayName}.");
            }
            referenceFilePaths.Add(refFileVm.FilePath);
            
            // Determine dropdown option for this reference file
            string dropdownOption = refFileVm.Config.MappedDropdownOption ?? string.Empty;
            
            if (string.IsNullOrWhiteSpace(dropdownOption) && Config.AutoMapFilenames)
            {
                // Try to auto-map
                var autoMappings = _matchingService.MapFilenamesToDropdownsFlexible(
                    new List<string> { refFileVm.FilePath },
                    dropdownOptions,
                    Config.ManualDropdownMappings);
                
                if (autoMappings.TryGetValue(refFileVm.FilePath, out var mapped))
                {
                    dropdownOption = mapped;
                }
            }
            else if (Config.ManualDropdownMappings.TryGetValue(refFileVm.FilePath, out var manualMapped))
            {
                dropdownOption = manualMapped;
            }

            if (string.IsNullOrWhiteSpace(dropdownOption))
            {
                _loggingService.Log($"WARNING: Could not determine dropdown option for {refFileVm.DisplayName}. It will be skipped.", LogLevel.Warning);
                continue;
            }

            _loggingService.Log($"Mapped {refFileVm.DisplayName} to dropdown option: {dropdownOption}");
            referenceFileMatches.Add((refEans, dropdownOption, refFileVm.Config.Priority));
        }

        if (referenceFileMatches.Count == 0)
        {
            throw new InvalidOperationException("No valid reference files found. Please configure at least one reference file.");
        }

        // Determine no-match option
        var usedOptions = referenceFileMatches.Select(m => m.DropdownOption).ToList();
        var noMatchOption = dropdownOptions.FirstOrDefault(o => !usedOptions.Contains(o, StringComparer.OrdinalIgnoreCase));
        
        if (string.IsNullOrWhiteSpace(noMatchOption) && dropdownOptions.Count > referenceFileMatches.Count)
        {
            // Use the first unused option
            noMatchOption = dropdownOptions.FirstOrDefault(o => !usedOptions.Any(u => u.Equals(o, StringComparison.OrdinalIgnoreCase)));
        }
        
        if (string.IsNullOrWhiteSpace(noMatchOption))
        {
            _loggingService.Log("WARNING: Could not determine no-match option. Using first dropdown option.", LogLevel.Warning);
            noMatchOption = dropdownOptions.FirstOrDefault() ?? "No Match";
        }

        _loggingService.Log($"No-match option: {noMatchOption}");
        UpdateLogDisplay();

        // Debug: Log sample EANs for comparison
        if (mainEans.Count > 0)
        {
            var sampleMainEans = mainEans.Take(5).Select(kvp => kvp.Value).ToList();
            _loggingService.Log($"Sample main file EANs (first 5): {string.Join(", ", sampleMainEans)}");
        }
        foreach (var (eanSet, dropdownOption, priority) in referenceFileMatches)
        {
            if (eanSet.Count > 0)
            {
                var sampleEans = eanSet.Take(5).ToList();
                _loggingService.Log($"Sample reference file (priority {priority}) EANs (first 5): {string.Join(", ", sampleEans)}");
            }
        }
        UpdateLogDisplay();

        // Match EANs using flexible matching
        _loggingService.Log("Matching EANs...");
        UpdateLogDisplay();
        var rowStatuses = _matchingService.MatchEansFlexible(mainEans, referenceFileMatches, noMatchOption);

        // Count matches
        var summary = new RunSummary
        {
            TotalEansProcessed = mainEans.Count
        };

        foreach (var status in rowStatuses.Values)
        {
            if (usedOptions.Contains(status, StringComparer.OrdinalIgnoreCase))
            {
                // Find which reference file this matches
                var matchIndex = usedOptions.FindIndex(o => o.Equals(status, StringComparison.OrdinalIgnoreCase));
                if (matchIndex == 0)
                    summary.MatchesInReferenceA++;
                else if (matchIndex == 1)
                    summary.MatchesInReferenceB++;
            }
            else
            {
                summary.NoMatches++;
            }
        }

        _loggingService.Log($"Processing complete:");
        _loggingService.Log($"  Total EANs processed: {summary.TotalEansProcessed}");
        _loggingService.Log($"  Matches in reference files: {summary.MatchesInReferenceA + summary.MatchesInReferenceB}");
        _loggingService.Log($"  No matches: {summary.NoMatches}");
        UpdateLogDisplay();

        // Write statuses to main file
        var statusColumn = Config.GetStatusColumnNumber();
        _loggingService.Log($"Writing statuses to main file (column {Config.StatusColumn})...");
        UpdateLogDisplay();
        _excelWriter.WriteStatuses(mainWorksheet, rowStatuses, statusColumn);

        // Save main file safely
        _loggingService.Log("Saving main file...");
        UpdateLogDisplay();
        var processedFilePath = _excelWriter.SaveToProcessedFolder(mainWorkbook, MainFilePath);
        _loggingService.Log($"Processed file saved to: {processedFilePath}");
        _loggingService.Log($"Original file moved to: original folder");
        _loggingService.Log("File processing completed successfully.");
        UpdateLogDisplay();
    }

    private List<(string FilePath, string DisplayName)> GetReferenceFilesFromConfig()
    {
        var referenceFiles = new List<(string FilePath, string DisplayName)>();
        
        // Get reference files from the configuration collection
        foreach (var refFileVm in ReferenceFileConfigs.OrderBy(r => r.Config.Priority))
        {
            var filePath = refFileVm.FilePath;
            
            // Try legacy paths for backward compatibility if file path is empty
            if (string.IsNullOrWhiteSpace(filePath))
            {
                if (refFileVm.Config.Priority == 1 && !string.IsNullOrWhiteSpace(ReferenceAPath))
                {
                    filePath = ReferenceAPath;
                }
                else if (refFileVm.Config.Priority == 2 && !string.IsNullOrWhiteSpace(ReferenceBPath))
                {
                    filePath = ReferenceBPath;
                }
            }

            if (!string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath))
            {
                referenceFiles.Add((filePath, refFileVm.DisplayName));
            }
        }

        // Fallback to legacy reference files if no files in collection
        if (referenceFiles.Count == 0)
        {
            if (!string.IsNullOrWhiteSpace(ReferenceAPath) && File.Exists(ReferenceAPath))
            {
                referenceFiles.Add((ReferenceAPath, "Reference File A"));
            }
            if (!string.IsNullOrWhiteSpace(ReferenceBPath) && File.Exists(ReferenceBPath))
            {
                referenceFiles.Add((ReferenceBPath, "Reference File B"));
            }
        }

        return referenceFiles;
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

    // Configuration management methods
    [RelayCommand]
    private void AddReferenceFile()
    {
        var priority = ReferenceFileConfigs.Count + 1;
        var newFile = new ReferenceFileViewModel($"Reference File {GetNextReferenceFileLetter()}", priority);
        ReferenceFileConfigs.Add(newFile);
    }

    public void RemoveReferenceFile(ReferenceFileViewModel viewModel)
    {
        ReferenceFileConfigs.Remove(viewModel);
        // Renumber priorities
        for (int i = 0; i < ReferenceFileConfigs.Count; i++)
        {
            ReferenceFileConfigs[i].Config.Priority = i + 1;
        }
    }

    private string GetNextReferenceFileLetter()
    {
        if (ReferenceFileConfigs.Count == 0) return "A";
        var lastChar = (char)('A' + ReferenceFileConfigs.Count);
        return lastChar.ToString();
    }

    [RelayCommand]
    private void SaveConfig()
    {
        try
        {
            SyncConfigFromUI();
            _configService.SaveConfiguration(Config);
            _loggingService.Log("Configuration saved successfully.", LogLevel.Info);
            UpdateLogDisplay();
            MessageBox.Show("Configuration saved successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            _loggingService.Log($"Failed to save configuration: {ex.Message}", LogLevel.Error);
            UpdateLogDisplay();
            MessageBox.Show($"Failed to save configuration:\n\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private void LoadConfig()
    {
        try
        {
            var loadedConfig = _configService.LoadConfiguration();
            if (loadedConfig != null)
            {
                Config = loadedConfig;
                UpdateReferenceFilesFromConfig();
                UpdateUIFromConfig();
                _loggingService.Log("Configuration loaded successfully.", LogLevel.Info);
                UpdateLogDisplay();
                MessageBox.Show("Configuration loaded successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("No saved configuration found.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            _loggingService.Log($"Failed to load configuration: {ex.Message}", LogLevel.Error);
            UpdateLogDisplay();
            MessageBox.Show($"Failed to load configuration:\n\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private void ResetConfig()
    {
        var result = MessageBox.Show("Reset configuration to defaults? This will clear all current settings.", 
            "Confirm Reset", MessageBoxButton.YesNo, MessageBoxImage.Question);
        
        if (result == MessageBoxResult.Yes)
        {
            Config = _configService.GetDefaultConfiguration();
            ReferenceFileConfigs.Clear();
            ReferenceFileConfigs.Add(new ReferenceFileViewModel("Reference File A", 1));
            ReferenceFileConfigs.Add(new ReferenceFileViewModel("Reference File B", 2));
            ManualMappings.Clear();
            UpdateUIFromConfig();
            _loggingService.Log("Configuration reset to defaults.", LogLevel.Info);
            UpdateLogDisplay();
        }
    }

    private void SyncConfigFromUI()
    {
        // Sync reference files
        Config.ReferenceFiles.Clear();
        foreach (var refFileVm in ReferenceFileConfigs)
        {
            refFileVm.Config.FilePath = refFileVm.FilePath;
            Config.ReferenceFiles.Add(refFileVm.Config);
        }

        // Sync manual mappings
        Config.ManualDropdownMappings.Clear();
        foreach (var mapping in ManualMappings)
        {
            if (!string.IsNullOrWhiteSpace(mapping.FilePath) && !string.IsNullOrWhiteSpace(mapping.DropdownOption))
            {
                Config.ManualDropdownMappings[mapping.FilePath] = mapping.DropdownOption;
            }
        }

        // Sync expected dropdown count
        if (int.TryParse(ExpectedDropdownCountText, out int count))
        {
            Config.ExpectedDropdownCount = count;
        }
        else
        {
            Config.ExpectedDropdownCount = null;
        }
    }

    private void UpdateReferenceFilesFromConfig()
    {
        ReferenceFileConfigs.Clear();
        if (Config.ReferenceFiles.Count > 0)
        {
            foreach (var refConfig in Config.ReferenceFiles.OrderBy(r => r.Priority))
            {
                var vm = new ReferenceFileViewModel($"Reference File {GetReferenceFileLetter(refConfig.Priority)}", refConfig.Priority)
                {
                    FilePath = refConfig.FilePath,
                    Config = refConfig
                };
                ReferenceFileConfigs.Add(vm);
            }
        }
        else
        {
            // Default to 2 reference files
            ReferenceFileConfigs.Add(new ReferenceFileViewModel("Reference File A", 1));
            ReferenceFileConfigs.Add(new ReferenceFileViewModel("Reference File B", 2));
        }
    }

    private void UpdateUIFromConfig()
    {
        ExpectedDropdownCountText = Config.ExpectedDropdownCount?.ToString() ?? string.Empty;
        
        ManualMappings.Clear();
        foreach (var mapping in Config.ManualDropdownMappings)
        {
            ManualMappings.Add(new ManualMappingViewModel
            {
                FilePath = mapping.Key,
                DropdownOption = mapping.Value
            });
        }
    }

    private string GetReferenceFileLetter(int priority)
    {
        if (priority <= 0) return "A";
        var letter = (char)('A' + (priority - 1));
        return letter.ToString();
    }
}

# Laura Asset Build Review - EAN Excel Matcher

A **rigid and specific** portable Windows WPF application that updates Excel files by matching EAN values from a main file against two reference files, setting dropdown status values in **Column G** based on matches.

## 🚀 Quick Start - How to Use

### Step 1: Prepare Your Files
- **Main File**: Must have EANs in **Column C starting at Row 3**, and a dropdown in **Column G with 3 options**
- **Reference File A**: EANs in **Column C** (any row)
- **Reference File B**: EANs in **Column C** (any row)

### Step 2: Run the Application
1. Double-click `LauraAssetBuildReview.exe`
2. Click **"Browse..."** and select your **Main File**
3. Click **"Browse..."** and select **Reference File A**
4. Click **"Browse..."** and select **Reference File B**
5. Click **"Run"**

### Step 3: Check Results
- The main file is **automatically updated** with status values in Column G
- Check the log panel for processing details
- A log file is saved next to your main file

**That's it!** The program matches EANs and fills in Column G automatically.

---

## ⚠️ IMPORTANT: Rigid Column Requirements

This application is **very specific** about Excel file structure. Files must match these exact requirements or the application will fail:

### Main Excel File Requirements

**CRITICAL - These requirements are non-negotiable:**

1. **Column C (EAN Column)**:
   - **Location**: Column C (the 3rd column)
   - **Start Row**: Row 3 (EANs must start at C3, NOT C1 or C2)
   - **Format**: EAN values must be numeric (8-14 digits)
   - **Validation**: Only numeric values between 8-14 digits are processed
   - **Empty Cells**: Empty cells in Column C are skipped
   - **Data Type**: Can be stored as text or number in Excel (both are handled)

2. **Column G (Status Column)**:
   - **Location**: Column G (the 7th column)
   - **Data Validation**: Must have a **pre-existing dropdown data validation list**
   - **Dropdown Source**: Must be a comma-separated list (e.g., `"Recieved - KING01042,Missing - Need to request,Recieved - KING01058"`)
   - **Number of Options**: Must have **exactly 3 dropdown options**
   - **Output**: Status values are written to Column G starting at Row 3 (matching each EAN row)
   - **Overwrite**: Existing values in Column G will be **overwritten**

3. **Worksheet**:
   - **Sheet Selection**: Only the **first worksheet** is used (Sheet1)
   - **Other Sheets**: All other worksheets are ignored

### Reference File A Requirements

**CRITICAL - These requirements are non-negotiable:**

1. **Column C (EAN Column)**:
   - **Location**: Column C (the 3rd column)
   - **Start Row**: Can start from Row 1 (scans entire used range)
   - **Format**: EAN values must be numeric (8-14 digits)
   - **Validation**: Only numeric values between 8-14 digits are processed
   - **Empty Cells**: Empty cells are skipped

2. **Worksheet**:
   - **Sheet Selection**: Only the **first worksheet** is used (Sheet1)
   - **Other Sheets**: All other worksheets are ignored

3. **Filename Mapping**:
   - The filename (without extension) is used to match against dropdown options
   - Example: If file is `"CGI BRIEF - KING01042.xlsx"`, it will match dropdown options containing `"KING01042"` or `"CGI BRIEF"`

### Reference File B Requirements

**CRITICAL - These requirements are non-negotiable:**

1. **Column C (EAN Column)**:
   - **Location**: Column C (the 3rd column)
   - **Start Row**: Can start from Row 1 (scans entire used range)
   - **Format**: EAN values must be numeric (8-14 digits)
   - **Validation**: Only numeric values between 8-14 digits are processed
   - **Empty Cells**: Empty cells are skipped

2. **Worksheet**:
   - **Sheet Selection**: Only the **first worksheet** is used (Sheet1)
   - **Other Sheets**: All other worksheets are ignored

3. **Filename Mapping**:
   - The filename (without extension) is used to match against dropdown options
   - Example: If file is `"KING01058 - Bathrooms.xlsx"`, it will match dropdown options containing `"KING01058"` or `"Bathrooms"`

## Column Summary

| File Type | Column C (EAN) | Column G (Status) | Worksheet |
|-----------|----------------|-------------------|-----------|
| **Main File** | EANs starting at **Row 3** | Dropdown output starting at **Row 3** | **First sheet only** |
| **Reference A** | EANs from **Row 1** onwards | Not used | **First sheet only** |
| **Reference B** | EANs from **Row 1** onwards | Not used | **First sheet only** |

## Features

- **EAN Matching**: Automatically matches EAN values from a main Excel file against two reference files
- **Smart Dropdown Mapping**: Intelligently maps reference filenames to dropdown options using case-insensitive matching
- **Safe File Writing**: Uses temporary files to prevent corruption during write operations
- **Comprehensive Logging**: Real-time UI logging and timestamped log files
- **File Comparison**: Compare Column G between two files to verify program output
- **Portable Executable**: Single-file, self-contained Windows executable (no installer required)

## Prerequisites

- **.NET 8.0 SDK** or later (for building from source)
- **Windows 10/11** (64-bit) (for running the executable)

## Building the Application

### Step 1: Install .NET SDK

Download and install the .NET 8.0 SDK from [Microsoft's official website](https://dotnet.microsoft.com/download/dotnet/8.0).

### Step 2: Build the Executable

Open a terminal/command prompt in the project directory and run:

```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

This command will:
- Build the application in Release mode
- Target Windows x64 platform
- Create a self-contained executable (includes .NET runtime)
- Package everything into a single file

### Step 3: Locate the Executable

After building, the executable will be located at:

```
bin\Release\net8.0-windows\win-x64\publish\LauraAssetBuildReview.exe
```

You can copy this single `.exe` file to any Windows machine and run it directly - no installation required!

## Usage

### Step 1: Prepare Your Excel Files

**Before running the application, ensure your files meet the rigid requirements:**

1. **Main File**:
   - Must have EANs in **Column C starting at Row 3**
   - Must have a dropdown data validation in **Column G with exactly 3 options**
   - The dropdown must be a comma-separated list source
   - Example dropdown options: `"Recieved - KING01042,Missing - Need to request,Recieved - KING01058"`

2. **Reference File A**:
   - Must have EANs in **Column C** (any row from 1 onwards)
   - Filename should contain text that matches one of the dropdown options

3. **Reference File B**:
   - Must have EANs in **Column C** (any row from 1 onwards)
   - Filename should contain text that matches one of the dropdown options

### Step 2: Launch the Application

Double-click `LauraAssetBuildReview.exe` to launch the application.

### Step 3: Select Files

1. Click **"Browse..."** next to **"Main File"** and select your main Excel file
2. Click **"Browse..."** next to **"Reference File A"** and select the first reference file
3. Click **"Browse..."** next to **"Reference File B"** and select the second reference file

### Step 4: Run the Process

1. Click the **"Run"** button
2. The application will:
   - Validate all files exist and are `.xlsx` format
   - Validate Column C has EAN data starting at Row 3 (main file)
   - Validate Column G has a dropdown with exactly 3 options (main file)
   - Read EANs from all files (Column C only)
   - Map reference filenames to dropdown options
   - Match EANs and set status values in Column G
   - Save the updated main file (overwrites original)

### Step 5: Review Results

1. Check the **log panel** in the UI for processing details:
   - Selected files
   - Detected dropdown options
   - Filename-to-dropdown mapping
   - EAN processing statistics
   - Warnings and errors

2. A **log file** will be saved next to your main Excel file with a timestamp:
   - Format: `[MainFileName]_EANMatchLog_yyyyMMdd_HHmmss.txt`
   - Example: `MyFile_EANMatchLog_20260202_123045.txt`

### Step 6: File Comparison (Optional)

To verify the program output:

1. Click **"Browse..."** next to **"File 1 (Manual)"** and select your manually populated file
2. Click **"Browse..."** next to **"File 2 (Program)"** and select the program-generated file
3. Click **"Compare Files"** to compare Column G values between the two files
4. Review the comparison results in the log panel:
   - Total rows compared
   - Matching rows
   - Mismatching rows
   - Missing values

## Matching Logic

For each EAN in the main file (Column C, starting at Row 3):

1. **If found in Reference File A (Column C)** → Sets Column G to the dropdown option matching Reference A's filename
2. **Else if found in Reference File B (Column C)** → Sets Column G to the dropdown option matching Reference B's filename
3. **Else** → Sets Column G to the remaining dropdown option (no-match option)

**Priority**: Reference A takes precedence if an EAN appears in both reference files.

**EAN Normalization**:
- EANs are treated as text to preserve leading zeros
- Whitespace is trimmed
- Common formatting characters (dashes, spaces, dots) are removed
- Only numeric values with 8-14 digits are considered valid EANs

## Dropdown Mapping

The application automatically maps reference filenames to dropdown options using a three-pass strategy:

1. **Exact Match** (case-insensitive): Filename exactly matches a dropdown option
2. **Containment Match**: Filename contains dropdown option text or vice versa
3. **Identifier Match**: Extracts alphanumeric codes (e.g., "KING01058") from both filename and dropdown options and matches them

The remaining dropdown option (not matched to either filename) is used for "no match" cases.

**Example**:
- Filename: `"CGI BRIEF - KING01042.xlsx"`
- Dropdown options: `"Recieved - KING01042"`, `"Missing - Need to request"`, `"Recieved - KING01058"`
- Mapping: `"Recieved - KING01042"` matches the filename (contains "KING01042")

## Log Files

Log files are automatically created next to the main Excel file with the naming pattern:

```
[MainFileName]_EANMatchLog_yyyyMMdd_HHmmss.txt
```

Example: `MyFile_EANMatchLog_20260202_123045.txt`

Log files contain:
- Selected file paths
- Detected dropdown options
- Filename-to-dropdown mapping decisions
- EAN processing statistics (total processed, matches in A/B, no matches)
- Warnings (unexpected dropdown structure, missing values, unreadable cells)
- Errors with stack traces and friendly messages
- Processing duration

## Project Structure

```
LauraAssetBuildReview/
├── Models/
│   ├── RunSummary.cs          # Processing statistics
│   └── MappingResult.cs       # Filename-to-dropdown mapping
├── Services/
│   ├── ExcelReader.cs         # Excel reading operations (Column C, dropdowns)
│   ├── ExcelWriter.cs         # Excel writing with safe overwrite (Column G)
│   ├── MatchingService.cs     # EAN matching and dropdown mapping
│   ├── FileComparisonService.cs # Column G comparison between files
│   └── LoggingService.cs      # UI and file logging
├── ViewModels/
│   └── MainViewModel.cs       # MVVM view model
├── MainWindow.xaml            # WPF UI definition
├── MainWindow.xaml.cs         # UI code-behind
└── README.md                  # This file
```

## Technical Details

- **Framework**: .NET 8.0
- **UI Framework**: WPF (Windows Presentation Foundation)
- **MVVM Pattern**: CommunityToolkit.Mvvm
- **Excel Library**: ClosedXML (with DocumentFormat.OpenXml for dropdown reading)
- **Architecture**: MVVM with service layer separation

## Error Handling

The application includes comprehensive error handling:

- **File Validation**: Checks file existence and `.xlsx` format before processing
- **Excel Structure Validation**: Verifies worksheets and columns exist
- **Column Validation**: Ensures Column C has data from Row 3 in main file
- **Dropdown Validation**: Ensures Column G has exactly 3 dropdown options and can be mapped
- **EAN Validation**: Only processes numeric values with 8-14 digits
- **Safe File Operations**: Uses temporary files to prevent corruption
- **User-Friendly Messages**: Clear error messages in the UI
- **Detailed Logging**: Stack traces and detailed errors in log files

## Performance

- Uses `HashSet` for O(1) EAN lookup performance
- Loads all reference EANs into memory before processing
- Single-pass iteration through main file rows
- Efficient Excel reading/writing with ClosedXML

## Troubleshooting

### "No dropdown options found"
- **Cause**: Column G does not have data validation with a dropdown list
- **Solution**: 
  - Ensure Column G has data validation
  - The dropdown must be a comma-separated list source
  - Try creating the dropdown in Excel if it doesn't exist
  - Check that the dropdown applies to Column G

### "Could not match reference file to dropdown option"
- **Cause**: Dropdown option names don't contain text from the reference filename
- **Solution**: 
  - Ensure dropdown option names contain or match the reference filename (without extension)
  - Dropdown options should be descriptive enough to match filenames
  - Check the log for specific mapping details
  - Example: If filename is `"KING01042.xlsx"`, dropdown should contain `"KING01042"` or similar

### "No EANs found in main file"
- **Cause**: Column C doesn't have valid EAN values starting at Row 3
- **Solution**:
  - Ensure EANs start at **Row 3** in Column C (NOT Row 1 or Row 2)
  - Ensure EANs are numeric (8-14 digits)
  - Check that Column C is the 3rd column (not Column A, B, D, etc.)

### "File is locked" or "Access denied"
- **Cause**: Excel file is open in another program
- **Solution**:
  - Close the Excel file in Excel before running the application
  - Ensure no other programs have the file open
  - Check file permissions
  - Ensure the file is not read-only

### "Application won't start"
- **Cause**: Missing .NET runtime or incompatible system
- **Solution**:
  - For self-contained builds, no runtime installation is needed
  - Ensure you're running Windows 10/11 (64-bit)
  - Check Windows Event Viewer for detailed error messages

### "Dropdown has wrong number of options"
- **Cause**: Column G dropdown doesn't have exactly 3 options
- **Solution**:
  - The application requires **exactly 3 dropdown options**
  - Update your dropdown to have exactly 3 options
  - Check the log to see how many options were detected

### Comparison shows mismatches due to quotes
- **Cause**: Extra quotes in program-generated values
- **Solution**:
  - This has been fixed in the latest version
  - Re-run the program to regenerate the file without quotes
  - The comparison normalizes quotes automatically

## Support

For issues or questions:
1. Check the log files for detailed error information
2. Verify your Excel files meet the rigid column requirements
3. Ensure Column C has EANs starting at Row 3 (main file)
4. Ensure Column G has exactly 3 dropdown options

## License

This project uses the following open-source libraries:

- **ClosedXML**: MIT License
- **CommunityToolkit.Mvvm**: MIT License

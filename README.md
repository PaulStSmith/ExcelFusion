# Excel Fusion

A command-line tool for extracting, manipulating, and rebuilding Microsoft Excel files with full VBA (Visual Basic for Applications) support. Excel Fusion treats Excel files as ZIP archives and provides comprehensive two-way functionality for file decomposition and reconstruction.

## Features

- **Complete Excel File Extraction**: Decomposes Excel files into individual XML components, worksheets, and resources
- **VBA Code Extraction**: Extracts Visual Basic for Applications source code, modules, classes, and forms
- **Reference Management**: Preserves and manages VBA project references with GUID validation
- **Full Reconstruction**: Rebuilds functional Excel files from extracted components
- **VBA Compilation**: Automatically compiles VBA projects during reconstruction
- **Flexible Output**: Supports various Excel formats (.xlsx, .xlsm, .xltx, etc.)
- **Cross-platform**: Built on .NET 8 for modern cross-platform compatibility

## Requirements

- **.NET 8 Runtime** or higher
- **Microsoft Excel** installed (required for VBA operations)
- **VBA Project Access** permissions (may require manual configuration)

## Installation

### From Source
```bash
git clone <repository-url>
cd ExcelFusion
dotnet build --configuration Release
```

### Usage
```shell
ExcelFusion <command> [options]
```

## Commands

### Extract Command
Extracts all components from an Excel file, including VBA source code and references.

```shell
ExcelFusion extract <excel-file> [--output <folder>]
ExcelFusion e <excel-file> [--o <folder>]
```

**Options:**
- `--output`, `--out`, `-o`: Specify output directory (defaults to Excel filename without extension)

**Example:**
```shell
ExcelFusion extract MyWorkbook.xlsm --output ./extracted
```

**Output Structure:**
```
extracted/
├── [Content_Types].xml          # Excel content type definitions
├── _rels/                       # Relationship mappings
├── xl/                          # Main Excel content
│   ├── workbook.xml            # Workbook structure
│   ├── worksheets/             # Individual worksheet data
│   │   ├── sheet1.xml
│   │   └── sheet2.xml
│   ├── sharedStrings.xml       # Shared string table
│   └── styles.xml              # Formatting and styles
└── .vba/                       # VBA components (if present)
    ├── Module1.bas             # VBA modules
    ├── Sheet1.cls              # Worksheet code-behind
    ├── ThisWorkbook.cls        # Workbook code-behind
    ├── UserForm1.frm           # User forms
    └── VBAProject.proj         # Reference metadata (JSON)
```

### Build Command
Reconstructs an Excel file from extracted components, including VBA code compilation.

```shell
ExcelFusion build <folder> [--output <filename>] [--extension <ext>]
ExcelFusion b <folder> [--o <filename>] [--e <ext>]
```

**Options:**
- `--output`, `--out`, `-o`: Specify output Excel filename (defaults to folder name)
- `--extension`, `--ext`, `-e`: Specify file extension (defaults to `.xlsx`)

**Example:**
```shell
ExcelFusion build ./extracted --output NewWorkbook.xlsm --extension xlsm
```

### License Command
Displays the MIT license information.

```shell
ExcelFusion license
ExcelFusion l
```

## VBA Integration

Excel Fusion provides comprehensive VBA support through Microsoft Office Interop:

### VBA Component Types
- **Standard Modules** (`.bas`): General VBA code modules
- **Class Modules** (`.cls`): Object-oriented VBA classes and worksheet code-behind
- **User Forms** (`.frm`): VBA user interface forms
- **Document Modules**: Workbook and worksheet-specific code

### Reference Management
- Automatically extracts VBA project references
- Preserves reference GUIDs and version information
- Validates reference integrity during reconstruction
- Supports both file-based and registered COM references

### VBA Compilation
- Automatically compiles VBA projects after code injection
- Provides detailed error reporting for compilation failures
- Handles VBA project access permission prompts

## Advanced Usage

### Batch Processing
```shell
# Extract multiple files
for file in *.xlsm; do ExcelFusion extract "$file"; done

# Rebuild multiple projects
for dir in */; do ExcelFusion build "$dir"; done
```

### Custom Extensions
```shell
# Create Excel template
ExcelFusion build myproject --extension xltx

# Create macro-enabled workbook
ExcelFusion build myproject --extension xlsm
```

## Command-Line Options

### Flexible Syntax
All commands and options are case-insensitive and support multiple formats:

```shell
# All equivalent extract commands
ExcelFusion extract file.xlsx --output folder
ExcelFusion EXTRACT file.xlsx --OUT folder
ExcelFusion e file.xlsx -o folder
ExcelFusion E file.xlsx /o folder
```

### Option Aliases
- **Output**: `--output`, `--out`, `-o`, `/o`
- **Extension**: `--extension`, `--ext`, `-e`, `/e`

## Error Handling

Excel Fusion provides comprehensive error handling for common scenarios:

- **File Not Found**: Clear error messages for missing Excel files or folders
- **VBA Access Denied**: Interactive prompts for granting VBA project access
- **COM Exceptions**: Detailed error reporting for Excel automation issues
- **Compilation Errors**: Specific feedback for VBA compilation failures

## Technical Implementation

- **ZIP Archive Manipulation**: Direct manipulation of Excel's underlying ZIP structure
- **COM Interop**: Integration with Excel Object Model for VBA operations
- **JSON Serialization**: Modern reference metadata storage
- **Stream Processing**: Efficient handling of large Excel files
- **Memory Management**: Proper cleanup of COM objects and file handles

## Troubleshooting

### VBA Access Issues
If you encounter "Programmatic access to Visual Basic Project is not trusted" errors:

1. Open Excel → File → Options → Trust Center → Trust Center Settings
2. Navigate to "Macro Settings"
3. Check "Trust access to the VBA project object model"
4. Restart Excel and try again

### Performance Considerations
- VBA operations require Excel automation and may be slower for large projects
- The `.vba` folder is automatically excluded from ZIP compression during build
- Large Excel files are processed using buffered streams for memory efficiency

## License

MIT License - Copyright © 2024 ByteForge

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
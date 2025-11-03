# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

ExcelFusion is a command-line tool for extracting and rebuilding Excel files with VBA components. It treats Excel files as ZIP archives and provides two-way functionality:
- **Extract**: Decomposes Excel files into constituent parts (XML files, VBA code, references)
- **Build**: Reconstructs Excel files from extracted components

## Build and Development Commands

```bash
# Build the project
dotnet build

# Run the application 
dotnet run -- <command> [options]

# Clean build artifacts
dotnet clean

# Restore NuGet packages
dotnet restore
```

## Application Usage Commands

```bash
# Extract Excel file components
ExcelFusion extract <file.xlsx> [--output <folder>]
# Short form: ExcelFusion e <file.xlsx> [--o <folder>]

# Build Excel file from components
ExcelFusion build <folder> [--output <filename>] [--extension <ext>]
# Short form: ExcelFusion b <folder> [--o <filename>] [--e <ext>]

# Show license information
ExcelFusion license
# Short form: ExcelFusion l
```

## Architecture Overview

### Core Components

- **Program.cs**: Entry point using System.CommandLine for CLI parsing with three main commands (extract, build, license)
- **ExcelFileCreator.cs**: Handles building Excel files from extracted components, including VBA injection via COM Interop
- **VbaExtractor.cs**: Extracts VBA source code and references using Excel COM Interop
- **ZipHelpers.cs**: Low-level ZIP operations for Excel file manipulation (Excel files are ZIP archives)

### Data Flow

1. **Extract Process**: Excel file → ZIP extraction → VBA component extraction → File system structure
2. **Build Process**: File system structure → ZIP compression → VBA component injection → Excel file

### Key Dependencies

- **Microsoft Office Interop**: COM references for Excel automation (Excel, VBIDE, Office.Core)
- **System.CommandLine**: Modern CLI framework for command parsing
- **.NET 8**: Target framework with nullable reference types enabled

### VBA Integration

- VBA components are stored in `.vba` subfolder during extraction
- Components are categorized by type (.bas, .cls, .frm extensions)
- References are serialized as JSON in `.proj` files
- Build process uses Excel COM Interop to inject VBA code and compile

### File Structure During Operations

```
extracted-folder/
├── [Content_Types].xml
├── _rels/
├── xl/
│   ├── worksheets/
│   ├── workbook.xml
│   └── ...
└── .vba/                    # VBA components (ignored during ZIP compression)
    ├── Module1.bas
    ├── Sheet1.cls
    ├── ThisWorkbook.cls
    └── ProjectName.proj     # References metadata as JSON
```

## Important Implementation Notes

- Excel COM Interop requires Excel to be installed and may show Excel application window
- VBA access permissions may need to be granted manually (handled via user prompt)
- ZIP operations preserve Excel's internal XML structure
- Error handling includes COM exception management for VBA operations
- AllowUnsafeBlocks is enabled for COM Interop operations
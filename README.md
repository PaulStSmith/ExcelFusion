# Excel Fusion
It's a simple command line tool to extract each component of 
an Excel file and save it as separate files, and also to 
reconstruct the Excel file from the components.

## Usage
``` shell
ExcelFusion <command> [options]
```

### Commands
Possible commands are:
- `e` or `extract` to extract the components of an Excel file
    
    Usage:

        ExcelFusion e <file> [--o <output folder>]

    Options:
    - `o`, `out`, or `output` to specify the output directory
 
        If the output directory is not specified, the components will be saved in a directory with the same name as the Excel file.

- `b` or `build` to build an Excel file from the components.
    
    Usage:

        ExcelFusion b <file> [--o <output folder> --e <extension>]

    Options:
    - `o`, `out`, or `output` to specify the name of the new Excel file.
 
        If the output file is not specified, the new Excel file will have the same name as the directory containing the components.

    - `e`, `ext`, or `extension` to specify the extension of the new Excel file.

        If the extension is not specified, the new Excel file will have the default extension `.xlsx`.

- `l` or `license` to show the license information. This command does not have any options.

    Usage:

        ExcelFusion l

**Note 01**: The commands are case-insensitive, so you can use `e`, `E`, `extract`, or `EXTRACT` interchangeably.

**Note 02**: The options are case-insensitive, so you can use `o`, `O`, `out`, `OUT`, `output`, or `OUTPUT` interchangeably, and can be called out using `/`, `-`, or `--`. Therefore, `/o`, `-o`, and `--o` mean the same option `Output`.

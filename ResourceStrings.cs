using System;

#nullable enable 

namespace ExcelFusion
{
    /// <summary>
    /// Auxiliary class that holds strings used in the program.
    /// </summary>
    internal static class ResourceStrings
    {
        public const String Compressing = "Compressing {0} ...";
        public const String CreateArgumentDescription = "A path to a folder containing the structure of an Excel file.";
        public const String CreateDescription = "Creates an Excel file based on a folder.";
        public const String CreateExtDescription = "The name of the Extension of the Excel file.";
        public const String CreateOutDescription = "The name of the Excel file output.";
        public const String Description = "Extracts files from an Excel file and/or integrates such files back into and Excel file.";
        public const String ExcelClosed = "Excel closed.";
        public const String ExcelClosing = "Closing Excel...";
        public const String ExcelOpen = "Excel opened.";
        public const String ExcelOpening = "Opening Excel...";
        public const String Extracing = "Extracting {0} to ‘{1}’ ...";
        public const String ExtractArgumentDescription = "A path to an Excel file.";
        public const String ExtractDescription = "Extracts files from an Excel file.";
        public const String ExtractOutDescription = "The folder where to extact the Excel file.";
        public const String FileNotFoundMessage = "Excel file '{0}' not found.";
        public const String FolderNotFoundMessage = "Folder '{0}' not found.";
        public const String GrantAccess = "Please, grant access to the Visual Basic Project.";
        public const String Header = @"  ___            _ ___        _           " + "\r\n"
                                   + @" | __|_ ____ ___| | __|  _ __(_)___ _ _   " + "\r\n"
                                   + @" | _|\ \ / _/ -_) | _| || (_-< / _ \ ' \  " + "\r\n"
                                   + @" |___/_\_\__\___|_|_| \_,_/__/_\___/_||_| " + "\r\n"
                                   + @"==========================================" + "\r\n"
                                   + @"Author: Paulo Santos " + "\r\n"
                                   + @"MIT License - 2024   " + "\r\n";
        public const String ItsA = "it's a {0}. Exporting.";
        public const String Open = "‘{0}’ Open.";
        public const String Opening = "Opening ‘{0}’...";
        public const String Processing = "Processing {0} ...";
        public const String ReferenceAdded = "Reference to {0} added to the project. GUID = ‘{1}’";
        public const String ReferenceFailed = "Failed to add reference to {0}.";
        public const String ReferenceGuidError = "GUID in reference to ‘{0}’ in the project file differs from GUID in the VB project.";
        public const String ReferenceGuidProject = "     in project file : {0}";
        public const String ReferenceGuiVbProject = "     in VB project   : {0}";
        public const String ReferenceRemoved = "Referece removed from VB proejct.";
        public const String TryAgain = "Try again? Y/N";
        public const String CouldNotDeserialize = "Could not deserialize file ‘{0}’";
        public const String OutputFolderNotSpecified = "Output folder not specified.";
        public const String RootCommandDescription = "Manipulates Excel files.";
        public const String VbaCompileError = "An error occurred while compiling the VBA project.";
        public const String MitLicense = "Copyright © 2024 Paulo Santos \r\n" +
                                         "\r\n" +
                                         "Permission is hereby granted, free of charge, to any person obtaining a copy \r\n" +
                                         "of this software and associated documentation files (the “Software”), to deal \r\n" +
                                         "in the Software without restriction, including without limitation the rights \r\n" +
                                         "to use, copy, modify, merge, publish, distribute, sublicense, and/or sell \r\n" +
                                         "copies of the Software, and to permit persons to whom the Software is \r\n" +
                                         "furnished to do so, subject to the following conditions: \r\n" +
                                         "\r\n"+
                                         "The above copyright notice and this permission notice shall be included in all \r\n" +
                                         "copies or substantial portions of the Software. \r\n" +
                                         "\r\n"+
                                         "THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR \r\n" +
                                         "IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, \r\n" +
                                         "FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE \r\n" +
                                         "AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER \r\n" +
                                         "LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, \r\n" +
                                         "OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE \r\n" +
                                         "SOFTWARE. \r\n";

        public const String LicenseDescription = "Displays the MIT license";
    }
}

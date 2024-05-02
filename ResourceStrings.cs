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
        public const String Extracing = "Extracting {0} ...";
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
    }
}

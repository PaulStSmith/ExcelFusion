#nullable enable 

namespace ExcelFusion
{
    /// <summary>
    /// Represents the options for the extract command.
    /// </summary>
    public class ExtractOptions
    {
        /// <summary>
        /// Gets or sets the path to an Excel file.
        /// </summary>
        public string? ExcelFile { get; set; }

        /// <summary>
        /// Gets or sets the output folder to extract the Excel file.
        /// </summary>
        public string? Out { get; set; }
    }
}

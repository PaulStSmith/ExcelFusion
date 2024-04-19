namespace AutomateExcel
{
    /// <summary>
    /// Represents the options for the extract command.
    /// </summary>
    public class CreateOptions
    {
        /// <summary>
        /// Gets or sets a path that contains the structure of an Excel file.
        /// </summary>
        public string? Folder { get; set; }

        /// <summary>
        /// Gets or sets the name of the Excel file to generate.
        /// </summary>
        public string? Out { get; set; }


        /// <summary>
        /// Gets or sets the name of the extension of the Excel file to generate.
        /// </summary>
        public string? Ext { get; set; }
    }
}

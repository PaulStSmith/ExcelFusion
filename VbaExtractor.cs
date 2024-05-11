using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace ExcelFusion
{
    /// <summary>
    /// Contains methods to extract Visual Basic source code from an Excel file.
    /// </summary>
    internal static class VbaExtractor
    {
        /// <summary>
        /// Options to serialize JSON.
        /// </summary>
        private static readonly JsonSerializerOptions jsonOpts = new() { WriteIndented = true };

        /// <summary>
        /// Extracts the VBA code from the Excel file specified within the <see cref="ExtractOptions"/> object.
        /// </summary>
        /// <param name="options">An <see cref="ExtractOptions"/> object containing the data to access the Excel file.</param>
        public static void ExtractVbaSourceCode(ExtractOptions options)
        {
            /*
             * If for some reason the file does not exist, display a message and exit.
             */
            if (!File.Exists(options.ExcelFile))
            {
                Console.WriteLine(ResourceStrings.FileNotFoundMessage, options.ExcelFile);
                return;
            }

            /*
             * Open Excel and the Excel file
             */
            Console.WriteLine(ResourceStrings.ExcelOpening);
            var xl = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = true,
                EnableEvents = false,
                DisplayAlerts = false,
                ScreenUpdating = false,
            };
            try
            {
                Console.WriteLine(ResourceStrings.ExcelOpen);
                Console.WriteLine(ResourceStrings.Opening, options.ExcelFile);
                var xlFilePath = (new FileInfo(options.ExcelFile)).FullName;
                var wb = xl.Workbooks.Open(xlFilePath);
                wb.Activate();
                Console.WriteLine(ResourceStrings.Open, options.ExcelFile);

                /*
                 * Check if we have a VB project to export.
                 */
                if (wb.HasVBProject)
                {
                    /*
                     * This while permits retry if we fail due to lack of permission.
                     */
                    while (true)
                    {
                        try
                        {
                            ExtractVbProject(options, wb);
                        }
                        catch (COMException ex)
                        {
                            if (!ProgramHelpers.HandleException(ex))
                                continue;
                        }
                        break;
                    }

                    ExtractReferences(options, wb);
                }

                wb.Close(SaveChanges: false);
            }
            finally
            {
                Console.WriteLine(ResourceStrings.ExcelClosing);
                xl.Quit();
                Console.WriteLine(ResourceStrings.ExcelClosed);
            }
        }

        /// <summary>
        /// Extracts all the references from the Visual Basic project.
        /// </summary>
        /// <param name="options">An <see cref="ExtractOptions"/> object containing the data to access the Excel file.</param>
        /// <p7aram name="wb">A <see cref="Workbook"/> object containing the Visual Basic project to extract.</param>
        private static void ExtractReferences(ExtractOptions options, Workbook wb)
        {
#pragma warning disable CS8604 // Possible null reference argument.
            if (!CheckArgs(options, wb) || !wb.HasVBProject) return;

            var dir = Path.Combine(options.Out, ".vba");
            var proj = wb.VBProject;
            var projFile = Path.Combine(dir, proj.Name + ".proj");
            var lst = new List<ReferenceInfo>();

            foreach (Reference rf in proj.References)
            {
                if (rf.IsBroken || rf.BuiltIn) continue;

                lst.Add(new ReferenceInfo()
                {
                    Guid = new System.Guid(rf.Guid),
                    Name = rf.Name,
                    FullPath = rf.FullPath,
                    Version = float.Parse($"{rf.Major}.{rf.Minor}"),
                    Type = rf.Type
                });
            }

            using var writer = new StreamWriter(projFile);
            writer.WriteLine(JsonSerializer.Serialize(lst, jsonOpts));
            writer.Close();
#pragma warning restore CS8604
        }

        /// <summary>
        /// Checks if the arguments are valid.
        /// </summary>
        /// <param name="options">An <see cref="ExtractOptions"/> object containing the data to access the Excel file.</param>
        /// <param name="wb">A <see cref="Workbook"/> object containing the Visual Basic project to extract.</param>
        /// <returns>True if the arguments are valid; otherwise, false.</returns>
        private static bool CheckArgs(ExtractOptions options, Workbook wb)
        {
            ArgumentNullException.ThrowIfNull(options);
            ArgumentNullException.ThrowIfNull(wb);
            if (string.IsNullOrEmpty(options.Out))
            {
                Console.WriteLine(ResourceStrings.OutputFolderNotSpecified);
                return false;
            }
            return true;
        }

        /// <summary>
        /// Exports the Visual Basic project contained in the specified <see cref="Workbook"/>.
        /// </summary>
        /// <param name="options">An <see cref="ExtractOptions"/> object containing the data to access the Excel file.</param>
        /// <p7aram name="wb">A <see cref="Workbook"/> object containing the Visual Basic project to extract.</param>
        public static void ExtractVbProject(ExtractOptions options, Workbook wb)
        {
#pragma warning disable CS8604 // Possible null reference argument.
            if (!CheckArgs(options, wb) || !wb.HasVBProject) return;

            var proj = wb.VBProject;
            foreach (VBComponent comp in proj.VBComponents)
            {
                /*
                 * Check if we need to ignore a component.
                 */
                Console.Write(ResourceStrings.Processing, $"{wb.Name}.{comp.Name}");
                /*
                 * Establish the file extension for the component to be exported.
                 */
                var ext = comp.Type switch
                {
                    vbext_ComponentType.vbext_ct_MSForm => ".frm",
                    vbext_ComponentType.vbext_ct_Document => ".cls",
                    vbext_ComponentType.vbext_ct_StdModule => ".bas",
                    vbext_ComponentType.vbext_ct_ClassModule => ".cls",
                    _ => ".bin",
                };

                /*
                 * Defines the name of the exported file.
                 * The VB components are exported to the “.vba” folder.
                 */
                var dir = Path.Combine(options.Out, ".vba");
                var filePath = Path.Combine(dir, comp.Name + ext);
                Console.WriteLine(ResourceStrings.ItsA, comp.Type.ToString()[(comp.Type.ToString().LastIndexOf('_') + 1)..]);
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                /*
                 * Extracts the component to the file.
                 */
                comp.Export(filePath);
            }
#pragma warning restore CS8604
        }
    }
}
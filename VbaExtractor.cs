using ExcelFusion.Properties;
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
                Console.WriteLine(Resources.FileNotFoundMessage, options.ExcelFile);
                return;
            }

            /*
             * Open Excel and the Excel file
             */
            Console.WriteLine(Resources.ExcelOpening);
            Microsoft.Office.Interop.Excel.Application? xl = null;
            Workbooks? workbooks = null;
            Workbook? wb = null;
            var workbookClosed = false;
            int? excelProcessId = null;
            try
            {
                xl = new Microsoft.Office.Interop.Excel.Application
                {
                    Visible = true,
                    EnableEvents = false,
                    DisplayAlerts = false,
                    ScreenUpdating = false,
                };
                excelProcessId = GetExcelProcessId(xl);
                Console.WriteLine(Resources.ExcelOpen);
                Console.WriteLine(Resources.Opening, options.ExcelFile);
                var xlFilePath = (new FileInfo(options.ExcelFile)).FullName;
                workbooks = xl.Workbooks;
                wb = workbooks.Open(xlFilePath);
                wb.Activate();
                Console.WriteLine(Resources.Open, options.ExcelFile);

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
                workbookClosed = true;
            }
            finally
            {
                Console.WriteLine(Resources.ExcelClosing);
                if (wb != null && !workbookClosed)
                    CloseWorkbook(wb);

                if (xl != null)
                    xl.Quit();

                ReleaseComObject(wb);
                ReleaseComObject(workbooks);
                ReleaseComObject(xl);
                CleanupComReferences();
                TerminateExcelProcess(excelProcessId);
                Console.WriteLine(Resources.ExcelClosed);
            }
        }

        /// <summary>
        /// Gets the process identifier for the specified Excel application.
        /// </summary>
        /// <param name="application">The Excel application to inspect.</param>
        /// <returns>The Excel process identifier, or null when it cannot be determined.</returns>
        private static int? GetExcelProcessId(Microsoft.Office.Interop.Excel.Application application)
        {
            if (!OperatingSystem.IsWindows())
                return null;

            try
            {
                _ = GetWindowThreadProcessId(new IntPtr(application.Hwnd), out var processId);
                return processId == 0 ? null : processId;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Terminates the specific Excel process if graceful COM shutdown left it running.
        /// </summary>
        /// <param name="processId">The process identifier captured from the Excel application.</param>
        private static void TerminateExcelProcess(int? processId)
        {
            if (!processId.HasValue)
                return;

            try
            {
                using var process = System.Diagnostics.Process.GetProcessById(processId.Value);
                if (process.HasExited)
                    return;

                process.Kill();
                process.WaitForExit(5000);
            }
            catch
            {
            }
        }

        /// <summary>
        /// Gets the process identifier associated with a window handle.
        /// </summary>
        /// <param name="hWnd">The window handle to inspect.</param>
        /// <param name="processId">The process identifier associated with the handle.</param>
        /// <returns>The identifier of the thread that created the window.</returns>
        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        /// <summary>
        /// Closes an Excel workbook while suppressing cleanup-time exceptions.
        /// </summary>
        /// <param name="workbook">The workbook to close.</param>
        private static void CloseWorkbook(Workbook workbook)
        {
            try
            {
                workbook.Close(SaveChanges: false);
            }
            catch
            {
            }
        }

        /// <summary>
        /// Releases a COM object while suppressing cleanup-time exceptions.
        /// </summary>
        /// <param name="comObject">The COM object to release.</param>
        private static void ReleaseComObject(object? comObject)
        {
            if (!OperatingSystem.IsWindows() || comObject == null || !Marshal.IsComObject(comObject))
                return;

            try
            {
                Marshal.FinalReleaseComObject(comObject);
            }
            catch
            {
            }
        }

        /// <summary>
        /// Runs garbage collection to release remaining runtime-callable wrappers.
        /// </summary>
        private static void CleanupComReferences()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
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
            References? refs = null;
            var projFile = Path.Combine(dir, proj.Name + ".proj");
            var lst = new List<ReferenceInfo>();

            try
            {
                refs = proj.References;
                foreach (Reference rf in refs)
                {
                    try
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
                    finally
                    {
                        ReleaseComObject(rf);
                    }
                }
            }
            finally
            {
                ReleaseComObject(refs);
                ReleaseComObject(proj);
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
                Console.WriteLine(Resources.OutputFolderNotSpecified);
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
            VBComponents? components = null;
            try
            {
                components = proj.VBComponents;
                foreach (VBComponent comp in components)
                {
                    try
                    {
                        /*
                         * Check if we need to ignore a component.
                         */
                        Console.Write(Resources.Processing, $"{wb.Name}.{comp.Name}");
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
                        Console.WriteLine(Resources.ItsA, comp.Type.ToString()[(comp.Type.ToString().LastIndexOf('_') + 1)..]);
                        if (!Directory.Exists(dir))
                            Directory.CreateDirectory(dir);

                        /*
                         * Extracts the component to the file.
                         */
                        comp.Export(filePath);
                    }
                    finally
                    {
                        ReleaseComObject(comp);
                    }
                }
            }
            finally
            {
                ReleaseComObject(components);
                ReleaseComObject(proj);
            }
#pragma warning restore CS8604
        }
    }
}

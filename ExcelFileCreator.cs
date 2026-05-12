using ExcelFusion.Exceptions;
using ExcelFusion.Properties;
using Microsoft.Vbe.Interop;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace ExcelFusion
{
    /// <summary>
    /// Contains methods to create an Excel file.
    /// </summary>
    internal static class ExcelFileCreator
    {

        /// <summary>
        /// Created the Excel file, based on the specified <see cref="CreateOptions"/>.
        /// </summary>
        /// <param name="options">A <see cref="CreateOptions"/> object containing information to generate the Excel file.</param>
        public static void CreateExcelFile(CreateOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);

            /*
             * Check if the folder exists
             */
            if (!Directory.Exists(options.Folder))
            {
                Console.WriteLine(Resources.FolderNotFoundMessage, options.Folder);
                return;
            }

            /*
             * Compress all the folder to a ZIP file with an Excel extension.
             */
            var di = new DirectoryInfo(options.Folder);
            options.Folder = di.FullName;
            options.Ext ??= ".xlsx";
            if (string.IsNullOrEmpty(options.Out))
                options.Out = di.FullName + (options.Ext.StartsWith('.') ? options.Ext : "." + options.Ext);

            if (File.Exists(options.Out))
                File.Delete(options.Out);

            using var fs = new FileStream(options.Out, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Create);
            ZipHelpers.ZipAppend(zip, di);
        }

        /// <summary>
        /// Includes all the VBA components within the ‘.vba’ folder into the Excel file created.
        /// </summary>
        /// <param name="options">A <see cref="CreateOptions"/> object containing information to generate the Excel file.</param>
        public static void IncludeVbaComponents(CreateOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);
            if (!Directory.Exists(options.Folder))
            {
                Console.WriteLine(Resources.FolderNotFoundMessage, options.Folder);
                return;
            }

            var vbaFolder = Path.Combine(options.Folder, ".vba");
            if (Directory.Exists(vbaFolder))
            {
                Microsoft.Office.Interop.Excel.Application? xl = null;
                Microsoft.Office.Interop.Excel.Workbook? wb = null;
                VBProject? proj = null;
                object? btnCompile = null;
                var saveChanges = false;
                int? excelProcessId = null;

                /*
                 * Open Excel and the Excel file
                 */
                try
                {
                    Console.WriteLine(Resources.ExcelOpening);
                    xl = new Microsoft.Office.Interop.Excel.Application
                    {
                        Visible = true,
                        EnableEvents = false,
                        DisplayAlerts = false,
                        ScreenUpdating = false,
                    };
                    excelProcessId = GetExcelProcessId(xl);
                    Console.WriteLine(Resources.ExcelOpen);
                    var start = DateTime.Now;
                    Console.WriteLine(Resources.Opening, options.Out);
#pragma warning disable CS8604 // options.Out is not null
                    var xlFilePath = new FileInfo(options.Out).FullName;
#pragma warning restore CS8604 // 
                    wb = xl.Workbooks.Open(xlFilePath, AddToMru: false);
                    wb.Activate();
                    Console.WriteLine(Resources.Open, options.Out);

                    /*
                     * Get the list of files of the VBA project.
                     */
                    proj = wb.VBProject;
                    var vbDi = new DirectoryInfo(vbaFolder);
                    var files = vbDi.GetFiles().Where((x) => (".bas;.cls;.frm").Contains(x.Extension, StringComparison.InvariantCultureIgnoreCase)).ToList<FileInfo>();

                    InjectCodeInDocComponents(proj, files);
                    InjectCodeInComponents(proj, files);
                    InjectReferences(proj, vbDi);

                    /*
                     * Try to compile the VBA project
                     */
                    btnCompile = proj.VBE.CommandBars.FindControl(Type: 1, Id: 578);
                    try
                    {
                        if (btnCompile != null)
                        {
                            dynamic compileButton = btnCompile;
                            if (compileButton.Enabled)
                                compileButton.Execute();
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new VbaCompilationException(Resources.VbaCompileError, ex);
                    }

                    saveChanges = true;
                }
                finally
                {
                    if (wb != null)
                        CloseWorkbook(wb, saveChanges);

                    if (xl != null)
                        xl.Quit();

                    ReleaseComObject(btnCompile);
                    ReleaseComObject(proj);
                    ReleaseComObject(wb);
                    ReleaseComObject(xl);
                    CleanupComReferences();
                    TerminateExcelProcess(excelProcessId);
                }
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
        /// <param name="saveChanges">A value indicating whether workbook changes should be saved.</param>
        private static void CloseWorkbook(Microsoft.Office.Interop.Excel.Workbook workbook, bool saveChanges)
        {
            try
            {
                workbook.Close(SaveChanges: saveChanges);
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
        /// Processes all the references for the VBA project.
        /// </summary>
        /// <param name="proj">A <see cref="VBProject"/> object.</param>
        /// <param name="dir">A <see cref="DirectoryInfo"/> object that contains information about the ‘.vba’ folder.</param>
        private static void InjectReferences(VBProject proj, DirectoryInfo dir)
        {
            var file = dir.GetFiles("*.proj").FirstOrDefault();
            if (file == null) return;

            using var reader = file.OpenText();
            var text = reader.ReadToEnd();
            List<ReferenceInfo>? lst;
            try
            {
                lst = JsonSerializer.Deserialize<List<ReferenceInfo>>(text);
                if (lst == null)
                    throw new Exception();
            }
            catch (Exception)
            {
                Console.WriteLine(Resources.CouldNotDeserialize, file);
                return;
            }

            foreach(var item in lst)
            {
                try
                {
                    var rf = proj.References.Item(item.Name);
                    var vbGuid = new Guid(rf.Guid);
                    if (vbGuid == item.Guid)
                        continue;

                    Console.WriteLine(Resources.ReferenceGuidError, item.Name);
                    Console.WriteLine(Resources.ReferenceGuidProject, item.Guid);
                    Console.WriteLine(Resources.ReferenceGuiVbProject, vbGuid);
                    Console.WriteLine(Resources.ReferenceRemoved);
                    proj.References.Remove(rf);
                }
                catch (Exception) { }

                /*
                 * Adds new references to the project.
                 */
                if (File.Exists(item.FullPath))
                {
                    try
                    {
                        var rf = proj.References.AddFromFile(Path.GetFullPath(item.FullPath));
                        Console.WriteLine(Resources.ReferenceAdded, Path.GetFileName(item.FullPath), item.Guid);
                        continue;
                    }
                    catch (Exception)
                    {
                        Console.WriteLine(Resources.ReferenceFailed, Path.GetFileName(item.FullPath));
                    }
                }
            }
            reader.Close();
        }

        /// <summary>
        /// Processes all the files not related to documents. <see cref="InjectCodeInDocComponents(VBProject, List{FileInfo})"/>.
        /// </summary>
        /// <param name="proj">A <see cref="VBProject"/> object.</param>
        /// <param name="files">A <see cref="List{FileInfo}"/> that contains the files.</param>
        private static void InjectCodeInComponents(VBProject proj, List<FileInfo> files)
        {
            /*
             * Processes all the other files
             */
            foreach (var file in files)
            {
                Console.WriteLine(Resources.Processing, file);
                VBComponent comp;
                try
                {
                    comp = proj.VBComponents.Item(Path.GetFileNameWithoutExtension(file.Name));
                    proj.VBComponents.Remove(comp);
                }
                catch { }

                proj.VBComponents.Import(file.FullName);
            }
        }

        /// <summary>
        /// Process the code for all documents within the project.
        /// </summary>
        /// <param name="proj">A <see cref="VBProject"/> object.</param>
        /// <param name="files">A <see cref="List{FileInfo}"/> that contains the files.</param>
        private static void InjectCodeInDocComponents(VBProject proj, List<FileInfo> files)
        {
            /*
             * Processes the documents -- they cannot be removed.
             */
            foreach (var doc in proj.VBComponents.OfType<VBComponent>().Where((x) => x.Type == vbext_ComponentType.vbext_ct_Document))
            {
                var file = files.FirstOrDefault((x) => x.Name.StartsWith(doc.Name[(doc.Name.LastIndexOf('.') + 1)..]));
                if (file != null)
                {
                    Console.WriteLine(Resources.Processing, file.Name);

                    using var reader = file.OpenText();
                    var lines = reader.ReadToEnd().Replace("\r", "").Split('\n');
                    var text = string.Join("\r\n", lines, 9, lines.Length - 9);

                    doc.CodeModule.DeleteLines(1, doc.CodeModule.CountOfLines);
                    doc.CodeModule.AddFromString(text);

                    files.Remove(file);
                }
            }
        }
    }
}

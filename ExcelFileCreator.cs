using Microsoft.Vbe.Interop;
using System.Text.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using ExcelFusion.Exceptions;

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
                Console.WriteLine(ResourceStrings.FolderNotFoundMessage, options.Folder);
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
                Console.WriteLine(ResourceStrings.FolderNotFoundMessage, options.Folder);
                return;
            }

            var vbaFolder = Path.Combine(options.Folder, ".vba");
            if (Directory.Exists(vbaFolder))
            {
                /*
                 * Open Excel and the Excel file
                 */
                Console.WriteLine(ResourceStrings.ExcelOpening);
                var xl = new Microsoft.Office.Interop.Excel.Application
                {
                    Visible = true,
                    EnableEvents = true,
                    DisplayAlerts = true,
                    ScreenUpdating = true,
                };
                Console.WriteLine(ResourceStrings.ExcelOpen);
                var start = DateTime.Now;
                Console.WriteLine(ResourceStrings.Opening, options.Out);
                var wb = xl.Workbooks.Open(options.Out);
                wb.Activate();
                Console.WriteLine(ResourceStrings.Open, options.Out);

                /*
                 * Get the list of files of the VBA project.
                 */
                var proj = wb.VBProject;
                var vbDi = new DirectoryInfo(vbaFolder);
                var files = vbDi.GetFiles().Where((x) => (".bas;.cls;.frm").Contains(x.Extension, StringComparison.InvariantCultureIgnoreCase)).ToList<FileInfo>();

                InjectCodeInDocComponents(proj, files);
                InjectCodeInComponents(proj, files);
                InjectReferences(proj, vbDi);

                /*
                 * Try to compile the VBA project
                 */
                var btnCompile = proj.VBE.CommandBars.FindControl(Type: 1, Id: 578);
                try
                {
                    if (btnCompile!= null && btnCompile.Enabled)
                        btnCompile?.Execute();
                }
                catch (Exception ex)
                {
                    throw new VbaCompilationException(ResourceStrings.VbaCompileError, ex);
                }

                wb.Close(SaveChanges: true);
                xl.Quit();
            }
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
                Console.WriteLine(ResourceStrings.CouldNotDeserialize, file);
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

                    Console.WriteLine(ResourceStrings.ReferenceGuidError, item.Name);
                    Console.WriteLine(ResourceStrings.ReferenceGuidProject, item.Guid);
                    Console.WriteLine(ResourceStrings.ReferenceGuiVbProject, vbGuid);
                    Console.WriteLine(ResourceStrings.ReferenceRemoved);
                    proj.References.Remove(rf);
                }
                catch (Exception) { }

                /*
                 * Adds new references to the project.
                 */
                if (!item.Guid.HasValue && File.Exists(item.FullPath))
                {
                    try
                    {
                        var rf = proj.References.AddFromFile(item.FullPath);
                        Console.WriteLine(ResourceStrings.ReferenceAdded, Path.GetFileName(item.FullPath), item.Guid);
                        continue;
                    }
                    catch (Exception)
                    {
                        Console.WriteLine(ResourceStrings.ReferenceFailed, Path.GetFileName(item.FullPath));
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
                Console.WriteLine(ResourceStrings.Processing, file);
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
                    Console.WriteLine(ResourceStrings.Processing, file.Name);

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
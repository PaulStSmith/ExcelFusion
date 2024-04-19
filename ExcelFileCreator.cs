using Microsoft.Vbe.Interop;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace AutomateExcel
{
    internal static class ExcelFileCreator
    {

        /// <summary>
        /// Created the Excel file, based on the specified <see cref="CreateOptions"/>.
        /// </summary>
        /// <param name="options">A <see cref="CreateOptions"/> object containing information to generate the Excel file.</param>
        public static void CreateExcelFile(CreateOptions options)
        {
            /*
             * Compress all the folder to a ZIP file with an Excel extension.
             */
            var di = new DirectoryInfo(options.Folder);
            options.Folder = di.FullName;
            options.Ext ??= ".xlsx";
            if (string.IsNullOrEmpty(options.Out))
                options.Out = di.FullName + (options.Ext.StartsWith(".") ? options.Ext : "." + options.Ext);

            if (File.Exists(options.Out))
                File.Delete(options.Out);

            using var fs = new FileStream(options.Out, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Create);
            ZipHelpers.ZipAppend(zip, di);
            zip.Dispose();
            fs.Close();
        }

        /// <summary>
        /// Includes all the VBA components within the ‘.vba’ folder into the Excel file created.
        /// </summary>
        /// <param name="options">A <see cref="CreateOptions"/> object containing information to generate the Excel file.</param>
        public static void IncludeVbaComponents(CreateOptions options)
        {
            var vbaFolder = Path.Combine(options.Folder, ".vba");
            if (Directory.Exists(vbaFolder))
            {
                /*
                 * Open Excel and the Excel file
                 */
                Console.WriteLine("Opening Excel...");
                var xl = new Microsoft.Office.Interop.Excel.Application
                {
                    Visible = true
                };
                Console.WriteLine("Excel open.");
                var start = DateTime.Now;
                Console.WriteLine($"Opening ‘{options.Out}’...");
                var wb = xl.Workbooks.Open(options.Out);
                wb.Activate();
                Console.WriteLine($"‘{options.Out}’ Open.");

                /*
                 * Get the list of files of the VBA project.
                 */
                var proj = wb.VBProject;
                var vbDi = new DirectoryInfo(vbaFolder);
                var files = vbDi.GetFiles().Where((x) => (".bas;.cls;.frm").Contains(x.Extension.ToLowerInvariant())).ToList<FileInfo>();

                ProcessDocuments(proj, files);
                ProcessFiles(proj, files);
                ProcessReferences(proj, vbDi);

                wb.Close(SaveChanges: true);
                xl.Quit();
            }
        }

        /// <summary>
        /// Processes all the references for the VBA project.
        /// </summary>
        /// <param name="proj">A <see cref="VBProject"/> object.</param>
        /// <param name="dir">A <see cref="DirectoryInfo"/> object that contains information about the ‘.vba’ folder.</param>
        private static void ProcessReferences(VBProject proj, DirectoryInfo dir)
        {
            var file = dir.GetFiles("*.proj").FirstOrDefault();
            if (file == null) return;

            using var reader = file.OpenText();
            var text = reader.ReadToEnd();
            var lst = JsonConvert.DeserializeObject<List<ReferenceInfo>>(text);
            foreach(var item in lst)
            {
                try
                {
                    var rf = proj.References.Item(item.Name);
                }
                catch (Exception)
                {
                    /*
                     * Adds new references to the project.
                     */
                    if (!item.Guid.HasValue && File.Exists(item.FullPath))
                    {
                        try
                        {
                            var rf = proj.References.AddFromFile(item.FullPath);
                            Console.WriteLine($@"Reference to {Path.GetFileName(item.FullPath)} added to the project. GUID = ‘{item.Guid}’");
                            continue;
                        }
                        catch (Exception)
                        {
                            Console.WriteLine($@"Failed to add reference to {Path.GetFileName(item.FullPath)}.");
                        }
                    }
                }
            }
            reader.Close();
        }

        /// <summary>
        /// Processes all the files not related to documents. <see cref="ProcessDocuments(VBProject, List{FileInfo})"/>.
        /// </summary>
        /// <param name="proj">A <see cref="VBProject"/> object.</param>
        /// <param name="files">A <see cref="List{FileInfo}"/> that contains the files.</param>
        private static void ProcessFiles(VBProject proj, List<FileInfo> files)
        {
            /*
             * Processes all the other files
             */
            foreach (var file in files)
            {
                Console.WriteLine($"Processing {file.Name} ...");
                VBComponent comp;
                try
                {
                    comp = proj.VBComponents.Item(Path.GetFileNameWithoutExtension(file.Name));
                    proj.VBComponents.Remove(comp);
                }
                catch (Exception)
                {
                }

                proj.VBComponents.Import(file.FullName);
            }
        }

        /// <summary>
        /// Process the code for all documents within the project.
        /// </summary>
        /// <param name="proj">A <see cref="VBProject"/> object.</param>
        /// <param name="files">A <see cref="List{FileInfo}"/> that contains the files.</param>
        private static void ProcessDocuments(VBProject proj, List<FileInfo> files)
        {
            /*
             * Processes the documents -- they cannot be removed.
             */
            foreach (var doc in proj.VBComponents.OfType<VBComponent>().Where((x) => x.Type == vbext_ComponentType.vbext_ct_Document))
            {
                var file = files.Where((x) => x.Name.StartsWith(doc.Name.Substring(doc.Name.LastIndexOf(".") + 1))).FirstOrDefault();
                if (file != null)
                {
                    Console.WriteLine($"Processing {file.Name} ...");

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
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace AutomateExcel
{
    /// <summary>
    /// Contains methods to extract Visual Basic source code from an Excel file.
    /// </summary>
    internal static class VbaExtractor
    {

        /// <summary>
        /// Extracts the VBA code from the Excel file specified within the <see cref="ExtractOptions"/> object.
        /// </summary>
        /// <param name="options">An <see cref="ExtractOptions"/> object containing the data to access the Excel file.</param>
        public static void ExtractVbaSourceCode(ExtractOptions options)
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
            Console.WriteLine($"Opening ‘{options.ExcelFile}’...");
            var wb = xl.Workbooks.Open(options.ExcelFile);
            wb.Activate();
            Console.WriteLine($"‘{options.ExcelFile}’ Open.");

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
            Console.WriteLine("Closing Excel...");
            xl.Quit();
            Console.WriteLine("Closed.");
        }

        /// <summary>
        /// Extracts all the references from the Visual Basic project.
        /// </summary>
        /// <param name="options">An <see cref="ExtractOptions"/> object containing the data to access the Excel file.</param>
        /// <p7aram name="wb">A <see cref="Workbook"/> object containing the Visual Basic project to extract.</param>
        private static void ExtractReferences(ExtractOptions options, Workbook wb)
        {
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
            writer.WriteLine(JsonConvert.SerializeObject(lst, Formatting.Indented));
            writer.Close();
        }

        /// <summary>
        /// Exports the Visual Basic project contained in the specified <see cref="Workbook"/>.
        /// </summary>
        /// <param name="options">An <see cref="ExtractOptions"/> object containing the data to access the Excel file.</param>
        /// <p7aram name="wb">A <see cref="Workbook"/> object containing the Visual Basic project to extract.</param>
        public static void ExtractVbProject(ExtractOptions options, Workbook wb)
        {
            var proj = wb.VBProject;
            foreach (VBComponent comp in proj.VBComponents)
            {
                /*
                 * Check if we need to ignore a component.
                 */
                Console.Write($"Processing {wb.Name}.{comp.Name} ... ");
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
                Console.WriteLine($"it's a {comp.Type.ToString().Substring(comp.Type.ToString().LastIndexOf('_') + 1)}. Exporting.");
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                /*
                 * Extracts the component to the file.
                 */
                comp.Export(filePath);
            }
        }
    }
}
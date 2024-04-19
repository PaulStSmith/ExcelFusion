using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.NamingConventionBinder;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

#nullable enable 

namespace AutomateExcel
{
    internal class Program
    {
        /// <summary>
        /// Entrypoint for the program.
        /// </summary>
        /// <param name="args">Command line arguments passed to the program.</param>
        static void Main(string[] args)
        {
            var cmdExtract = new Command("e", "Extracts the specified Excel file") {
                new Argument<string>(name:"ExcelFile", description:"A path to an Excel file."),
                new Option<string>(aliases: ProgramHelpers.GenerateAliases("out"), description:"The folder where to extact the Excel file.")
            };

            var cmdCreate = new Command("c", "Creates an Excel file based on a folder") {
                new Argument<string>(name:"Folder", description:"A path to a folder containing the structure of an Excel file."),
                new Option<string>(aliases: ProgramHelpers.GenerateAliases("out"), description:"The name of the Excel file output."),
                new Option<string>(aliases: ProgramHelpers.GenerateAliases("ext"), description:"The name of the Extension of the Excel file.", getDefaultValue:()=>"xlsx")
            };

            ConfigureExportHandler(cmdExtract);
            ConfigureCreateHandler(cmdCreate);

            var root = new RootCommand("Manipulates Excel files.")
            {
                cmdExtract,
                cmdCreate
            };

            root.Invoke(args);
            Console.WriteLine("Press [Enter] to continue.");
            Console.ReadLine();
        }

        /// <summary>
        /// Configures the command handler for the create command.
        /// </summary>
        /// <param name="cmdCreate">The command data.</param>
        private static void ConfigureCreateHandler(Command cmdCreate)
        {
            cmdCreate.Handler = CommandHandler.Create<CreateOptions>((options) =>
            {
                /*
                 * Check if the folder exists
                 */
                if (!Directory.Exists(options.Folder))
                {
                    Console.WriteLine("Folder not found.");
                    return 99;
                }

                ExcelFileCreator.CreateExcelFile(options);
                ExcelFileCreator.
                                IncludeVbaComponents(options);

                return 0;
            });
        }

        /// <summary>
        /// Configures the command handler for the extract command.
        /// </summary>
        /// <param name="cmdExtract">The command data.</param>
        private static void ConfigureExportHandler(Command cmdExtract)
        {
            cmdExtract.Handler = CommandHandler.Create<ExtractOptions>((options) =>
            {
                /*
                 * If the file doesn't exist display message and exit.
                 */
                if (!File.Exists(options.ExcelFile))
                {
                    Console.WriteLine($"Excel file '{options.ExcelFile}' not found");
                    return 99;
                }

                /*
                 * Guarantees that we have an output folder.
                 */
                options.Out = options.Out?.Trim();
                if (string.IsNullOrEmpty(options.Out))
                    options.Out = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(options.ExcelFile)) ?? "", Path.GetFileNameWithoutExtension(options.ExcelFile));

                /*
                 * Extracts the Excel file to the Output folder, creating the directory if needed.
                 */
                Console.WriteLine($"Extracting {Path.GetFileName(options.ExcelFile)} to ‘{options.Out}’");
                if (!Directory.Exists(options.Out))
                    Directory.CreateDirectory(options.Out);
                ZipHelpers.ExtractFiles(options);
                VbaExtractor.

                /*
                 * Extracts the VBA source code.
                 */
                ExtractVbaSourceCode(options);

                return 0;
            });
        }
    }
}

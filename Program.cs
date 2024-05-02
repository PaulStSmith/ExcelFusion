using System;
using System.CommandLine;
using System.CommandLine.NamingConventionBinder;
using System.IO;

namespace ExcelFusion
{
    internal class Program
    {
        /// <summary>
        /// Entrypoint for the program.
        /// </summary>
        /// <param name="args">Command line arguments passed to the program.</param>
        static void Main(string[] args)
        {
            var cmdExtract = new Command("e", ResourceStrings.ExtractDescription) {
                new Argument<string>(name:"ExcelFile", description:ResourceStrings.ExtractArgumentDescription),
                new Option<string>(aliases: ProgramHelpers.GenerateAliases("out"), description:ResourceStrings.ExtractOutDescription)
            };

            var cmdCreate = new Command("c", ResourceStrings.CreateDescription) {
                new Argument<string>(name:"Folder", description:ResourceStrings.CreateArgumentDescription),
                new Option<string>(aliases: ProgramHelpers.GenerateAliases("out"), description:ResourceStrings.CreateOutDescription),
                new Option<string>(aliases: ProgramHelpers.GenerateAliases("ext"), description:ResourceStrings.CreateExtDescription, getDefaultValue:()=>"xlsx")
            };

            ConfigureExportHandler(cmdExtract);
            ConfigureCreateHandler(cmdCreate);

            var root = new RootCommand("Manipulates Excel files.")
            {
                cmdExtract,
                cmdCreate
            };

            Console.WriteLine(ResourceStrings.Header);
            root.Invoke(args);
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
                    Console.WriteLine(ResourceStrings.FolderNotFoundMessage);
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
                    Console.WriteLine(ResourceStrings.FileNotFoundMessage, options.ExcelFile);
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

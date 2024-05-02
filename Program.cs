using System;
using System.CommandLine;
using System.CommandLine.Builder;
using System.CommandLine.NamingConventionBinder;
using System.CommandLine.Parsing;
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
            /*
             * Prepare the extract command.
             */
            var cmdExtract = new Command("extract", ResourceStrings.ExtractDescription) {
                new Argument<string>(name:"ExcelFile", description:ResourceStrings.ExtractArgumentDescription),
                new Option<string>(aliases: ProgramHelpers.GenerateAliases("output"), description:ResourceStrings.ExtractOutDescription)
            };
            cmdExtract.AddAlias("e");

            /*
             * Prepare the build command.
             */
            var cmdBuild = new Command("build", ResourceStrings.CreateDescription) {
                new Argument<string>(name:"Folder", description:ResourceStrings.CreateArgumentDescription),
                new Option<string>(aliases: ProgramHelpers.GenerateAliases("output"), description:ResourceStrings.CreateOutDescription),
                new Option<string>(aliases: ProgramHelpers.GenerateAliases("extension"), description:ResourceStrings.CreateExtDescription, getDefaultValue:()=>"xlsx")
            };
            cmdBuild.AddAlias("b");

            /*
             * Prepare the license command.
             */
            var cmdLicense = new Command("license", ResourceStrings.LicenseDescription);
            cmdLicense.AddAlias("l");

            /*
             * Configure the handlers for each command.
             */
            ConfigureExportHandler(cmdExtract);
            ConfigureCreateHandler(cmdBuild);
            ConfigureLicenseHandler(cmdLicense);

            /*
             * Create the root command and add the subcommands.
             */
            var root = new RootCommand(ResourceStrings.RootCommandDescription)
            {
                cmdExtract,
                cmdBuild,
                cmdLicense
            };

            /*
             * Build the command line parser.
             */
            var parser = new CommandLineBuilder(root)
                                    .UseHelp()
                                    .UseEnvironmentVariableDirective()
                                    .UseParseDirective()
                                    .UseSuggestDirective()
                                    .RegisterWithDotnetSuggest()
                                    .UseTypoCorrections()
                                    .UseParseErrorReporting()
                                    .UseExceptionHandler()
                                    .CancelOnProcessTermination()
                                    .Build();

            /*
             * Display the header and invoke the parser.
             */
            Console.WriteLine(ResourceStrings.Header);
            parser.Invoke(args);
        }

        /// <summary>
        /// Configures the command handler for the license command.
        /// </summary>
        /// <param name="cmdLicense">The command data.</param>
        private static void ConfigureLicenseHandler(Command cmdLicense)
        {
            /*
             * Display the license information.
             */
            cmdLicense.Handler = CommandHandler.Create(() =>
            {
                Console.WriteLine(ResourceStrings.MitLicense);
                return 0;
            });
        }

        /// <summary>
        /// Configures the command handler for the create command.
        /// </summary>
        /// <param name="cmdCreate">The command data.</param>
        private static void ConfigureCreateHandler(Command cmdCreate)
        {
            /*
             * Create the Excel file.
             */
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
                ExcelFileCreator.IncludeVbaComponents(options);

                return 0;
            });
        }

        /// <summary>
        /// Configures the command handler for the extract command.
        /// </summary>
        /// <param name="cmdExtract">The command data.</param>
        private static void ConfigureExportHandler(Command cmdExtract)
        {
            /*
             * Extract the Excel file.
             */
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
                Console.WriteLine(ResourceStrings.Extracing, Path.GetFileName(options.ExcelFile), options.Out);
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

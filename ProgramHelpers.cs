using ExcelFusion;
using System;
using System.Runtime.InteropServices;

/// <summary>
/// Contains helper functions for the program.
/// </summary>
internal static class ProgramHelpers
{

    /// <summary>
    /// Returns a path relative to the specified <paramref name="basePath"/>.
    /// </summary>
    /// <param name="basePath">The base to make the <paramref name="fullName"/> relative to.</param>
    /// <param name="fullName">The full path to make relative to <paramref name="basePath"/>.</param>
    /// <returns></returns>
    public static string GetRelativePath(string basePath, string fullName)
    {
        if (fullName.StartsWith(basePath, StringComparison.OrdinalIgnoreCase))
        {
            return fullName.Substring(basePath.Length + 1);
        }
        return fullName;
    }

    /// <summary>
    /// Handles the specified <see cref="Exception"/>.
    /// </summary>
    /// <param name="ex">The exception to handle.</param>
    /// <returns>True if the exception was handled, false otherwise.</returns>
    public static bool HandleException(COMException ex)
    {
        var clr = Console.ForegroundColor;
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine(ex.Message);

        if (ex.Message.StartsWith("Programmatic access to Visual Basic Project is not trusted"))
        {
            Console.ForegroundColor = clr;
            Console.WriteLine(ResourceStrings.GrantAccess);
            Console.WriteLine(ResourceStrings.TryAgain);
            var c = Console.ReadKey();
            if (c.Key == ConsoleKey.Y)
                return false;
        }
        else
        {
            Exception e = ex;
            while (e.InnerException != null)
            {
                Console.WriteLine(e.InnerException.Message);
                e = e.InnerException;
            }
        }
        Console.ForegroundColor = clr;
        return true;
    }

    /// <summary>
    /// Generates the aliases for the specified option
    /// </summary>
    /// <param name="option">The option to generate aliases.</param>
    /// <returns>Aliases for the specified option.</returns>
    public static string[] GenerateAliases(string option)
    {
        return new string[] { $"/{option}", $"--{option}" };
    }
}
using ExcelFusion.Properties;
using System;
using System.IO;
using System.IO.Compression;
using System.Threading;

#nullable enable

namespace ExcelFusion
{
    internal static class ZipHelpers
    {

        /// <summary>
        /// Extract the files inside the Excel file specified within <see cref="ExtractOptions"/>.
        /// </summary>
        /// <param name="options">An <see cref="ExtractOptions"/> object containing the data necessary to extract the files.</param>
        public static void ExtractFiles(ExtractOptions options)
        {
            if (options.Out == null || options.ExcelFile == null) return;
            using var fs = new FileStream(options.ExcelFile, FileMode.Open, FileAccess.Read);
            using var zip = new ZipArchive(fs);
            foreach (var entry in zip.Entries)
            {
                var dest = Path.Combine(options.Out, entry.FullName.Replace('/', '\\'));
                var dir = Path.GetDirectoryName(dest) ?? "";
                Console.WriteLine(Resources.Extracing, entry.FullName, dir);
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                var buffer = new byte[4096];
                using var fsEntry = new FileStream(dest, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                using var zipStream = entry.Open();
                while (true)
                {
                    var r = zipStream.Read(buffer, 0, buffer.Length);
                    if (r <= 0) break;
                    fsEntry.Write(buffer, 0, r);
                }
                zipStream.Close();
                fsEntry.Close();
            }
            fs.Close();
        }

        /// <summary>
        /// Appends the folder specified by <paramref name="di"/> into the zip archive specified by <paramref name="zip"/>.
        /// </summary>
        /// <param name="zip">A <see cref="ZipArchive"/> that represents a ZIP file.</param>
        /// <param name="di">A <see cref="DirectoryInfo"/> that specifies a folder.</param>
        public static void ZipAppend(ZipArchive zip, DirectoryInfo di)
        {
            ZipAppend(zip, di, null);
        }

        /// <summary>
        /// Appends the folder specified by <paramref name="di"/> into the zip archive specified by <paramref name="zip"/>.
        /// </summary>
        /// <param name="zip">A <see cref="ZipArchive"/> that represents a ZIP file.</param>
        /// <param name="di">A <see cref="DirectoryInfo"/> that specifies a folder.</param>
        /// <param name="basePath">The base path of the zip file.</param>
        public static void ZipAppend(ZipArchive zip, DirectoryInfo di, string? basePath)
        {
            /*
             * Ignore folder named ‘.vba’
             */
            if (di.Name.StartsWith(".vba")) return;

            basePath ??= di.FullName;
            foreach (var d in di.GetDirectories())
                ZipAppend(zip, d, basePath);

            foreach (var f in di.GetFiles())
            {
                var buffer = new byte[4096];
                var entryName = ProgramHelpers.GetRelativePath(basePath, f.FullName);
                var entry = zip.CreateEntry(entryName.Replace('\\', '/'));
                Console.WriteLine(Resources.Compressing , entryName);
                using var strm = entry.Open();
                using var fs = f.OpenRead();
                while (true)
                {
                    var r = fs.Read(buffer, 0, buffer.Length);
                    if (r == 0) break;
                    strm.Write(buffer, 0, r);
                    Thread.Sleep(1);
                }
                strm.Flush();
                fs.Close();
            }
        }
    }
}
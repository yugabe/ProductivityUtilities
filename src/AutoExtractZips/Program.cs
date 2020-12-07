using System;
using System.IO;
using System.IO.Compression;
using System.Linq;

var allFiles = new DirectoryInfo(args?[0] ?? Directory.GetCurrentDirectory()).EnumerateFiles("*.zip", SearchOption.AllDirectories).ToList();

foreach (var (file, number) in allFiles.Select((f, i) => (f, i + 1)))
{
    {
        using var zip = new ZipArchive(file.OpenRead(), ZipArchiveMode.Read);
        var destination = file.Directory!.CreateSubdirectory(Path.GetFileNameWithoutExtension(file.Name));
        zip.ExtractToDirectory(destination.FullName);
        Console.Write($"{number} / {allFiles.Count} | ");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("Extracted");
        Console.ResetColor();
        Console.WriteLine($" {zip.Entries.Count} files to {destination.FullName}.");
    }
    file.Delete();
}

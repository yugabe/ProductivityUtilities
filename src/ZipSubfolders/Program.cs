using System.Diagnostics;
using System.IO.Compression;

// Zips the current folder or the folder provided as the first string argument each to its own ZIP file named after the folder. Uses maximum available parallelism.

var root = args?.Length > 0 ? args[0] : Directory.GetCurrentDirectory();
var folders = new DirectoryInfo(root).GetDirectories();
var outerStopwatch = Stopwatch.StartNew();
double totalFileSizeInMb = 0, totalZipSizeInMb = 0;
Console.BufferWidth = Console.WindowWidth = new[] { Console.BufferWidth, Console.WindowWidth, folders.Length.ToString().Length * 2 + 1 + 30 + 5 + 20 + 5 + 16 + 5 + +5 + 3 * 7 + 1 }.Max();

void Write(string joinString, char padChar, string state, string number, string folder, string files, string totalSize, string zippedSize, string time)
{
    string PadRightMax(string text, int padMax)
    {
        text ??= "";
        return text.Length > padMax ? text[0..padMax] : text.PadRight(padMax, padChar);
    }
    var output = string.Join(joinString, new[] { PadRightMax(number, folders.Length.ToString().Length * 2 + 1), PadRightMax(folder, 30), PadRightMax(files, 5), PadRightMax(totalSize, 10), PadRightMax(state, 5), PadRightMax(zippedSize, 20), PadRightMax(time, 5), outerStopwatch.Elapsed.ToString("mm\\:ss") });
    Console.WriteLine(output);
}
Write("_|_", '_', "State", "#", "Folder", "Files", "TotalSize", "Zipped Size", "Time");

Parallel.ForEach(folders.Select((s, i) => (s, i)), item =>
{
    var (subfolder, subfolderIndex) = item;
    var stopwatch = Stopwatch.StartNew();
    using var zipFile = File.Create($"{subfolder.FullName}.zip");
    using var zip = new ZipArchive(zipFile, ZipArchiveMode.Create);

    var files = subfolder.GetFiles("*", SearchOption.AllDirectories);
    var filesInMb = (double)files.Sum(f => f.Length) / (1024 * 1024);
    totalFileSizeInMb += filesInMb;
    Write(" | ", ' ', "Start", $"{(subfolderIndex + 1)}/{folders.Length}", Path.GetRelativePath(root, subfolder.FullName), files.Length.ToString(), $"{filesInMb:N2} MB", "...", "...");
    foreach (var file in files)
        zip.CreateEntryFromFile(file.FullName, Path.GetRelativePath(subfolder.FullName, file.FullName));
    totalZipSizeInMb += (double)zipFile.Length / (1024 * 1024);
    Write(" | ", ' ', "Done", $"{(subfolderIndex + 1)}/{folders.Length}", Path.GetRelativePath(root, subfolder.FullName), files.Length.ToString(), $"{filesInMb:N2} MB", $"{(double)zipFile.Length / (1024 * 1024):N2} MB ({((double)zipFile.Length / (1024 * 1024) / filesInMb) * 100:N2}%)", stopwatch.Elapsed.ToString("mm\\:ss"));
});

Console.WriteLine($"Finished zipping {folders.Length} folders in {outerStopwatch.Elapsed:mm\\:ss}. Total compression ratio: {(totalZipSizeInMb / totalFileSizeInMb) * 100:N2}%");
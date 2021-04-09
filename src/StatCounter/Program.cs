using Shell32;
using System;
using System.IO;
using System.Linq;
using System.Text;

Console.WindowWidth = Console.BufferWidth = 200;
var rootDirectory = new DirectoryInfo(args.ElementAtOrDefault(0) ?? Directory.GetCurrentDirectory());
Console.WriteLine($"Running StatCounter in folder \"{rootDirectory.FullName}\"");

var shell = (dynamic)Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application")!)!;
Folder root = shell.NameSpace(rootDirectory.FullName);
var headers = Enumerable.Range(0, short.MaxValue)
    .Select(i => (index: i, header: root.GetDetailsOf(null, i)))
    .Where(e => !string.IsNullOrEmpty(e.header))
    .ToLookup(e => e.header, e => e.index);
var lengthProp = headers["Length"].Single();
var slidesProp = headers["Slides"].Single();

var sum = (length: new TimeSpan(), slides: 0);
var s = new StringBuilder().AppendLine("Fejezet\tLecke\tÖsszes hossz\tÖsszes diaszám\tFájlnév\tHossz\tDiaszám");
foreach (var file in rootDirectory.EnumerateFiles("*.*", SearchOption.AllDirectories).Where(f => f.Extension is ".mp4" or ".pptx" && !f.FullName.Contains("Nyers")).OrderBy(f => f.FullName))
{
    Folder folder = shell.NameSpace(file.DirectoryName);
    var folderItem = folder.ParseName(file.Name);

    var length = folder.GetDetailsOf(folderItem, lengthProp);
    if (TimeSpan.TryParse(length, out var timeSpan))
        sum = (sum.length + timeSpan, sum.slides);

    var slides = folder.GetDetailsOf(folderItem, slidesProp);
    if (int.TryParse(slides, out var slidesVal) && slidesVal < 35)
        sum = (sum.length, sum.slides + slidesVal);

    s.AppendLine(string.Join('\t', file.Directory!.Parent!.Name, file.Directory.Name, "", "", file.Name, length, slides));
    Console.WriteLine($"{Path.GetRelativePath(rootDirectory.FullName, file.DirectoryName!),-100} {file.Name,-70}{length,-12}{(slidesVal < 35 ? slides : $"({slides})")}");
}

var outFile = $"{rootDirectory.Name}.tsv";
File.WriteAllText(outFile, s.ToString());

Console.WriteLine($"Total: {sum.length}, {sum.slides} slides\nSaved data to file \"{outFile}\".");

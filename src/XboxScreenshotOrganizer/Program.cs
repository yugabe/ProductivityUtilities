using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Metadata.Profiles.Exif;
using System.Globalization;

const string exifDateFormat = "yyyy:MM:dd HH:mm:ss";
var quality = int.TryParse(args.ElementAtOrDefault(0), out var q) ? Math.Max(1, Math.Min(100, q)) : 75;
var author = args.ElementAtOrDefault(1) ?? "YuGabe";

var directory = new DirectoryInfo(Directory.GetCurrentDirectory());

Parallel.ForEach(directory.EnumerateFiles("*.png"), png =>
{
    using var image = Image.Load(png.OpenRead());

    var game = png.Name[..^24].Replace('_', ' ').Replace("  ", " ");
    var date = new DateTime(int.Parse(png.Name[^23..^19]), int.Parse(png.Name[^18..^16]), int.Parse(png.Name[^15..^13]),
        int.Parse(png.Name[^12..^10]), int.Parse(png.Name[^9..^7]), int.Parse(png.Name[^6..^4]));
    var exif = image.Metadata.ExifProfile = new();
    exif.SetValue(ExifTag.DateTime, DateTime.Now.ToString(exifDateFormat));
    exif.SetValue(ExifTag.DateTimeOriginal, date.ToString(exifDateFormat));
    exif.SetValue(ExifTag.XPAuthor, author);
    var title = $"{game} | Screenshot taken on Xbox at {date:yyyy-MM-dd HH:mm:ss}";
    exif.SetValue(ExifTag.ImageDescription, title);
    exif.SetValue(ExifTag.XPTitle, title);
    exif.SetValue(ExifTag.XPSubject, game);
    exif.SetValue(ExifTag.Software, "Xbox");

    var targetFileName = $"{png.FullName[..^4]}.jpg";
    image.SaveAsJpeg(targetFileName, new() { Quality = quality });

    var (original, @new) = (((float)png.Length) / 1024 / 1024, ((float)new FileInfo(targetFileName).Length) / 1024 / 1024);
    Console.WriteLine($"[{original.ToString("N2", CultureInfo.InvariantCulture),-5}MB => {@new.ToString("N2", CultureInfo.InvariantCulture),-5}MB (\e[32m{(100 * @new / original).ToString("N2", CultureInfo.InvariantCulture):-5}%\e[0m)] {png.Name[..^4]}");
});

using Aspose.Words;
using Novacode;
using OfficeOpenXml;

var now = DateTime.Now.ToString("yyyyddMM-HHmmss");

if (args?.Length != 4 || args.Any(string.IsNullOrWhiteSpace))
{
    static string ReadNonEmpty(string message, string? defaultValue = null)
    {
        string? result;
        do
        {
            Console.WriteLine($"{message}{(defaultValue == null ? "" : $" (leave empty to use default '{defaultValue}')")}:");
            result = Console.ReadLine() is var r && string.IsNullOrWhiteSpace(r) ? defaultValue : r;
        }
        while (string.IsNullOrWhiteSpace(result));
        return result;
    }

    var currentDirectory = new DirectoryInfo(Directory.GetCurrentDirectory());

    Console.WriteLine($"Usage:\nmergegen 'template.xlsx' 'template.docx' 'outputdirectory' 'outputnamefield'\n{(args?.Length == 0 ? "" : $"Arguments provided: {string.Join(", ", args!.Select(a => $"\"{a}\""))}\n")}Ctrl+C to exit or provide the parameters below.\n");

    args = [
        ReadNonEmpty("Template Excel workbook", args?.ElementAtOrDefault(0) ?? currentDirectory.EnumerateFiles("*.xlsx").FirstOrDefault()?.Name),
        ReadNonEmpty("Template Word document", args?.ElementAtOrDefault(1) ?? currentDirectory.EnumerateFiles("*.docx").FirstOrDefault()?.Name),
        ReadNonEmpty("Output directory", args?.ElementAtOrDefault(2) ?? $"MergeGen-{now}"),
        ReadNonEmpty("Output name field in template Excel", args?.ElementAtOrDefault(3) ?? "Filename")
    ];
}

string[] headers;
string[][] items;
int nameFieldIndex;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

using (var package = new ExcelPackage(new FileInfo(args[0])))
{
    var sheet = package.Workbook.Worksheets[0];
    headers = sheet.Cells["1:1"].Select(r => r.Text).TakeWhile(r => !string.IsNullOrWhiteSpace(r)).ToArray();
    nameFieldIndex = headers.Select((h, i) => (h, i)).First(e => string.Compare(e.h, args[3], true) == 0).i;
    items = Enumerable.Range(2, sheet.Cells.Rows - 1).Select(r => Enumerable.Range(1, headers.Length).Select(c => sheet.Cells[r, c].Text).ToArray()).TakeWhile(r => !r.All(c => string.IsNullOrWhiteSpace(c))).ToArray();
}

using var memoryStream = new MemoryStream();
using (var fileStream = new FileStream(args[1], FileMode.Open, FileAccess.Read))
    fileStream.CopyTo(memoryStream);

foreach (var row in items)
{
    memoryStream.Position = 0;
    var document = new Document(memoryStream);
    document.MailMerge.Execute(headers, row);
    var outputDir = Directory.CreateDirectory(args[2]).FullName;
    var fileName = Path.Combine(outputDir, row[nameFieldIndex]);
    if (File.Exists(fileName))
    {
        var backupName = Path.Combine(outputDir, $"{Path.GetFileNameWithoutExtension(fileName)} ({now}).{Path.GetExtension(fileName)}");
        File.Move(fileName, backupName);
        Console.WriteLine($"{Path.GetRelativePath(Directory.GetCurrentDirectory(), backupName)} backed up.");
    }
    using var docStream = new MemoryStream();
    document.Save(docStream, SaveFormat.Docx);
    docStream.Position = 0;

    using var docx = DocX.Load(docStream);
    docx.RemoveParagraphAt(0);
    var allHeaders = new[] { docx.Headers.first, docx.Headers.odd, docx.Headers.even }.Where(h => h != null);
    var allFooters = new[] { docx.Footers.first, docx.Footers.odd, docx.Footers.even }.Where(h => h != null);
    var allHeaderFooterParagraphs = allHeaders.SelectMany(h => h.Paragraphs).Concat(allFooters.SelectMany(f => f.Paragraphs));
    foreach (var picture in allHeaderFooterParagraphs.SelectMany(p => p.Pictures).Where(p => p.Height == 328))
        picture.Remove();

    foreach (var paragraph in allHeaderFooterParagraphs.Where(p => p.Text.Contains("Aspose")))
        paragraph.Remove(false);

    docx.SaveAs(fileName);
    Console.WriteLine($"{Path.GetRelativePath(Directory.GetCurrentDirectory(), fileName)} created.");
}

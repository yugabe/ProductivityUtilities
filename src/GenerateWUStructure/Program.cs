using GemBox.Presentation;
using OfficeOpenXml;
using System.IO;
using System.Linq;

static string N(string name) => new(name.SelectMany(c => c is ':' ? " -" : c.ToString()).Where(c => !Path.GetInvalidFileNameChars().Contains(c)).ToArray());
ComponentInfo.SetLicense("FREE-LIMITED-KEY");
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
using var excel = new ExcelPackage(new FileInfo(args.ElementAtOrDefault(2) ?? @"C:\Users\Gábor (BME)\OneDrive - BME\Projects\Webuni - ASP.NET fejlesztői alapképzés\tematika v1.xlsx"));
var sheet = excel.Workbook.Worksheets.First();
foreach (var fejezet in Enumerable.Range(2, 44).Select(i => new Row(sheet.Cells[i, 1].GetValue<int>(), sheet.Cells[i, 2].GetValue<string>(), sheet.Cells[i, 3].GetValue<int>(), sheet.Cells[i, 4].GetValue<string>())).ToList().GroupBy(r => (r.FejezetSzam, r.FejezetCim)))
{
    foreach (var (lecke, sorszam) in fejezet.OrderBy(l => l.LeckeSzam).Select((l, i) => (l, i + 1)))
    {
        var presentation = PresentationDocument.Load(args.ElementAtOrDefault(1) ?? @"C:\Users\Gábor (BME)\OneDrive - BME\Projects\Webuni - ASP.NET fejlesztői alapképzés\sablon.pptx");

        var leckeDir = new DirectoryInfo(args.ElementAtOrDefault(0) ?? @"C:\Users\Gábor (BME)\OneDrive - BME\Projects\Webuni - ASP.NET fejlesztői alapképzés\Anyagok").CreateSubdirectory(@$"{N($"{fejezet.Key.FejezetSzam} - {fejezet.Key.FejezetCim}")}\{N($"{fejezet.Key.FejezetSzam}.{sorszam} - {lecke.LeckeCim}")}");
        leckeDir.CreateSubdirectory(args.ElementAtOrDefault(4) ?? "Nyers");

        foreach (var textContent in presentation.MasterSlides.Select(m => m.TextContent).Concat(presentation.MasterSlides.SelectMany(m => m.LayoutSlides).Select(l => l.TextContent)).Concat(presentation.Slides.Select(s => s.TextContent)))
            textContent.Replace(args.ElementAtOrDefault(3) ?? "XYZ", lecke.LeckeCim);

        presentation.Save(Path.Combine(leckeDir.FullName, $"{N(lecke.LeckeCim)}.pptx"));
    }
}

record Row(int FejezetSzam, string FejezetCim, int LeckeSzam, string LeckeCim) { }

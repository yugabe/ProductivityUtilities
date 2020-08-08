using System;
using System.IO;
using System.Linq;

namespace SkewSrt
{
    public class Program
    {
        public static int Main(string[] args)
        {
            if (args?.Length != 3 || !TryParseStamp(args[1], out var firstSubStamp) || !TryParseStamp(args[2], out var secondSubStamp))
            {
                Console.WriteLine(
                    "---SkewSrt---\n" +
                    "Resynchronizes the subtitles in the given .srt file if they are out of sync. " +
                    "The subs' timestamps are skewed between the first and last timestamps, so you have to provide the correct times " +
                    "for the first and last subtitles that appear for your video file (the times they should appear). The original file will be backed up.\n" +
                    "Usage:   skewsrt filename firstsubstamp lastsubstamp\n" +
                    "Example: mysub.srt 00:01:05,311 01:34:11,210");
                return 1;
            }
            var file = new FileInfo(args[0]);
            if (!file.Exists)
            {
                Console.WriteLine($"Could not find file: '{file.FullName}'.");
                return 2;
            }
            var lines = File.ReadAllLines(file.FullName);

            TimeSpan startStamp = default, endStamp = default;
            lines.First(l => TryParseStamps(l, out startStamp, out _));
            lines.Last(l => TryParseStamps(l, out endStamp, out _));

            var (inFileDiff, desiredDiff) = (endStamp - startStamp, secondSubStamp - firstSubStamp);
            TimeSpan Skew(TimeSpan value) => firstSubStamp + (value - startStamp) / inFileDiff * desiredDiff;
            var contents = lines.Select(l => TryParseStamps(l, out var start, out var end)
                ? $"{Skew(start).ToString(TimeSpanFormat)} --> {Skew(end).ToString(TimeSpanFormat)}"
                : l).ToArray();

            var fileNameNumber = 1;
            var backupFileName = $"{args[0]}.orig";
            while (File.Exists(backupFileName))
                backupFileName = $"{args[0]}.orig ({++fileNameNumber})";
            file.MoveTo(backupFileName);
            File.WriteAllLines(args[0], contents);

            return 0;
        }

        public static bool TryParseStamps(string text, out TimeSpan start, out TimeSpan end)
        {
            var split = text.Split(" --> ", StringSplitOptions.RemoveEmptyEntries);
            start = end = default;
            return split.Length == 2 && TryParseStamp(split[0], out start) && TryParseStamp(split[1], out end);
        }

        private const string TimeSpanFormat = @"hh\:mm\:ss\,fff";
        public static bool TryParseStamp(string text, out TimeSpan value) => TimeSpan.TryParseExact(text, TimeSpanFormat, null, out value);
    }
}

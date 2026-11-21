using System.Text;
using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// What the README promises about threads: an instance belongs to one thread at a time, and separate instances
/// share nothing, so many of them may work at once and each ends up with exactly its own data.
/// </summary>
/// <remarks>
/// The test cannot prove the absence of shared state, only look for it where it would show: sixteen threads
/// released together by a barrier each build a package, stream a workbook and read a CSV, many times over, and
/// read back what they wrote. A static cache, a shared string table that two packages reached, or a settings
/// object that one reader changed for the others would put one thread's text in another's workbook.
/// </remarks>
public sealed class SpreadsheetThreadingTests
{
    private const int Workers = 16;
    private const int Rounds = 25;

    [Fact]
    public async Task Separate_instances_used_on_many_threads_at_once_each_keep_their_own_data()
    {
        using Barrier start = new(Workers);

        Task<List<string>>[] tasks = Enumerable.Range(0, Workers)
            .Select(worker => Task.Factory.StartNew(
                () =>
                {
                    start.SignalAndWait();
                    List<string> failures = [];
                    for (int round = 0; round < Rounds; round++)
                    {
                        string text = $"worker {worker} round {round} Özlem Çelik";
                        Check(failures, "package", text, FromPackage(worker, text));
                        Check(failures, "stream writer", text, FromStreamWriter(worker, text));
                        Check(failures, "csv", text, FromCsv(text));
                    }

                    return failures;
                },
                CancellationToken.None,
                TaskCreationOptions.LongRunning,
                TaskScheduler.Default))
            .ToArray();

        List<string>[] results = await Task.WhenAll(tasks);

        Assert.All(results, failures => Assert.Empty(failures));
    }

    private static void Check(List<string> failures, string source, string expected, string actual)
    {
        if (!string.Equals(expected, actual, StringComparison.Ordinal))
        {
            failures.Add($"{source}: expected '{expected}', read '{actual}'");
        }
    }

    private static string FromPackage(int worker, string text)
    {
        byte[] bytes;
        using (SpreadsheetPackage package = new())
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add($"W{worker}");
            sheet.Cells["A1"].Value = text;
            bytes = package.GetAsByteArray();
        }

        using MemoryStream stream = new(bytes);
        using SpreadsheetWorkbookReader reader = SpreadsheetWorkbookReader.Open(stream);
        return reader.ReadRows(0).Single()[0].ToString();
    }

    private static string FromStreamWriter(int worker, string text)
    {
        using MemoryStream stream = new();
        using (SpreadsheetStreamWriter writer = new(stream, $"S{worker}"))
        {
            writer.StartRow(1);
            writer.WriteCell(1, text);
            writer.Complete();
        }

        stream.Position = 0;
        using SpreadsheetWorkbookReader reader = SpreadsheetWorkbookReader.Open(stream);
        return reader.ReadRows(0).Single()[0].ToString();
    }

    private static string FromCsv(string text)
    {
        byte[] bytes = new UTF8Encoding(false).GetBytes($"{text}\n");
        using MemoryStream stream = new(bytes);
        using SpreadsheetCsvReader reader = SpreadsheetCsvReader.Open(stream);
        return reader.ReadRows().Single()[0].ToString();
    }
}

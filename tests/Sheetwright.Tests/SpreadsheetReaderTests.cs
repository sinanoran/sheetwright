using Xunit;

namespace Sheetwright.Tests;

public sealed class SpreadsheetReaderTests
{
    private sealed record PortfolioRow(string Name, decimal Arr, DateTime InvestedDate, bool Active);

    [Fact]
    public void ReadRows_ExtractsTypedDataCorrectly()
    {
        // 1. Arrange: Write a multi-row grid workbook
        byte[] bytes;
        using (SpreadsheetPackage package = new())
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Investments");
            sheet.Cells["A1"].Value = "Company";
            sheet.Cells["B1"].Value = "ARR";
            sheet.Cells["C1"].Value = "Date";
            sheet.Cells["D1"].Value = "Active";

            sheet.Cells["A2"].Value = "Startup Alpha";
            sheet.Cells["B2"].Value = 1_200_000m;
            sheet.Cells["C2"].Value = new DateTime(2024, 5, 1);
            sheet.Cells["D2"].Value = true;

            sheet.Cells["A3"].Value = "Beta SaaS";
            sheet.Cells["B3"].Value = 850_000m;
            sheet.Cells["C3"].Value = new DateTime(2025, 2, 15);
            sheet.Cells["D3"].Value = false;

            bytes = package.GetAsByteArray();
        }

        // 2. Act: Read back using SpreadsheetReader
        using SpreadsheetReader reader = new(bytes);
        IReadOnlyList<PortfolioRow> rows = reader.ReadRows("Investments", startRow: 2, endRow: 3, row =>
            new PortfolioRow(
                row.GetText("A") ?? string.Empty,
                row.GetDecimal("B") ?? 0m,
                row.GetDateTime("C") ?? DateTime.MinValue,
                row.GetBoolean("D") ?? false
            ));

        // 3. Assert
        Assert.Equal(2, rows.Count);

        Assert.Equal("Startup Alpha", rows[0].Name);
        Assert.Equal(1_200_000m, rows[0].Arr);
        Assert.Equal(new DateTime(2024, 5, 1), rows[0].InvestedDate);
        Assert.True(rows[0].Active);

        Assert.Equal("Beta SaaS", rows[1].Name);
        Assert.Equal(850_000m, rows[1].Arr);
        Assert.Equal(new DateTime(2025, 2, 15), rows[1].InvestedDate);
        Assert.False(rows[1].Active);
    }

    [Fact]
    public void Formats_ConstantPresets_AreAvailable()
    {
        Assert.Equal("$#,##0.00;($#,##0.00);\"-\"", SpreadsheetNumberFormats.CurrencyUsd);
        Assert.Equal("0.0%", SpreadsheetNumberFormats.Percent);
        Assert.Equal("[$₺-tr-TR] #,##0.00;([$₺-tr-TR] #,##0.00);\"-\"", SpreadsheetNumberFormats.CurrencyTry);
    }
}

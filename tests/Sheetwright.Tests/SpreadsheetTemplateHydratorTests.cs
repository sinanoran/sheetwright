using Xunit;

namespace Sheetwright.Tests;

public sealed class SpreadsheetTemplateHydratorTests
{
    [Fact]
    public void HydrateTemplate_UpdatesSpecifiedCells_WhilePreservingStructure()
    {
        // 1. Arrange: Create a template workbook
        byte[] templateBytes;
        using (SpreadsheetPackage package = new())
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("PortfolioReport");
            sheet.Cells["A1"].Value = "Company Name";
            sheet.Cells["B1"].Value = "Placeholder";
            sheet.Cells["A2"].Value = "Annual Revenue";
            sheet.Cells["B2"].Value = 0D;
            sheet.Cells["A3"].Value = "Reporting Date";
            sheet.Cells["B3"].Value = new DateTime(2025, 1, 1);
            sheet.Cells["A4"].Value = "Is Audited";
            sheet.Cells["B4"].Value = false;

            // Add a defined name for B2
            package.Workbook.Names.Add("CompanyRevenue", sheet.Cells["B2"]);

            templateBytes = package.GetAsByteArray();
        }

        // 2. Act: Hydrate the template with new data
        byte[] hydratedBytes;
        using (SpreadsheetTemplateHydrator hydrator = new(templateBytes))
        {
            hydrator.SetCellValue("PortfolioReport", "B1", "Acme Fintech Corp");
            hydrator.SetNamedCellValue("CompanyRevenue", 2_500_000.50m);
            hydrator.SetCellValue("PortfolioReport!B3", new DateTime(2026, 9, 30));
            hydrator.SetCellValue("B4", true);

            hydratedBytes = hydrator.ToByteArray();
        }

        // 3. Assert: Read back the hydrated workbook using SpreadsheetReader
        using SpreadsheetReader reader = new(hydratedBytes);
        Assert.Contains("PortfolioReport", reader.SheetNames);
        Assert.Equal("Company Name", reader.GetText("PortfolioReport", "A1"));
        Assert.Equal("Acme Fintech Corp", reader.GetText("PortfolioReport", "B1"));
        Assert.Equal(2_500_000.50m, reader.GetDecimal("PortfolioReport", "B2"));
        Assert.Equal(new DateTime(2026, 9, 30), reader.GetDateTime("PortfolioReport", "B3"));
        Assert.True(reader.GetBoolean("PortfolioReport", "B4"));
    }

    [Fact]
    public void HydrateTemplate_ThrowsOnMissingDefinedName()
    {
        byte[] templateBytes;
        using (SpreadsheetPackage package = new())
        {
            package.Workbook.Worksheets.Add("Sheet1").Cells["A1"].Value = "Test";
            templateBytes = package.GetAsByteArray();
        }

        using SpreadsheetTemplateHydrator hydrator = new(templateBytes);
        Assert.Throws<KeyNotFoundException>(() => hydrator.SetNamedCellValue("NonExistent", 123));
    }
}

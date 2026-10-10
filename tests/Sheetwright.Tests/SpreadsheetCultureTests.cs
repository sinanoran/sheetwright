using System.Globalization;
using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// A workbook is a file, so it must not depend on who wrote it.
/// </summary>
/// <remarks>
/// <para>
/// This repository runs with globalization on and expects to serve readers in
/// more than one country, which makes the machine's culture a real variable
/// rather than a hypothetical one. Everything written into a package therefore
/// goes through <see cref="CultureInfo.InvariantCulture"/>: the typed numeric
/// paths always did, and the fallback that renders anything else did not, so the
/// same object produced a different file on a Turkish machine than on an
/// American one.
/// </para>
/// <para>
/// Nothing here is about how a value should look. A caller who wants a value
/// rendered for a particular reader formats it and passes a string.
/// </para>
/// </remarks>
public sealed class SpreadsheetCultureTests
{
    [Fact]
    public void A_value_with_no_special_handling_is_written_the_same_way_in_any_culture()
    {
        // DateTimeOffset takes the fallback branch — it is not one of the types
        // the writer knows — so its text is whatever Convert.ToString was given
        // a provider for.
        DateTimeOffset moment = new(2026, 8, 27, 9, 30, 0, TimeSpan.Zero);

        Assert.Equal(WriteAsText("tr-TR", moment), WriteAsText("en-US", moment));
    }

    [Fact]
    public void A_number_is_written_with_a_full_stop_in_any_culture()
    {
        // Turkish writes 1,5. A workbook that did would be read back as fifteen,
        // or refused.
        Assert.Equal("1.5", WriteAsText("tr-TR", 1.5D));
        Assert.Equal("1.5", WriteAsText("en-US", 1.5D));
    }

    [Fact]
    public void A_table_heading_taken_from_a_number_is_the_same_in_any_culture()
    {
        // The heading has to match the header cell exactly or Excel repairs the
        // file, and the cell was written invariantly.
        Assert.Equal(HeadingFromCulture("tr-TR"), HeadingFromCulture("en-US"));
    }

    [Theory]
    [InlineData("tr-TR")]
    [InlineData("de-DE")]
    [InlineData("fr-FR")]
    public void A_conditional_format_compares_with_a_full_stop_in_any_culture(string culture)
    {
        // A formula is parsed by Excel in the file's own spelling, which is invariant. A rule
        // written as "greater than 1,5" on a German machine is a syntax error everywhere.
        string[] formulas = InCulture(culture, () =>
        {
            using WrittenWorkbook written = WrittenWorkbook.From(package =>
            {
                SpreadsheetCellValueConditionalFormattingRule rule = package.Workbook.Worksheets.Add("Data")
                    .ConditionalFormatting.AddCellValue("A1:A9", SpreadsheetConditionalFormattingOperator.Between);
                rule.Formula.Value = 1.5D;
                rule.Formula2.Value = 1234.75M;
            });

            return string.Join("|", ConditionalFormattingRules(written)
                .SelectMany(rule => rule.Elements<DocumentFormat.OpenXml.Spreadsheet.Formula>())
                .Select(formula => formula.Text));
        }).Split('|');

        Assert.Equal(["1.5", "1234.75"], formulas);
    }

    [Theory]
    [InlineData("tr-TR")]
    [InlineData("de-DE")]
    public void A_colour_scale_anchor_is_written_with_a_full_stop_in_any_culture(string culture)
    {
        string anchors = InCulture(culture, () =>
        {
            using WrittenWorkbook written = WrittenWorkbook.From(package =>
            {
                SpreadsheetColorScaleConditionalFormattingRule rule = package.Workbook.Worksheets.Add("Data")
                    .ConditionalFormatting.AddThreeColorScale("A1:A9");
                rule.LowValue.SetNumber(-0.25D);
                rule.MiddleValue!.SetPercentile(37.5D);
                rule.HighValue.SetPercent(99.5D);
            });

            return string.Join(
                "|",
                ConditionalFormattingRules(written).Single()
                    .Descendants<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValueObject>()
                    .Select(anchor => anchor.Val!.Value));
        });

        Assert.Equal("-0.25|37.5|99.5", anchors);
    }

    [Fact]
    public void A_conditional_format_is_the_same_file_in_any_culture()
    {
        // Every kind of value a rule can carry, written on three machines. The rules and the
        // stylesheet that backs them have to come out byte for byte the same.
        Assert.Equal(ConditionalFormattingFromCulture("en-US"), ConditionalFormattingFromCulture("tr-TR"));
        Assert.Equal(ConditionalFormattingFromCulture("en-US"), ConditionalFormattingFromCulture("de-DE"));
    }

    private static string ConditionalFormattingFromCulture(string culture)
    {
        return InCulture(culture, () =>
        {
            using WrittenWorkbook written = WrittenWorkbook.From(package =>
            {
                SpreadsheetConditionalFormattingCollection rules = package.Workbook.Worksheets.Add("Data").ConditionalFormatting;

                SpreadsheetCellValueConditionalFormattingRule between =
                    rules.AddCellValue("A1:A9", SpreadsheetConditionalFormattingOperator.Between);
                between.Formula.Value = 0.1D;
                between.Formula2.Value = 2.5F;
                between.Style.NumberFormat.SetFormat("#,##0.00");

                rules.AddCellValue("B1:B9", SpreadsheetConditionalFormattingOperator.LessThan)
                    .Formula.Value = new DateTime(2026, 8, 27, 12, 0, 0);

                // DateTimeOffset takes the text fallback, and its text is what a culture changes.
                rules.AddCellValue("C1:C9", SpreadsheetConditionalFormattingOperator.Equal)
                    .Formula.Value = new DateTimeOffset(2026, 8, 27, 9, 30, 0, TimeSpan.Zero);

                // Excel's TRUE, not .NET's True — and not whatever a culture's casing makes of it.
                rules.AddCellValue("D1:D9", SpreadsheetConditionalFormattingOperator.Equal).Formula.Value = true;

                SpreadsheetColorScaleConditionalFormattingRule scale = rules.AddThreeColorScale("E1:E9");
                scale.LowValue.SetNumber(1.25D);
                scale.MiddleValue!.SetPercentile(62.5D);
            });

            return string.Concat(
                string.Concat(written.Sheet("Data")
                    .Elements<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting>()
                    .Select(block => block.OuterXml)),
                written.Stylesheet.DifferentialFormats!.OuterXml,
                written.Stylesheet.NumberingFormats!.OuterXml);
        });
    }

    private static IEnumerable<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingRule> ConditionalFormattingRules(
        WrittenWorkbook written) =>
        written.Sheet("Data")
            .Elements<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting>()
            .SelectMany(block => block.Elements<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingRule>());

    private static string WriteAsText(string culture, object value)
    {
        return InCulture(culture, () =>
        {
            using WrittenWorkbook written = WrittenWorkbook.From(package =>
                package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = value);

            return written.Text("Data", "A1")!;
        });
    }

    private static string HeadingFromCulture(string culture)
    {
        return InCulture(culture, () =>
        {
            using WrittenWorkbook written = WrittenWorkbook.From(package =>
            {
                SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
                sheet.Cells[1, 1].Value = 1.5D;
                sheet.Cells[2, 1].Value = "row";
                sheet.Tables.Add(sheet.Cells["A1:A2"], "Numbers");
            });

            return written.WorkbookPart.WorksheetParts.Single()
                .TableDefinitionParts.Single()
                .Table!
                .GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.TableColumns>()!
                .Elements<DocumentFormat.OpenXml.Spreadsheet.TableColumn>()
                .Single().Name!.Value!;
        });
    }

    /// <remarks>
    /// Restored in a finally rather than left set: the culture flows with the
    /// async context, and a test that abandons a changed one hands it to whatever
    /// the runner schedules on that thread next.
    /// </remarks>
    private static string InCulture(string culture, Func<string> write)
    {
        CultureInfo original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(culture);
            return write();
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }
}

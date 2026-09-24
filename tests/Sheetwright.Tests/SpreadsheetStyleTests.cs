using DocumentFormat.OpenXml.Spreadsheet;
using Drawing = System.Drawing;
using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// Styles: the three layers that combine into one, and the table that stops
/// them multiplying.
/// </summary>
/// <remarks>
/// <para>
/// A cell's format is the sheet default, then the column, then the row, then the
/// cell — each overriding the last, and each recorded as a patch of only the
/// properties somebody set. Getting the precedence wrong is invisible until a
/// column format silently wins over the cell format a caller asked for.
/// </para>
/// <para>
/// The other half is deduplication. A style table with one entry per cell is how
/// a 50,000-row export becomes a file Excel refuses to open, so identical
/// formatting must resolve to one index.
/// </para>
/// </remarks>
public sealed class SpreadsheetStyleTests
{
    [Fact]
    public void Two_cells_styled_the_same_share_one_entry_in_the_style_table()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "a";
            sheet.Cells[2, 1].Value = "b";
            sheet.Cells[1, 1].Style.Font.SetBold(true);
            sheet.Cells[2, 1].Style.Font.SetBold(true);
        });

        Assert.Equal(
            written.Cell("Data", "A1")!.StyleIndex!.Value,
            written.Cell("Data", "A2")!.StyleIndex!.Value);
    }

    [Fact]
    public void Two_cells_styled_differently_do_not()
    {
        // The negative control. Without it, a repository that returned index 0
        // for everything would pass the test above.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "a";
            sheet.Cells[2, 1].Value = "b";
            sheet.Cells[1, 1].Style.Font.SetBold(true);
            sheet.Cells[2, 1].Style.Font.SetItalic(true);
        });

        Assert.NotEqual(
            written.Cell("Data", "A1")!.StyleIndex!.Value,
            written.Cell("Data", "A2")!.StyleIndex!.Value);
    }

    [Fact]
    public void An_unstyled_cell_points_at_the_default_format()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = "plain");

        Assert.Equal(0U, written.Cell("Data", "A1")!.StyleIndex!.Value);
    }

    [Fact]
    public void A_bold_cell_gets_a_bold_font()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "heading";
            sheet.Cells[1, 1].Style.Font.SetBold(true);
        });

        Font font = FontOf(written, "A1");

        Assert.NotNull(font.Bold);
        Assert.Null(font.Italic);
    }

    [Fact]
    public void A_font_keeps_its_name_and_size()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "heading";
            sheet.Cells[1, 1].Style.Font.SetName("Segoe UI");
            sheet.Cells[1, 1].Style.Font.SetSize(14D);
        });

        Font font = FontOf(written, "A1");

        Assert.Equal("Segoe UI", font.GetFirstChild<FontName>()!.Val!.Value);
        Assert.Equal(14D, font.GetFirstChild<FontSize>()!.Val!.Value);
    }

    [Fact]
    public void A_colour_becomes_alpha_first_hexadecimal()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "red";
            sheet.Cells[1, 1].Style.Font.Color.SetColor(Drawing.Color.FromArgb(255, 255, 0, 0));
        });

        Assert.Equal("FFFF0000", FontOf(written, "A1").GetFirstChild<Color>()!.Rgb!.Value);
    }

    [Theory]
    [InlineData("FF0000", "FFFF0000")]
    [InlineData("#FF0000", "FFFF0000")]
    [InlineData("80FF0000", "80FF0000")]
    public void A_colour_written_as_text_gains_an_alpha_channel_if_it_lacks_one(string given, string expected)
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "red";
            sheet.Cells[1, 1].Style.Font.Color.SetColor(given);
        });

        Assert.Equal(expected, FontOf(written, "A1").GetFirstChild<Color>()!.Rgb!.Value);
    }

    [Fact]
    public void A_fill_is_solid_so_the_colour_is_actually_seen()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "shaded";
            sheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor("EEEEEE");
        });

        CellFormat format = written.FormatOf(written.Cell("Data", "A1")!);
        PatternFill fill = written.Stylesheet.Fills!.Elements<Fill>()
            .ElementAt((int)format.FillId!.Value).PatternFill!;

        Assert.Equal(PatternValues.Solid, fill.PatternType!.Value);
        Assert.Equal("FFEEEEEE", fill.ForegroundColor!.Rgb!.Value);
        Assert.True(format.ApplyFill!.Value);
    }

    [Fact]
    public void A_number_format_is_registered_beyond_the_builtin_range()
    {
        // Excel reserves ids below 164 for its own formats. Reusing one would
        // silently reformat every cell that referenced the builtin.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = 1234.5D;
            sheet.Cells[1, 1].Style.NumberFormat.SetFormat("#,##0.00");
        });

        CellFormat format = written.FormatOf(written.Cell("Data", "A1")!);
        NumberingFormat numbering = written.Stylesheet.NumberingFormats!
            .Elements<NumberingFormat>().Single();

        Assert.True(format.NumberFormatId!.Value >= 164U);
        Assert.Equal(format.NumberFormatId!.Value, numbering.NumberFormatId!.Value);
        Assert.Equal("#,##0.00", numbering.FormatCode!.Value);
        Assert.True(format.ApplyNumberFormat!.Value);
    }

    [Fact]
    public void The_same_number_format_asked_for_twice_is_registered_once()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = 1D;
            sheet.Cells[2, 1].Value = 2D;
            sheet.Cells[1, 1].Style.NumberFormat.SetFormat("0.00");
            sheet.Cells[2, 1].Style.NumberFormat.SetFormat("0.00");
        });

        Assert.Single(written.Stylesheet.NumberingFormats!.Elements<NumberingFormat>());
    }

    [Fact]
    public void A_date_gets_a_date_format_nobody_had_to_ask_for()
    {
        // Without one, Excel shows the serial number and the export looks broken.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = new DateTime(2026, 8, 27));

        CellFormat format = written.FormatOf(written.Cell("Data", "A1")!);
        NumberingFormat numbering = written.Stylesheet.NumberingFormats!
            .Elements<NumberingFormat>().Single();

        Assert.Equal(format.NumberFormatId!.Value, numbering.NumberFormatId!.Value);
        Assert.Equal("dd/mm/yyyy", numbering.FormatCode!.Value);
    }

    [Fact]
    public void An_explicit_format_on_a_date_wins_over_the_implicit_one()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = new DateTime(2026, 8, 27);
            sheet.Cells[1, 1].Style.NumberFormat.SetFormat("yyyy-mm-dd");
        });

        Assert.Equal(
            "yyyy-mm-dd",
            written.Stylesheet.NumberingFormats!.Elements<NumberingFormat>().Single().FormatCode!.Value);
    }

    [Fact]
    public void A_date_replaced_by_a_number_stops_being_formatted_as_a_date()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = new DateTime(2026, 8, 27);
            sheet.Cells[1, 1].Value = 42;
        });

        Assert.Empty(written.Stylesheet.NumberingFormats!.Elements<NumberingFormat>());
        Assert.Equal(0U, written.Cell("Data", "A1")!.StyleIndex!.Value);
    }

    [Fact]
    public void A_date_read_back_is_a_date_and_not_its_serial_number()
    {
        DateTime moment = new(2026, 8, 27, 9, 30, 0);

        byte[] bytes;
        using (SpreadsheetPackage package = new())
        {
            package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = moment;
            bytes = package.GetAsByteArray();
        }

        using SpreadsheetPackage reopened = new(new MemoryStream(bytes));

        Assert.Equal(moment, reopened.Workbook.Worksheets["Data"]!.Cells[1, 1].Value);
    }

    [Fact]
    public void A_cell_style_beats_the_column_style_it_sits_in()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = 1D;
            sheet.Column(1).Style.NumberFormat.SetFormat("0.00");
            sheet.Cells[1, 1].Style.NumberFormat.SetFormat("0.0000");
        });

        CellFormat format = written.FormatOf(written.Cell("Data", "A1")!);
        NumberingFormat numbering = written.Stylesheet.NumberingFormats!
            .Elements<NumberingFormat>()
            .Single(x => x.NumberFormatId!.Value == format.NumberFormatId!.Value);

        Assert.Equal("0.0000", numbering.FormatCode!.Value);
    }

    [Fact]
    public void A_column_style_reaches_the_cells_that_were_never_styled()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = 1D;
            sheet.Column(1).Style.NumberFormat.SetFormat("0.00");
        });

        CellFormat format = written.FormatOf(written.Cell("Data", "A1")!);

        Assert.True(format.NumberFormatId!.Value >= 164U);
    }

    [Fact]
    public void A_worksheet_default_reaches_every_cell()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells.Style.Font.SetName("Segoe UI");
            sheet.Cells[1, 1].Value = "a";
            sheet.Cells[9, 9].Value = "b";
        });

        Assert.Equal("Segoe UI", FontOf(written, "A1").GetFirstChild<FontName>()!.Val!.Value);
        Assert.Equal("Segoe UI", FontOf(written, "I9").GetFirstChild<FontName>()!.Val!.Value);
    }

    [Fact]
    public void A_row_style_reaches_every_cell_in_that_row_and_no_other()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "a";
            sheet.Cells[1, 2].Value = "b";
            sheet.Cells[2, 1].Value = "c";
            sheet.Cells["1:1"].Style.Font.SetBold(true);
        });

        Assert.NotNull(FontOf(written, "A1").Bold);
        Assert.NotNull(FontOf(written, "B1").Bold);
        Assert.Equal(0U, written.Cell("Data", "A2")!.StyleIndex!.Value);
    }

    [Fact]
    public void Alignment_and_wrapping_are_applied_together()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "a long value";
            sheet.Cells[1, 1].Style.SetHorizontalAlignment(SpreadsheetHorizontalAlignment.Center);
            sheet.Cells[1, 1].Style.SetWrapText(true);
        });

        CellFormat format = written.FormatOf(written.Cell("Data", "A1")!);

        Assert.True(format.ApplyAlignment!.Value);
        Assert.Equal(HorizontalAlignmentValues.Center, format.Alignment!.Horizontal!.Value);
        Assert.True(format.Alignment!.WrapText!.Value);
    }

    [Fact]
    public void Unlocking_a_cell_is_what_makes_sheet_protection_useful()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "editable";
            sheet.Cells[1, 1].Style.SetLocked(false);
            sheet.Protect(lockEverything: true, password: "pw");
        });

        CellFormat format = written.FormatOf(written.Cell("Data", "A1")!);

        Assert.True(format.ApplyProtection!.Value);
        Assert.False(format.Protection!.Locked!.Value);
    }

    [Fact]
    public void A_border_around_a_range_is_drawn_only_on_the_outside()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            for (int row = 1; row <= 2; row++)
            {
                for (int column = 1; column <= 2; column++)
                {
                    sheet.Cells[row, column].Value = "x";
                }
            }

            sheet.Cells["A1:B2"].Style.Border.BorderAround(SpreadsheetBorderStyle.Medium, Drawing.Color.Black);
        });

        Border topLeft = BorderOf(written, "A1");
        Border bottomRight = BorderOf(written, "B2");

        Assert.Equal(BorderStyleValues.Medium, topLeft.LeftBorder!.Style!.Value);
        Assert.Equal(BorderStyleValues.Medium, topLeft.TopBorder!.Style!.Value);
        Assert.Equal(BorderStyleValues.None, topLeft.RightBorder!.Style!.Value);
        Assert.Equal(BorderStyleValues.None, topLeft.BottomBorder!.Style!.Value);

        Assert.Equal(BorderStyleValues.Medium, bottomRight.RightBorder!.Style!.Value);
        Assert.Equal(BorderStyleValues.Medium, bottomRight.BottomBorder!.Style!.Value);
        Assert.Equal(BorderStyleValues.None, bottomRight.LeftBorder!.Style!.Value);
    }

    [Fact]
    public void A_border_colour_is_only_written_where_there_is_a_border()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            sheet.Cells[1, 2].Value = "y";
            sheet.Cells["A1:B1"].Style.Border.BorderAround(SpreadsheetBorderStyle.Thin, Drawing.Color.Red);
        });

        Border border = BorderOf(written, "A1");

        Assert.Equal("FFFF0000", border.LeftBorder!.Color!.Rgb!.Value);
        Assert.Null(border.RightBorder!.Color);
    }

    [Fact]
    public void A_border_around_a_whole_column_is_refused_rather_than_drawn_a_million_times()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        Assert.Throws<InvalidOperationException>(
            () => sheet.Cells["A:A"].Style.Border.BorderAround(SpreadsheetBorderStyle.Thin, Drawing.Color.Black));
    }

    [Fact]
    public void Styling_a_cell_nobody_wrote_to_still_produces_the_cell()
    {
        // Otherwise a blank cell with a background colour is simply not there.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[3, 3].Style.Fill.BackgroundColor.SetColor("EEEEEE");
        });

        Assert.NotNull(written.Cell("Data", "C3"));
        Assert.NotEqual(0U, written.Cell("Data", "C3")!.StyleIndex!.Value);
    }

    private static Font FontOf(WrittenWorkbook written, string reference)
    {
        CellFormat format = written.FormatOf(written.Cell("Data", reference)!);
        return written.Stylesheet.Fonts!.Elements<Font>().ElementAt((int)format.FontId!.Value);
    }

    private static Border BorderOf(WrittenWorkbook written, string reference)
    {
        CellFormat format = written.FormatOf(written.Cell("Data", reference)!);
        return written.Stylesheet.Borders!.Elements<Border>().ElementAt((int)format.BorderId!.Value);
    }
}

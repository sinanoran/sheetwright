using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// The grid export: a list of rows, as a workbook somebody can open.
/// </summary>
/// <remarks>
/// This is the shape both callers in the repository use — the member directory
/// and the audit trail — so what it defends is what an export is for. A date that
/// arrives as text sorts April before January; a number as text cannot be summed;
/// a heading row that scrolls away makes a long export unreadable. None of those
/// fails loudly, and all three are the reason somebody asked for Excel rather
/// than CSV.
/// </remarks>
public sealed class SpreadsheetGridTests
{
    private sealed record Row(string Name, int Count, DateTime When);

    private static readonly Row[] TwoRows =
    [
        new("Ada", 3, new DateTime(2026, 1, 15, 9, 30, 0)),
        new("Grace", 7, new DateTime(2026, 2, 20, 14, 0, 0))
    ];

    private static readonly SpreadsheetGridColumn<Row>[] Columns =
    [
        new("Name", row => row.Name),
        new("Count", row => row.Count),
        new("When", row => row.When) { NumberFormat = "yyyy-mm-dd" }
    ];

    [Fact]
    public void The_headings_are_the_first_row()
    {
        using WrittenWorkbook written = Write();

        Assert.Equal("Name", written.Text("Report", "A1"));
        Assert.Equal("Count", written.Text("Report", "B1"));
        Assert.Equal("When", written.Text("Report", "C1"));
    }

    [Fact]
    public void The_headings_are_bold_so_they_read_as_headings()
    {
        using WrittenWorkbook written = Write();

        CellFormat format = written.FormatOf(written.Cell("Report", "A1")!);
        Font font = written.Stylesheet.Fonts!.Elements<Font>().ElementAt((int)format.FontId!.Value);

        Assert.NotNull(font.Bold);
    }

    [Fact]
    public void The_rows_follow_the_headings_in_the_order_they_were_given()
    {
        using WrittenWorkbook written = Write();

        Assert.Equal("Ada", written.Text("Report", "A2"));
        Assert.Equal("Grace", written.Text("Report", "A3"));
    }

    [Fact]
    public void A_number_is_written_as_a_number_and_not_as_text()
    {
        // Otherwise the column cannot be summed, which is most of the point.
        using WrittenWorkbook written = Write();

        Cell cell = written.Cell("Report", "B2")!;

        Assert.Equal(CellValues.Number, cell.DataType!.Value);
        Assert.Equal("3", cell.CellValue!.InnerText);
    }

    [Fact]
    public void A_date_is_written_as_a_date_and_carries_the_columns_format()
    {
        using WrittenWorkbook written = Write();

        Cell cell = written.Cell("Report", "C2")!;
        CellFormat format = written.FormatOf(cell);
        NumberingFormat numbering = written.Stylesheet.NumberingFormats!
            .Elements<NumberingFormat>()
            .Single(x => x.NumberFormatId!.Value == format.NumberFormatId!.Value);

        Assert.Equal(CellValues.Number, cell.DataType!.Value);
        Assert.Equal("yyyy-mm-dd", numbering.FormatCode!.Value);
    }

    [Fact]
    public void A_column_with_no_format_is_left_unformatted()
    {
        // The negative control for the format above: a number format applied to
        // every column would silently reformat the text ones.
        using WrittenWorkbook written = Write();

        Assert.Equal(
            "General",
            FormatCodeOf(written, "A2") ?? "General");
    }

    [Fact]
    public void The_heading_row_is_frozen_so_it_survives_scrolling()
    {
        using WrittenWorkbook written = Write();

        Pane pane = written.Sheet("Report").GetFirstChild<SheetViews>()!
            .Elements<SheetView>().Single().Pane!;

        Assert.Equal(1D, pane.VerticalSplit!.Value);
        Assert.Equal(PaneStateValues.Frozen, pane.State!.Value);
    }

    [Fact]
    public void The_columns_are_wide_enough_for_what_is_in_them()
    {
        using WrittenWorkbook written = Write();

        Column[] columns = [.. written.Sheet("Report").GetFirstChild<Columns>()!.Elements<Column>()];

        Assert.Equal(3, columns.Length);
        Assert.All(columns, column => Assert.True(column.Width!.Value > 0D));
    }

    [Fact]
    public void A_heading_longer_than_its_data_still_fits()
    {
        // Auto-fit runs after the headings are written, which is the only reason
        // this holds — measuring before would size the column to "3".
        byte[] bytes = SpreadsheetGrid.Write(
            "Report",
            [new SpreadsheetGridColumn<Row>("A considerably long heading", row => row.Count)],
            TwoRows);

        using WrittenWorkbook written = Open(bytes);
        Column column = written.Sheet("Report").GetFirstChild<Columns>()!.Elements<Column>().Single();

        Assert.True(column.Width!.Value >= "A considerably long heading".Length);
    }

    [Fact]
    public void An_empty_grid_is_a_workbook_with_headings_and_no_rows()
    {
        // What an export of a filtered-to-nothing grid produces. An empty file,
        // or a throw, would both read as a failure of the export rather than an
        // answer to the question.
        byte[] bytes = SpreadsheetGrid.Write("Report", Columns, Array.Empty<Row>());

        using WrittenWorkbook written = Open(bytes);

        Assert.Equal("Name", written.Text("Report", "A1"));
        Assert.Null(written.Cell("Report", "A2"));
    }

    [Fact]
    public void A_null_value_leaves_its_cell_empty()
    {
        byte[] bytes = SpreadsheetGrid.Write(
            "Report",
            [new SpreadsheetGridColumn<Row>("Name", _ => null)],
            TwoRows);

        using WrittenWorkbook written = Open(bytes);

        Assert.Null(written.Text("Report", "A2"));
    }

    [Fact]
    public void A_grid_with_no_columns_is_refused()
    {
        Assert.Throws<ArgumentException>(
            () => SpreadsheetGrid.Write("Report", Array.Empty<SpreadsheetGridColumn<Row>>(), TwoRows));
    }

    [Fact]
    public void Nulls_are_refused()
    {
        Assert.Throws<ArgumentNullException>(() => SpreadsheetGrid.Write(null!, Columns, TwoRows));
        Assert.Throws<ArgumentNullException>(() => SpreadsheetGrid.Write<Row>("Report", null!, TwoRows));
        Assert.Throws<ArgumentNullException>(() => SpreadsheetGrid.Write("Report", Columns, null!));
    }

    [Fact]
    public void The_media_type_is_the_one_Excel_is_registered_against()
    {
        // Hard to notice when wrong: the browser downloads a file that then will
        // not open, and it reads as a corrupt export.
        Assert.Equal(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            SpreadsheetGrid.ContentType);
    }

    private static WrittenWorkbook Write() => Open(SpreadsheetGrid.Write("Report", Columns, TwoRows));

    private static WrittenWorkbook Open(byte[] bytes) => WrittenWorkbook.Opening(bytes);

    private static string? FormatCodeOf(WrittenWorkbook written, string reference)
    {
        CellFormat format = written.FormatOf(written.Cell("Report", reference)!);

        return written.Stylesheet.NumberingFormats?
            .Elements<NumberingFormat>()
            .FirstOrDefault(x => x.NumberFormatId!.Value == format.NumberFormatId?.Value)?
            .FormatCode?.Value;
    }
}

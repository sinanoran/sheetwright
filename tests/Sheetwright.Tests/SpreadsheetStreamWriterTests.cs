using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// The streaming writer: a workbook written one row at a time, with every part assembled by hand.
/// </summary>
/// <remarks>
/// <para>
/// This is the one writer in the library that does not go through
/// <see cref="DocumentFormat.OpenXml"/> at all — the worksheet, the styles, the table parts, the
/// relationships and the content types are all written as XML text. Nothing checks the schema on
/// the way out, so these tests check it on the way back in: every workbook written here is run
/// through <see cref="OpenXmlValidator"/> as well as read.
/// </para>
/// <para>
/// The other thing being defended is the agreement between a table and its header row. The
/// worksheet writer can settle that after the fact, because it still holds every cell when it
/// writes the table part. Here the header row has been written and thrown away long before, so the
/// headings are resolved up front and the header cells are written from them.
/// </para>
/// </remarks>
public sealed class SpreadsheetStreamWriterTests
{
    [Fact]
    public void A_streamed_cell_is_typed_as_the_package_types_the_same_value()
    {
        DateTime moment = new(2026, 3, 4, 5, 6, 7, DateTimeKind.Unspecified);

        byte[] bytes = Streamed(writer =>
        {
            writer.StartRow(1);
            writer.WriteCell(1, "text");
            writer.WriteCell(2, 42);
            writer.WriteCell(3, 3.5D);
            writer.WriteCell(4, 12.75M);
            writer.WriteCell(5, true);
            writer.WriteCell(6, false);
            writer.WriteCell(7, moment);
            writer.WriteCell(8, 'x');
            writer.WriteCell(9, null);
            writer.WriteCell(10, DBNull.Value);
        });

        using SpreadsheetPackage reopened = new(new MemoryStream(bytes));
        SpreadsheetWorksheet read = reopened.Workbook.Worksheets["Data"]!;

        Assert.Equal("text", read.Cells[1, 1].Value);
        Assert.Equal(42D, read.Cells[1, 2].Value);
        Assert.Equal(3.5D, read.Cells[1, 3].Value);
        Assert.Equal(12.75D, read.Cells[1, 4].Value);
        Assert.Equal(true, read.Cells[1, 5].Value);
        Assert.Equal(false, read.Cells[1, 6].Value);
        Assert.Equal(moment, read.Cells[1, 7].Value);
        Assert.Equal("x", read.Cells[1, 8].Value);
        Assert.Null(read.Cells[1, 9].Value);
        Assert.Null(read.Cells[1, 10].Value);
    }

    [Fact]
    public void A_date_with_no_format_of_its_own_gets_one_as_it_does_everywhere_else()
    {
        // Unformatted, a date cell shows Excel's serial number and the export looks broken, so
        // SpreadsheetWorksheet gives one dd/mm/yyyy. The streamed writer has to agree, or the same
        // rows produce a different file depending on which writer wrote them.
        using WrittenWorkbook written = StreamedWorkbook(writer =>
        {
            uint hours = writer.GetStyleIndex(numberFormat: "yyyy-mm-dd hh:mm");
            writer.SetColumn(3, width: null, styleIndex: writer.GetStyleIndex(numberFormat: "yyyy"));

            writer.StartRow(1);
            writer.WriteCell(1, new DateTime(2026, 3, 4, 5, 6, 7, DateTimeKind.Unspecified));
            writer.WriteCell(2, new DateTime(2026, 3, 4, 5, 6, 7, DateTimeKind.Unspecified), hours);
            writer.WriteCell(3, new DateTime(2026, 3, 4, 5, 6, 7, DateTimeKind.Unspecified));
        });

        Assert.Equal("dd/mm/yyyy", FormatCodeOf(written, "A1"));
        Assert.Equal("yyyy-mm-dd hh:mm", FormatCodeOf(written, "B1"));

        // The column already formats C1, so nothing is imposed on the cell and the column's
        // format is what the reader sees.
        Assert.Null(written.Cell("Data", "C1")!.StyleIndex);
    }

    [Fact]
    public void Text_is_written_inline_because_there_is_no_shared_string_table_to_hold_open()
    {
        using WrittenWorkbook written = StreamedWorkbook(writer =>
        {
            writer.StartRow(1);
            writer.WriteCell(1, "Reference");
            writer.WriteCell(2, "  padded  ");
        });

        Assert.Empty(written.SharedStrings);
        Assert.Null(written.WorkbookPart.SharedStringTablePart);
        Assert.Equal("Reference", written.Text("Data", "A1"));

        // Excel drops the padding without xml:space, and a column of codes that were deliberately
        // padded is a column that no longer joins to anything.
        Assert.Equal("  padded  ", written.Text("Data", "B1"));
    }

    [Fact]
    public void A_table_heading_that_had_to_be_invented_is_written_into_the_header_cell_too()
    {
        // The same guarantee SpreadsheetWorksheet gives: Excel checks a table's column names
        // against the cells of its header row and offers to repair the file when they disagree.
        using WrittenWorkbook written = StreamedWorkbook(writer =>
        {
            writer.AddTable(1, 1, 2, 3, "Invoices", ["Amount", "Amount", "  "]);

            writer.StartRow(1);
            writer.WriteCell(1, "Amount");
            writer.WriteCell(2, "Amount");
            writer.WriteCell(3, "  ");
            writer.StartRow(2);
            writer.WriteCell(1, 1D);
        });

        Table table = TableOf(written);

        Assert.Equal<string[]>(
            ["Amount", "Amount2", "C"],
            [.. table.GetFirstChild<TableColumns>()!.Elements<TableColumn>().Select(x => x.Name!.Value!)]);

        Assert.Equal("Amount", written.Text("Data", "A1"));
        Assert.Equal("Amount2", written.Text("Data", "B1"));
        Assert.Equal("C", written.Text("Data", "C1"));
    }

    [Fact]
    public void A_header_cell_the_caller_never_wrote_is_written_anyway()
    {
        // A header row with a hole in it is a table naming a column that is not there, which is
        // the same repair dialogue by another route.
        using WrittenWorkbook written = StreamedWorkbook(writer =>
        {
            writer.AddTable(1, 1, 2, 3, "Invoices", ["Reference", "Issued", "Amount"]);

            writer.StartRow(1);
            writer.WriteCell(3, "Amount");
            writer.StartRow(2);
            writer.WriteCell(1, "INV-1");
        });

        Assert.Equal("Reference", written.Text("Data", "A1"));
        Assert.Equal("Issued", written.Text("Data", "B1"));
        Assert.Equal("Amount", written.Text("Data", "C1"));

        Assert.Equal<string[]>(
            ["A1", "B1", "C1"],
            [.. written.Sheet("Data").Descendants<Row>().First().Elements<Cell>().Select(x => x.CellReference!.Value!)]);
    }

    [Fact]
    public void A_table_is_written_with_the_autofilter_and_style_flags_the_worksheet_writer_uses()
    {
        using WrittenWorkbook streamed = StreamedWorkbook(writer =>
        {
            writer.AddTable(1, 1, 2, 2, "Invoices", ["Reference", "Amount"]);
            writer.StartRow(1);
            writer.WriteCell(1, "Reference");
            writer.WriteCell(2, "Amount");
            writer.StartRow(2);
            writer.WriteCell(1, "INV-1");
            writer.WriteCell(2, 10D);
        });

        using WrittenWorkbook packaged = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "Reference";
            sheet.Cells[1, 2].Value = "Amount";
            sheet.Cells[2, 1].Value = "INV-1";
            sheet.Cells[2, 2].Value = 10D;
            sheet.Tables.Add(sheet.Cells["A1:B2"], "Invoices");
        });

        Table streamedTable = TableOf(streamed);
        Table packagedTable = TableOf(packaged);

        Assert.Equal(packagedTable.Reference!.Value, streamedTable.Reference!.Value);
        Assert.Equal(packagedTable.AutoFilter!.Reference!.Value, streamedTable.AutoFilter!.Reference!.Value);
        Assert.Equal(packagedTable.Name!.Value, streamedTable.Name!.Value);
        Assert.Equal(packagedTable.DisplayName!.Value, streamedTable.DisplayName!.Value);

        TableStyleInfo streamedStyle = streamedTable.GetFirstChild<TableStyleInfo>()!;
        TableStyleInfo packagedStyle = packagedTable.GetFirstChild<TableStyleInfo>()!;

        Assert.Equal(packagedStyle.Name!.Value, streamedStyle.Name!.Value);
        Assert.Equal(packagedStyle.ShowRowStripes!.Value, streamedStyle.ShowRowStripes!.Value);
        Assert.Equal(packagedStyle.ShowColumnStripes!.Value, streamedStyle.ShowColumnStripes!.Value);
        Assert.Equal(packagedStyle.ShowFirstColumn!.Value, streamedStyle.ShowFirstColumn!.Value);
        Assert.Equal(packagedStyle.ShowLastColumn!.Value, streamedStyle.ShowLastColumn!.Value);
    }

    [Fact]
    public void Two_tables_each_get_their_own_part_and_relationship()
    {
        using WrittenWorkbook written = StreamedWorkbook(writer =>
        {
            writer.AddTable(1, 1, 2, 1, "First", ["Reference"]);
            writer.AddTable(4, 1, 5, 1, "Second", ["Total"]);

            writer.StartRow(1);
            writer.WriteCell(1, "Reference");
            writer.StartRow(2);
            writer.WriteCell(1, "INV-1");
            writer.StartRow(4);
            writer.WriteCell(1, "Total");
            writer.StartRow(5);
            writer.WriteCell(1, 1D);
        });

        WorksheetPart part = (WorksheetPart)written.WorkbookPart
            .GetPartById(written.Workbook.Sheets!.Elements<Sheet>().Single().Id!.Value!);

        Assert.Equal(2, part.TableDefinitionParts.Count());
        Assert.Equal<string[]>(
            ["A1:A2", "A4:A5"],
            [.. part.TableDefinitionParts.Select(x => x.Table!.Reference!.Value!).Order()]);
        Assert.Equal<uint[]>([1, 2], [.. part.TableDefinitionParts.Select(x => x.Table!.Id!.Value).Order()]);
    }

    [Fact]
    public void A_table_left_open_at_the_bottom_ends_at_the_last_row_written()
    {
        // The row count is the one thing a streaming caller usually does not have: the rows are a
        // query nobody counted first. The table part is written at Complete, so it does not have
        // to be known in advance.
        using WrittenWorkbook written = StreamedWorkbook(writer =>
        {
            writer.AddTable(1, 1, null, 2, "Invoices", ["Reference", "Amount"]);

            writer.StartRow(1);
            writer.WriteCell(1, "Reference");
            writer.WriteCell(2, "Amount");

            for (int row = 2; row <= 4; row++)
            {
                writer.StartRow(row);
                writer.WriteCell(1, $"INV-{row}");
                writer.WriteCell(2, (double)row);
            }
        });

        Table table = TableOf(written);

        Assert.Equal("A1:B4", table.Reference!.Value);
        Assert.Equal("A1:B4", table.AutoFilter!.Reference!.Value);
    }

    [Fact]
    public void A_table_left_open_over_a_header_row_and_nothing_else_is_the_header_row()
    {
        // An export of an empty result still has to open, and a range that ends above where it
        // starts is not one Excel can read.
        using WrittenWorkbook written = StreamedWorkbook(writer =>
        {
            writer.AddTable(1, 1, null, 2, "Invoices", ["Reference", "Amount"]);

            writer.StartRow(1);
            writer.WriteCell(1, "Reference");
            writer.WriteCell(2, "Amount");
        });

        Assert.Equal("A1:B1", TableOf(written).Reference!.Value);
    }

    [Fact]
    public void Bytes_reach_the_destination_while_the_rows_are_still_being_written()
    {
        // This is the whole point of the class: the rows go out as they arrive rather than being
        // held until the end. A workbook that is still being written is already mostly written.
        using MemoryStream destination = new();
        long lengthBeforeCompleting;

        using (SpreadsheetStreamWriter writer = new(destination, "Data"))
        {
            writer.StartRow(1);
            writer.WriteCell(1, "Reference");

            for (int row = 2; row <= 50_000; row++)
            {
                writer.StartRow(row);
                writer.WriteCell(1, $"INV-{row}");
            }

            lengthBeforeCompleting = destination.Length;
            writer.Complete();
        }

        Assert.True(
            lengthBeforeCompleting > 0L,
            "Nothing had reached the destination stream by the fifty-thousandth row.");

        Assert.True(
            lengthBeforeCompleting > destination.Length / 2L,
            $"Only {lengthBeforeCompleting.ToString()} of {destination.Length.ToString()} bytes had been written by the fifty-thousandth row.");
    }

    [Fact]
    public void Identical_formatting_resolves_to_one_entry_in_the_style_table()
    {
        using WrittenWorkbook written = StreamedWorkbook(writer =>
        {
            uint bold = writer.GetStyleIndex(bold: true);
            Assert.Equal(bold, writer.GetStyleIndex(bold: true));
            Assert.NotEqual(bold, writer.GetStyleIndex(bold: true, italic: true));

            writer.StartRow(1);
            writer.WriteCell(1, "Reference", bold);
            writer.WriteCell(2, "Notes", writer.GetStyleIndex(wrapText: true));
        });

        CellFormat[] formats = [.. written.Stylesheet.CellFormats!.Elements<CellFormat>()];

        // The default, bold, bold-italic, wrapped.
        Assert.Equal(4, formats.Length);
        Assert.Equal(4U, written.Stylesheet.CellFormats!.Count!.Value);

        Font boldFont = (Font)written.Stylesheet.Fonts!.ElementAt((int)formats[1].FontId!.Value);
        Assert.NotNull(boldFont.Bold);
        Assert.Null(boldFont.Italic);

        Assert.True(written.FormatOf(written.Cell("Data", "B1")!).Alignment!.WrapText!.Value);
    }

    [Fact]
    public void A_column_carries_its_width_visibility_and_style()
    {
        using WrittenWorkbook written = StreamedWorkbook(writer =>
        {
            writer.SetColumn(1, width: 30D);
            writer.SetColumn(3, width: null, hidden: true);
            writer.SetColumn(2, width: null, styleIndex: writer.GetStyleIndex(italic: true));

            writer.StartRow(1);
            writer.WriteCell(1, "Reference");
        });

        Column[] columns = [.. written.Sheet("Data").GetFirstChild<Columns>()!.Elements<Column>()];

        // Ascending, whatever order they were set in.
        Assert.Equal<uint[]>([1, 2, 3], [.. columns.Select(x => x.Min!.Value)]);
        Assert.Equal(30D, columns[0].Width!.Value);
        Assert.True(columns[0].CustomWidth!.Value);
        Assert.Null(columns[0].Hidden);
        Assert.NotEqual(0U, columns[1].Style!.Value);
        Assert.True(columns[2].Hidden!.Value);
    }

    [Fact]
    public void NaN_and_the_infinities_go_in_as_text_because_the_format_has_no_number_for_them()
    {
        using WrittenWorkbook written = StreamedWorkbook(writer =>
        {
            writer.StartRow(1);
            writer.WriteCell(1, double.NaN);
            writer.WriteCell(2, double.PositiveInfinity);
            writer.WriteCell(3, float.NegativeInfinity);
        });

        Assert.Equal("NaN", written.Text("Data", "A1"));
        Assert.Equal("Infinity", written.Text("Data", "B1"));
        Assert.Equal("-Infinity", written.Text("Data", "C1"));
    }

    [Fact]
    public void The_content_types_are_the_first_entry_in_the_archive()
    {
        // Excel reads the package sequentially and wants to know what a part is before it reaches
        // it. SpreadsheetPackageArchiveWriter puts them first for the same reason.
        byte[] bytes = Streamed(writer =>
        {
            writer.AddTable(1, 1, 2, 1, "Invoices", ["Reference"]);
            writer.StartRow(1);
            writer.WriteCell(1, "Reference");
            writer.StartRow(2);
            writer.WriteCell(1, "INV-1");
        });

        using ZipArchive archive = new(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);

        Assert.Equal("[Content_Types].xml", archive.Entries[0].FullName);
        Assert.Contains("xl/tables/table1.xml", archive.Entries.Select(x => x.FullName));
    }

    [Fact]
    public void The_author_lands_in_the_core_properties()
    {
        byte[] bytes = Streamed(
            writer =>
            {
                writer.StartRow(1);
                writer.WriteCell(1, "Reference");
            },
            author: "Sheetwright");

        using MemoryStream stream = new(bytes, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        Assert.Equal("Sheetwright", document.PackageProperties.Creator);
    }

    [Fact]
    public void Tens_of_thousands_of_rows_come_out_in_order()
    {
        const int rowCount = 20_000;

        byte[] bytes = Streamed(writer =>
        {
            writer.AddTable(1, 1, rowCount + 1, 2, "Rows", ["Id", "Name"]);

            writer.StartRow(1);
            writer.WriteCell(1, "Id");
            writer.WriteCell(2, "Name");

            for (int row = 1; row <= rowCount; row++)
            {
                writer.StartRow(row + 1);
                writer.WriteCell(1, row);
                writer.WriteCell(2, $"Row {row}");
            }
        });

        using SpreadsheetPackage reopened = new(new MemoryStream(bytes));
        SpreadsheetWorksheet read = reopened.Workbook.Worksheets["Data"]!;

        Assert.Equal(1D, read.Cells[2, 1].Value);
        Assert.Equal("Row 1", read.Cells[2, 2].Value);
        Assert.Equal((double)rowCount, read.Cells[rowCount + 1, 1].Value);
        Assert.Equal($"Row {rowCount}", read.Cells[rowCount + 1, 2].Value);
    }

    [Fact]
    public void The_destination_stream_is_left_open_for_the_caller_who_owns_it()
    {
        using MemoryStream destination = new();
        using (SpreadsheetStreamWriter writer = new(destination, "Data"))
        {
            writer.StartRow(1);
            writer.WriteCell(1, "Reference");
            writer.Complete();
        }

        Assert.True(destination.CanWrite);
        Assert.NotEqual(0L, destination.Length);
    }

    [Fact]
    public void Rows_and_cells_out_of_order_are_refused_rather_than_written()
    {
        // Nothing downstream notices: a row written out of order produces a package Excel offers
        // to repair, and there is no second pass here to put it right.
        using MemoryStream destination = new();
        using SpreadsheetStreamWriter writer = new(destination, "Data");

        writer.StartRow(2);
        writer.WriteCell(2, "b");

        Assert.Throws<InvalidOperationException>(() => writer.StartRow(2));
        Assert.Throws<InvalidOperationException>(() => writer.StartRow(1));
        Assert.Throws<InvalidOperationException>(() => writer.WriteCell(2, "again"));
        Assert.Throws<InvalidOperationException>(() => writer.WriteCell(1, "before"));
    }

    [Fact]
    public void A_cell_outside_the_grid_is_refused_as_it_is_everywhere_else()
    {
        using MemoryStream destination = new();
        using SpreadsheetStreamWriter writer = new(destination, "Data");

        Assert.Throws<ArgumentOutOfRangeException>(() => writer.StartRow(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => writer.StartRow(1_048_577));
        Assert.Throws<ArgumentOutOfRangeException>(() => writer.SetColumn(0, 10D));
        Assert.Throws<ArgumentOutOfRangeException>(() => writer.SetColumn(16_385, 10D));
        Assert.Throws<ArgumentOutOfRangeException>(() => writer.SetColumn(1, -1D));
        Assert.Throws<ArgumentOutOfRangeException>(() => writer.SetColumn(1, 10D, styleIndex: 7));
        Assert.Throws<ArgumentOutOfRangeException>(() => writer.AddTable(0, 1, 2, 1, "T", ["a"]));
        Assert.Throws<ArgumentOutOfRangeException>(() => writer.AddTable(2, 1, 1, 1, "T", ["a"]));

        writer.StartRow(1);
        Assert.Throws<ArgumentOutOfRangeException>(() => writer.WriteCell(16_385, "x"));
        Assert.Throws<ArgumentOutOfRangeException>(() => writer.WriteCell(1, "x", styleIndex: 7));
    }

    [Fact]
    public void A_table_needs_one_heading_per_column()
    {
        using MemoryStream destination = new();
        using SpreadsheetStreamWriter writer = new(destination, "Data");

        Assert.Throws<ArgumentException>(() => writer.AddTable(1, 1, 2, 3, "T", ["a", "b"]));
        Assert.Throws<ArgumentNullException>(() => writer.AddTable(1, 1, 2, 1, "T", null!));
        Assert.Throws<ArgumentNullException>(() => writer.AddTable(1, 1, 2, 1, null!, ["a"]));
    }

    [Fact]
    public void Columns_and_tables_cannot_be_added_once_a_row_has_been_written()
    {
        // Both are written ahead of the rows, so by here the bytes are already gone. Throwing is
        // the only honest answer; accepting the call and dropping it is the one that wastes an
        // afternoon.
        using MemoryStream destination = new();
        using SpreadsheetStreamWriter writer = new(destination, "Data");

        writer.StartRow(1);
        writer.WriteCell(1, "Reference");

        Assert.Throws<InvalidOperationException>(() => writer.SetColumn(1, 10D));
        Assert.Throws<InvalidOperationException>(() => writer.AddTable(1, 1, 2, 1, "T", ["a"]));
    }

    [Fact]
    public void A_table_whose_header_row_was_never_written_is_refused_at_completion()
    {
        using MemoryStream destination = new();
        using SpreadsheetStreamWriter writer = new(destination, "Data");

        writer.AddTable(4, 1, 5, 1, "Invoices", ["Reference"]);
        writer.StartRow(1);
        writer.WriteCell(1, "Reference");

        // Nothing can be put right by now, and the alternative is a table naming a column that is
        // not in the sheet: a file whoever opens it is asked to repair.
        Assert.Throws<InvalidOperationException>(writer.Complete);
    }

    [Fact]
    public void A_cell_before_the_first_row_is_refused()
    {
        using MemoryStream destination = new();
        using SpreadsheetStreamWriter writer = new(destination, "Data");

        Assert.Throws<InvalidOperationException>(() => writer.WriteCell(1, "Reference"));
    }

    [Fact]
    public void A_completed_workbook_accepts_nothing_further()
    {
        using MemoryStream destination = new();
        using SpreadsheetStreamWriter writer = new(destination, "Data");

        writer.StartRow(1);
        writer.WriteCell(1, "Reference");
        writer.Complete();

        Assert.Throws<InvalidOperationException>(writer.Complete);
        Assert.Throws<InvalidOperationException>(() => writer.StartRow(2));
        Assert.Throws<InvalidOperationException>(() => writer.WriteCell(2, "x"));
        Assert.Throws<InvalidOperationException>(() => writer.SetColumn(1, 10D));

        // And disposing afterwards is not a second close of the archive.
        writer.Dispose();
    }

    [Fact]
    public void An_empty_sheet_is_still_a_workbook()
    {
        byte[] bytes = Streamed(_ => { });

        using SpreadsheetPackage reopened = new(new MemoryStream(bytes));

        Assert.NotNull(reopened.Workbook.Worksheets["Data"]);
        Assert.Null(reopened.Workbook.Worksheets["Data"]!.Cells[1, 1].Value);
    }

    [Fact]
    public void A_destination_that_cannot_be_written_to_is_refused()
    {
        Assert.Throws<ArgumentNullException>(() => new SpreadsheetStreamWriter(null!, "Data"));
        Assert.Throws<ArgumentNullException>(() => new SpreadsheetStreamWriter(new MemoryStream(), null!));
        Assert.Throws<ArgumentException>(
            () => new SpreadsheetStreamWriter(new MemoryStream([1, 2, 3], writable: false), "Data"));
    }

    /// <summary>Writes a workbook and returns its bytes, having checked them against the schema.</summary>
    private static byte[] Streamed(Action<SpreadsheetStreamWriter> write, string sheetName = "Data", string? author = null)
    {
        using MemoryStream destination = new();
        using (SpreadsheetStreamWriter writer = new(destination, sheetName, author))
        {
            write(writer);
            writer.Complete();
        }

        byte[] bytes = destination.ToArray();
        AssertValid(bytes);
        return bytes;
    }

    private static WrittenWorkbook StreamedWorkbook(Action<SpreadsheetStreamWriter> write) =>
        WrittenWorkbook.Opening(Streamed(write));

    /// <remarks>
    /// Nothing checks the schema while this writer writes, so everything it writes is checked
    /// here. An element out of order or a missing attribute count is exactly the kind of fault
    /// that opens on one machine and offers to repair on another.
    /// </remarks>
    private static void AssertValid(byte[] bytes)
    {
        using MemoryStream stream = new(bytes, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);

        string[] errors = [.. new OpenXmlValidator().Validate(document)
            .Select(error => $"{error.Path?.XPath}: {error.Description}")];

        Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors));
    }

    private static Table TableOf(WrittenWorkbook written)
    {
        WorksheetPart part = (WorksheetPart)written.WorkbookPart
            .GetPartById(written.Workbook.Sheets!.Elements<Sheet>().First().Id!.Value!);

        return part.TableDefinitionParts.Single().Table!;
    }

    private static string? FormatCodeOf(WrittenWorkbook written, string reference)
    {
        CellFormat format = written.FormatOf(written.Cell("Data", reference)!);

        return written.Stylesheet.NumberingFormats?.Elements<NumberingFormat>()
            .FirstOrDefault(x => x.NumberFormatId!.Value == format.NumberFormatId!.Value)?.FormatCode?.Value;
    }
}

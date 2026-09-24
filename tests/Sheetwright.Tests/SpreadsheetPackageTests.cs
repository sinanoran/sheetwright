using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// The package itself: what a caller writes comes back out.
/// </summary>
/// <remarks>
/// The round trip is the load-bearing test in this suite. Sheetwright
/// assembles the OPC package by hand — a custom <c>System.IO.Packaging.Package</c>
/// over a <c>MemoryStream</c>, with the content types and the zip written by
/// <c>SpreadsheetPackageArchiveWriter</c> — so "does it produce a readable
/// workbook" is a question about this repository's code and not about the
/// library underneath it.
/// </remarks>
public sealed class SpreadsheetPackageTests
{
    [Fact]
    public void A_new_package_writes_a_workbook_that_opens()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = "hello");

        Assert.Equal("hello", written.Text("Data", "A1"));
    }

    [Fact]
    public void The_bytes_are_a_zip_and_not_a_bare_xml_document()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            package.Workbook.Worksheets.Add("Data"));

        Assert.True(written.Bytes.Length > 4);
        Assert.Equal("PK\u0003\u0004", System.Text.Encoding.ASCII.GetString(written.Bytes, 0, 4));
    }

    [Fact]
    public void Every_kind_of_value_survives_the_round_trip()
    {
        DateTime moment = new(2026, 8, 27, 13, 45, 0, DateTimeKind.Unspecified);

        byte[] bytes;
        using (SpreadsheetPackage package = new())
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "text";
            sheet.Cells[2, 1].Value = 42;
            sheet.Cells[3, 1].Value = 3.5D;
            sheet.Cells[4, 1].Value = 12.75M;
            sheet.Cells[5, 1].Value = true;
            sheet.Cells[6, 1].Value = false;
            sheet.Cells[7, 1].Value = moment;
            sheet.Cells[8, 1].Value = 'x';
            bytes = package.GetAsByteArray();
        }

        using SpreadsheetPackage reopened = new(new MemoryStream(bytes));
        SpreadsheetWorksheet read = reopened.Workbook.Worksheets["Data"]!;

        Assert.Equal("text", read.Cells[1, 1].Value);
        Assert.Equal(42D, read.Cells[2, 1].Value);
        Assert.Equal(3.5D, read.Cells[3, 1].Value);
        Assert.Equal(12.75D, read.Cells[4, 1].Value);
        Assert.Equal(true, read.Cells[5, 1].Value);
        Assert.Equal(false, read.Cells[6, 1].Value);
        Assert.Equal(moment, read.Cells[7, 1].Value);
        Assert.Equal("x", read.Cells[8, 1].Value);
    }

    [Fact]
    public void A_cell_nobody_wrote_to_reads_as_nothing()
    {
        // The negative control for the round trip above. Without it, a reader
        // that returned a default for everything would pass every case.
        byte[] bytes;
        using (SpreadsheetPackage package = new())
        {
            package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = "only me";
            bytes = package.GetAsByteArray();
        }

        using SpreadsheetPackage reopened = new(new MemoryStream(bytes));

        Assert.Null(reopened.Workbook.Worksheets["Data"]!.Cells[9, 9].Value);
    }

    [Fact]
    public void Writing_null_clears_a_cell_rather_than_writing_the_word()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "something";
            sheet.Cells[1, 1].Value = null;
        });

        Cell? cell = written.Cell("Data", "A1");

        Assert.NotNull(cell);
        Assert.Null(cell.CellValue);
        Assert.Null(cell.DataType);
    }

    [Fact]
    public void DBNull_is_emptiness_and_not_a_value_to_convert()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = DBNull.Value);

        Assert.Null(written.Cell("Data", "A1")!.CellValue);
    }

    [Fact]
    public void A_repeated_string_is_stored_once()
    {
        // The shared string table is the reason a 200,000-row export is not
        // 200,000 copies of the same status word.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "Approved";
            sheet.Cells[2, 1].Value = "Approved";
            sheet.Cells[3, 1].Value = "Rejected";
        });

        Assert.Equal(["Approved", "Rejected"], written.SharedStrings);
        Assert.Equal("Approved", written.Text("Data", "A2"));
    }

    [Fact]
    public void A_string_that_would_break_the_XML_is_cleaned_on_the_way_in()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = "be\u0001fore");

        Assert.Equal("before", written.Text("Data", "A1"));
    }

    [Fact]
    public void Leading_and_trailing_space_is_preserved_rather_than_collapsed()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = "  padded  ");

        SharedStringItem item = written.WorkbookPart.SharedStringTablePart!.SharedStringTable!
            .Elements<SharedStringItem>().Single();

        Assert.Equal(
            DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve,
            item.Text!.Space!.Value);
        Assert.Equal("  padded  ", written.Text("Data", "A1"));
    }

    [Fact]
    public void A_formula_is_written_without_its_equals_sign()
    {
        // Excel stores SUM(A1:A2); the leading = is how a person types it.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            package.Workbook.Worksheets.Add("Data").Cells[3, 1].Formula = "=SUM(A1:A2)");

        Assert.Equal("SUM(A1:A2)", written.Cell("Data", "A3")!.CellFormula!.Text);
    }

    [Fact]
    public void A_formula_replaces_the_value_that_was_there()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetRange cell = package.Workbook.Worksheets.Add("Data").Cells[1, 1];
            cell.Value = "stale";
            cell.Formula = "TODAY()";
        });

        Cell cell = written.Cell("Data", "A1")!;

        Assert.Equal("TODAY()", cell.CellFormula!.Text);
        Assert.Null(cell.CellValue);
    }

    [Fact]
    public void A_value_replaces_the_formula_that_was_there()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetRange cell = package.Workbook.Worksheets.Add("Data").Cells[1, 1];
            cell.Formula = "TODAY()";
            cell.Value = 1;
        });

        Assert.Null(written.Cell("Data", "A1")!.CellFormula);
        Assert.Equal("1", written.Text("Data", "A1"));
    }

    [Fact]
    public void An_empty_formula_removes_the_formula()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetRange cell = package.Workbook.Worksheets.Add("Data").Cells[1, 1];
            cell.Formula = "TODAY()";
            cell.Formula = null;
        });

        Assert.Null(written.Cell("Data", "A1")!.CellFormula);
    }

    [Fact]
    public void The_author_is_recorded_and_read_back()
    {
        byte[] bytes;
        using (SpreadsheetPackage package = new())
        {
            package.Workbook.Worksheets.Add("Data");
            package.Workbook.Properties.Author = "Sheetwright";
            bytes = package.GetAsByteArray();
        }

        using SpreadsheetPackage reopened = new(new MemoryStream(bytes));

        Assert.Equal("Sheetwright", reopened.Workbook.Properties.Author);
    }

    [Fact]
    public void Saving_to_a_stream_writes_the_same_bytes_as_asking_for_them()
    {
        using SpreadsheetPackage package = new();
        package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = "hello";

        using MemoryStream destination = new();
        package.SaveAs(destination);

        Assert.Equal(package.GetAsByteArray(), destination.ToArray());
    }

    [Fact]
    public void Asking_for_the_bytes_twice_gives_the_same_answer()
    {
        // Finalising is destructive — it disposes the document and closes the
        // package — so the second call has to be a no-op rather than a repeat.
        using SpreadsheetPackage package = new();
        package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = "hello";

        Assert.Equal(package.GetAsByteArray(), package.GetAsByteArray());
    }

    [Fact]
    public void A_finalised_package_refuses_further_edits()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        package.GetAsByteArray();

        Assert.Throws<InvalidOperationException>(() => sheet.Cells[1, 1].Value = "too late");
    }

    [Fact]
    public void A_package_opened_for_reading_refuses_to_be_written_to()
    {
        byte[] bytes;
        using (SpreadsheetPackage package = new())
        {
            package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = "hello";
            bytes = package.GetAsByteArray();
        }

        using SpreadsheetPackage reopened = new(new MemoryStream(bytes));
        SpreadsheetWorksheet sheet = reopened.Workbook.Worksheets["Data"]!;

        Assert.Throws<InvalidOperationException>(() => sheet.Cells[1, 1].Value = "no");
        Assert.Throws<InvalidOperationException>(() => reopened.GetAsByteArray());
    }

    [Fact]
    public void Opening_reads_from_the_start_of_a_stream_somebody_left_at_the_end()
    {
        byte[] bytes;
        using (SpreadsheetPackage package = new())
        {
            package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = "hello";
            bytes = package.GetAsByteArray();
        }

        MemoryStream source = new(bytes);
        source.Position = source.Length;

        using SpreadsheetPackage reopened = new(source);

        Assert.Equal("hello", reopened.Workbook.Worksheets["Data"]!.Cells[1, 1].Value);
    }

    [Fact]
    public void Opening_something_that_is_not_a_workbook_fails_rather_than_returning_an_empty_one()
    {
        using MemoryStream notAWorkbook = new([1, 2, 3, 4]);

        Assert.ThrowsAny<Exception>(() => new SpreadsheetPackage(notAWorkbook));
    }

    [Fact]
    public void Disposing_without_saving_does_not_throw()
    {
        SpreadsheetPackage package = new();
        package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = "hello";

        package.Dispose();
    }

    [Fact]
    public void Disposing_after_saving_does_not_throw()
    {
        SpreadsheetPackage package = new();
        package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = "hello";
        package.GetAsByteArray();

        package.Dispose();
    }

    [Fact]
    public void A_null_destination_is_refused()
    {
        using SpreadsheetPackage package = new();

        Assert.Throws<ArgumentNullException>(() => package.SaveAs(null!));
    }

    [Fact]
    public void A_null_source_is_refused()
    {
        Assert.Throws<ArgumentNullException>(() => new SpreadsheetPackage((Stream)null!));
    }
}

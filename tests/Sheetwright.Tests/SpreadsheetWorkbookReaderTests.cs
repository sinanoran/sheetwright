using System.Globalization;
using System.IO.Compression;
using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// The streaming workbook reader: a file made by something else, read a row at a time.
/// </summary>
/// <remarks>
/// <para>
/// Two kinds of file. Those Sheetwright's own writer made, which prove the two halves of the library agree;
/// and those assembled by hand, byte for byte, which are what other programs write and the writer never does:
/// inline strings, rich text with a phonetic guide, a date system from 1904, the strict namespaces, cells with no
/// reference, formulas with no stored result.
/// </para>
/// <para>
/// A cell is read as what the file says it is. The cases here are the ones where a reader that guessed would be
/// wrong without any error: a date is a number with a date format, the 1900 leap-year mistake moves every date
/// before March by a day, and a number too big for a decimal must come back as the text the file holds and not as
/// zero.
/// </para>
/// </remarks>
public sealed class SpreadsheetWorkbookReaderTests
{
    private static List<SpreadsheetRow> Read(MemoryStream stream, SpreadsheetReaderOptions? options = null, int sheet = 0)
    {
        using SpreadsheetWorkbookReader reader = SpreadsheetWorkbookReader.Open(stream, options);
        return [.. reader.ReadRows(sheet)];
    }

    private static List<SpreadsheetRow> Read(XlsxBuilder workbook, SpreadsheetReaderOptions? options = null) => Read(workbook.ToStream(), options);

    private static SpreadsheetCellValue Cell(string cell, string? styles = null, int? style = null, bool date1904 = false, string? sharedStrings = null)
    {
        string xml = cell.Replace("{S}", style is int s ? $" s=\"{s}\"" : string.Empty, StringComparison.Ordinal);
        List<SpreadsheetRow> rows = Read(XlsxBuilder.Workbook(XlsxBuilder.Row(1, xml), sharedStrings, styles, date1904));
        return rows.Single()[0];
    }

    [Fact]
    public void A_workbook_the_writer_made_reads_back_cell_for_cell()
    {
        byte[] bytes;
        using (SpreadsheetPackage package = new())
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("People");
            sheet.Cells["A1"].Value = "Name";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["C1"].Value = "When";
            sheet.Cells["D1"].Value = "Active";
            sheet.Cells["A2"].Value = "Özlem Çelik";
            sheet.Cells["B2"].Value = 12_480.5m;
            sheet.Cells["C2"].Value = new DateTime(2026, 11, 20, 13, 30, 0);
            sheet.Cells["D2"].Value = true;
            sheet.Cells["A3"].Value = "Ada";
            sheet.Cells["B3"].Value = -3;
            sheet.Cells["C3"].Value = new DateTime(2026, 1, 5);
            sheet.Cells["D3"].Value = false;
            sheet.Cells["A5"].Value = "after a gap";
            bytes = package.GetAsByteArray();
        }

        List<SpreadsheetRow> rows = Read(new MemoryStream(bytes));

        Assert.Equal([1, 2, 3, 5], rows.Select(row => row.RowNumber));
        Assert.Equal("Name", rows[0][0].Text);
        Assert.Equal("Özlem Çelik", rows[1][0].Text);
        Assert.Equal(SpreadsheetCellKind.Number, rows[1][1].Kind);
        Assert.Equal(12_480.5m, rows[1][1].Number);
        Assert.Equal(SpreadsheetCellKind.Date, rows[1][2].Kind);
        Assert.Equal(new DateTime(2026, 11, 20, 13, 30, 0), rows[1][2].Date);
        Assert.True(rows[1][3].Boolean);
        Assert.Equal(-3m, rows[2][1].Number);
        Assert.Equal(new DateTime(2026, 1, 5), rows[2][2].Date);
        Assert.False(rows[2][3].Boolean);
        Assert.Equal("after a gap", rows[3][0].Text);
    }

    [Fact]
    public void The_sheets_are_listed_in_tab_order_and_found_by_name_without_regard_to_case()
    {
        byte[] bytes;
        using (SpreadsheetPackage package = new())
        {
            package.Workbook.Worksheets.Add("First").Cells["A1"].Value = "one";
            package.Workbook.Worksheets.Add("Second").Cells["A1"].Value = "two";
            bytes = package.GetAsByteArray();
        }

        using SpreadsheetWorkbookReader reader = SpreadsheetWorkbookReader.Open(new MemoryStream(bytes));

        Assert.Equal(["First", "Second"], reader.Sheets.Select(sheet => sheet.Name));
        Assert.Equal([0, 1], reader.Sheets.Select(sheet => sheet.Index));
        Assert.Equal("two", reader.ReadRows("sECOND").Single()[0].Text);
        Assert.Equal("one", reader.ReadRows(0).Single()[0].Text);
        Assert.Throws<KeyNotFoundException>(() => reader.ReadRows("Third"));
        Assert.Throws<ArgumentOutOfRangeException>(() => reader.ReadRows(2));
    }

    [Fact]
    public void A_hidden_sheet_is_listed_as_hidden_and_a_chart_sheet_is_not_listed()
    {
        XlsxBuilder builder = XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", "x")), sheetName: "Shown");
        builder.Replace(
            "xl/workbook.xml",
            $"<workbook xmlns=\"{XlsxBuilder.Main}\" xmlns:r=\"{XlsxBuilder.Rel}\"><sheets>"
            + "<sheet name=\"Shown\" sheetId=\"1\" r:id=\"rId1\"/>"
            + "<sheet name=\"Secret\" sheetId=\"2\" state=\"veryHidden\" r:id=\"rId1\"/>"
            + "<sheet name=\"Chart\" sheetId=\"3\" r:id=\"rId9\"/>"
            + "</sheets></workbook>");
        builder.Replace(
            "xl/_rels/workbook.xml.rels",
            $"<Relationships xmlns=\"{XlsxBuilder.PackageRel}\">"
            + $"<Relationship Id=\"rId1\" Type=\"{XlsxBuilder.Rel}/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
            + $"<Relationship Id=\"rId9\" Type=\"{XlsxBuilder.Rel}/chartsheet\" Target=\"chartsheets/sheet1.xml\"/>"
            + "</Relationships>");

        using SpreadsheetWorkbookReader reader = SpreadsheetWorkbookReader.Open(builder.ToStream());

        Assert.Equal(["Shown", "Secret"], reader.Sheets.Select(sheet => sheet.Name));
        Assert.Equal(SpreadsheetSheetVisibility.Visible, reader.Sheets[0].Visibility);
        Assert.Equal(SpreadsheetSheetVisibility.VeryHidden, reader.Sheets[1].Visibility);
    }

    [Fact]
    public void Inline_strings_keep_their_spaces_and_rich_text_is_joined_without_its_phonetic_guide()
    {
        string shared = "<si><r><t>Hel</t></r><r><rPr><b/></rPr><t xml:space=\"preserve\">lo </t></r><r><t>there</t></r><rPh sb=\"0\" eb=\"1\"><t>ヨミ</t></rPh></si>";
        List<SpreadsheetRow> rows = Read(XlsxBuilder.Workbook(
            XlsxBuilder.Row(
                1,
                "<c r=\"A1\" t=\"inlineStr\"><is><t xml:space=\"preserve\">  padded  </t></is></c>",
                XlsxBuilder.Shared("B1", 0)),
            sharedStrings: shared));

        Assert.Equal("  padded  ", rows[0][0].Text);
        Assert.Equal("Hello there", rows[0][1].Text);
    }

    [Fact]
    public void A_shared_string_that_is_not_in_the_table_is_a_malformed_file()
    {
        XlsxBuilder builder = XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Shared("A1", 7)), sharedStrings: "<si><t>only</t></si>");

        Assert.Throws<InvalidDataException>(() => Read(builder));
    }

    [Fact]
    public void A_formula_is_read_as_the_value_stored_with_it_and_never_evaluated()
    {
        List<SpreadsheetRow> rows = Read(XlsxBuilder.Workbook(
            XlsxBuilder.Row(
                1,
                "<c r=\"A1\"><f>1+1</f><v>2</v></c>",
                "<c r=\"B1\" t=\"str\"><f>\"a\"&amp;\"b\"</f><v>ab</v></c>",
                "<c r=\"C1\"><f>NOW()</f></c>",
                "<c r=\"D1\" t=\"b\"><f>TRUE()</f><v>1</v></c>")));

        Assert.Equal(2m, rows[0][0].Number);
        Assert.Equal("ab", rows[0][1].Text);
        Assert.True(rows[0][2].IsEmpty);
        Assert.True(rows[0][3].Boolean);
        Assert.Equal(4, rows[0].Count);
    }

    [Fact]
    public void Booleans_errors_and_iso_dates_have_their_own_kinds()
    {
        Assert.False(Cell("<c r=\"A1\" t=\"b\"><v>0</v></c>").Boolean);
        Assert.True(Cell("<c r=\"A1\" t=\"b\"><v>1</v></c>").Boolean);

        SpreadsheetCellValue error = Cell("<c r=\"A1\" t=\"e\"><v>#DIV/0!</v></c>");
        Assert.Equal(SpreadsheetCellKind.Error, error.Kind);
        Assert.Equal("#DIV/0!", error.Text);

        SpreadsheetCellValue iso = Cell("<c r=\"A1\" t=\"d\"><v>2026-11-20T13:30:00</v></c>");
        Assert.Equal(new DateTime(2026, 11, 20, 13, 30, 0), iso.Date);
    }

    [Theory]
    [InlineData("12.5", "12.5")]
    [InlineData("-0.5", "-0.5")]
    [InlineData("1E3", "1000")]
    [InlineData("0.1", "0.1")]
    [InlineData("  7  ", "7")]
    public void A_number_is_a_decimal(string raw, string expected)
    {
        SpreadsheetCellValue value = Cell(XlsxBuilder.Number("A1", raw));

        Assert.Equal(SpreadsheetCellKind.Number, value.Kind);
        Assert.Equal(decimal.Parse(expected, CultureInfo.InvariantCulture), value.Number);
    }

    [Theory]
    [InlineData("1E+300")]
    [InlineData("NaN")]
    [InlineData("INF")]
    public void A_number_a_decimal_cannot_hold_comes_back_as_the_text_the_file_has(string raw)
    {
        SpreadsheetCellValue value = Cell(XlsxBuilder.Number("A1", raw));

        Assert.Equal(SpreadsheetCellKind.Text, value.Kind);
        Assert.Equal(raw, value.Text);
    }

    [Theory]
    [InlineData(14)]
    [InlineData(22)]
    [InlineData(30)]
    [InlineData(45)]
    [InlineData(58)]
    public void A_number_in_a_built_in_date_format_is_a_date(int formatId)
    {
        string serial = new DateTime(2026, 11, 20).ToOADate().ToString("R", CultureInfo.InvariantCulture);
        string styles = XlsxBuilder.Styles(string.Empty, 0, formatId);

        SpreadsheetCellValue value = Cell(XlsxBuilder.Number("A1", serial, style: 1), styles);

        Assert.Equal(SpreadsheetCellKind.Date, value.Kind);
        Assert.Equal(new DateTime(2026, 11, 20), value.Date);
    }

    [Theory]
    [InlineData("yyyy-mm-dd hh:mm", true)]
    [InlineData("dd/mm/yy", true)]
    [InlineData("[h]:mm:ss", true)]
    [InlineData("[$-409]mmmm d, yyyy", true)]
    [InlineData("#,##0.00", false)]
    [InlineData("#,##0.00 \"mm\"", false)]
    [InlineData("0.00E+00", false)]
    [InlineData("[Red]0.00", false)]
    [InlineData("General", false)]
    [InlineData("0 \\d", false)]
    public void A_custom_format_is_a_date_format_by_what_it_says_outside_quotes_and_brackets(string code, bool isDate)
    {
        string styles = XlsxBuilder.Styles($"<numFmt numFmtId=\"164\" formatCode=\"{System.Security.SecurityElement.Escape(code)}\"/>", 0, 164);
        string serial = new DateTime(2026, 11, 20, 6, 0, 0).ToOADate().ToString("R", CultureInfo.InvariantCulture);

        SpreadsheetCellValue value = Cell(XlsxBuilder.Number("A1", serial, style: 1), styles);

        Assert.Equal(isDate ? SpreadsheetCellKind.Date : SpreadsheetCellKind.Number, value.Kind);
    }

    [Fact]
    public void A_number_in_the_general_format_is_not_a_date_whatever_its_size()
    {
        SpreadsheetCellValue value = Cell(XlsxBuilder.Number("A1", "46346", style: 0), XlsxBuilder.Styles(string.Empty, 0));

        Assert.Equal(SpreadsheetCellKind.Number, value.Kind);
    }

    [Fact]
    public void A_time_with_no_date_is_on_30_december_1899()
    {
        SpreadsheetCellValue value = Cell(XlsxBuilder.Number("A1", "0.75", style: 1), XlsxBuilder.Styles(string.Empty, 0, 21));

        Assert.Equal(new DateTime(1899, 12, 30, 18, 0, 0), value.Date);
    }

    [Theory]
    [InlineData("1", 1900, 1, 1)]
    [InlineData("59", 1900, 2, 28)]
    [InlineData("61", 1900, 3, 1)]
    [InlineData("36526", 2000, 1, 1)]
    public void The_1900_date_system_counts_a_29th_of_february_that_never_was(string serial, int year, int month, int day)
    {
        SpreadsheetCellValue value = Cell(XlsxBuilder.Number("A1", serial, style: 1), XlsxBuilder.Styles(string.Empty, 0, 14));

        Assert.Equal(new DateTime(year, month, day), value.Date);
    }

    [Fact]
    public void The_day_that_never_was_has_no_date_and_stays_a_number()
    {
        SpreadsheetCellValue value = Cell(XlsxBuilder.Number("A1", "60", style: 1), XlsxBuilder.Styles(string.Empty, 0, 14));

        Assert.Equal(SpreadsheetCellKind.Number, value.Kind);
        Assert.Equal(60m, value.Number);
    }

    [Theory]
    [InlineData("0", 1904, 1, 1)]
    [InlineData("1", 1904, 1, 2)]
    [InlineData("42370", 2020, 1, 2)]
    public void A_workbook_in_the_1904_date_system_counts_from_1904(string serial, int year, int month, int day)
    {
        SpreadsheetCellValue value = Cell(XlsxBuilder.Number("A1", serial, style: 1), XlsxBuilder.Styles(string.Empty, 0, 14), date1904: true);

        Assert.Equal(new DateTime(year, month, day), value.Date);
    }

    [Fact]
    public void A_serial_past_the_last_date_stays_a_number()
    {
        SpreadsheetCellValue value = Cell(XlsxBuilder.Number("A1", "3000000", style: 1), XlsxBuilder.Styles(string.Empty, 0, 14));

        Assert.Equal(SpreadsheetCellKind.Number, value.Kind);
    }

    [Fact]
    public void Cells_with_no_reference_follow_one_another_and_a_gap_is_empty_cells()
    {
        List<SpreadsheetRow> rows = Read(XlsxBuilder.Workbook(
            "<row><c><v>1</v></c><c><v>2</v></c></row><row><c r=\"C1\"><v>3</v></c></row>"));

        Assert.Equal([1, 2], rows.Select(row => row.RowNumber));
        Assert.Equal(2, rows[0].Count);
        Assert.Equal(3, rows[1].Count);
        Assert.True(rows[1][0].IsEmpty);
        Assert.True(rows[1][1].IsEmpty);
        Assert.Equal(3m, rows[1][2].Number);
    }

    [Fact]
    public void A_row_ends_at_its_last_cell_with_a_value_and_not_at_a_styled_empty_one()
    {
        List<SpreadsheetRow> rows = Read(XlsxBuilder.Workbook(
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\" s=\"4\"/><c r=\"C1\" s=\"4\"><v></v></c></row>"));

        Assert.Equal(1, rows[0].Count);
    }

    [Fact]
    public void Empty_rows_are_skipped_and_return_when_asked_for()
    {
        XlsxBuilder Make() => XlsxBuilder.Workbook(
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row><row r=\"2\"><c r=\"A2\" s=\"1\"/></row><row r=\"3\"/><row r=\"4\"><c r=\"A4\"><v>4</v></c></row>");

        Assert.Equal([1, 4], Read(Make()).Select(row => row.RowNumber));

        List<SpreadsheetRow> all = Read(Make(), new SpreadsheetReaderOptions { IncludeEmptyRows = true });
        Assert.Equal([1, 2, 3, 4], all.Select(row => row.RowNumber));
        Assert.True(all[1].IsEmpty);
    }

    [Fact]
    public void A_workbook_in_the_strict_namespaces_reads_the_same()
    {
        const string main = "http://purl.oclc.org/ooxml/spreadsheetml/main";
        const string rel = "http://purl.oclc.org/ooxml/officeDocument/relationships";

        XlsxBuilder builder = new();
        builder.Add("_rels/.rels", $"<Relationships xmlns=\"{XlsxBuilder.PackageRel}\"><Relationship Id=\"rId1\" Type=\"{rel}/officeDocument\" Target=\"/xl/workbook.xml\"/></Relationships>");
        builder.Add("xl/workbook.xml", $"<workbook xmlns=\"{main}\" xmlns:r=\"{rel}\"><sheets><sheet name=\"S\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
        builder.Add("xl/_rels/workbook.xml.rels", $"<Relationships xmlns=\"{XlsxBuilder.PackageRel}\"><Relationship Id=\"rId1\" Type=\"{rel}/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
        builder.Add("xl/worksheets/sheet1.xml", $"<worksheet xmlns=\"{main}\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>strict</t></is></c></row></sheetData></worksheet>");

        Assert.Equal("strict", Read(builder)[0][0].Text);
    }

    [Fact]
    public void A_relationship_may_name_a_part_by_an_absolute_path_or_climb_with_dots()
    {
        XlsxBuilder builder = XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", "found")));
        builder.Replace(
            "xl/_rels/workbook.xml.rels",
            $"<Relationships xmlns=\"{XlsxBuilder.PackageRel}\"><Relationship Id=\"rId1\" Type=\"{XlsxBuilder.Rel}/worksheet\" Target=\"../xl/worksheets/./sheet1.xml\"/></Relationships>");

        Assert.Equal("found", Read(builder)[0][0].Text);
    }

    [Fact]
    public void A_relationship_that_climbs_out_of_the_package_is_refused()
    {
        XlsxBuilder builder = XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", "x")));
        builder.Replace(
            "xl/_rels/workbook.xml.rels",
            $"<Relationships xmlns=\"{XlsxBuilder.PackageRel}\"><Relationship Id=\"rId1\" Type=\"{XlsxBuilder.Rel}/worksheet\" Target=\"../../../etc/passwd\"/></Relationships>");

        Assert.Throws<InvalidDataException>(() => Read(builder));
    }

    [Fact]
    public void An_external_relationship_is_never_followed()
    {
        XlsxBuilder builder = XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", "x")));
        builder.Replace(
            "xl/_rels/workbook.xml.rels",
            $"<Relationships xmlns=\"{XlsxBuilder.PackageRel}\">"
            + $"<Relationship Id=\"rId1\" Type=\"{XlsxBuilder.Rel}/worksheet\" Target=\"http://example.invalid/sheet.xml\" TargetMode=\"External\"/></Relationships>");

        using SpreadsheetWorkbookReader reader = SpreadsheetWorkbookReader.Open(builder.ToStream());

        Assert.Empty(reader.Sheets);
    }

    [Fact]
    public void Something_that_is_not_a_zip_is_not_a_workbook()
    {
        using MemoryStream stream = new("name,city\nAda,London\n"u8.ToArray());

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => SpreadsheetWorkbookReader.Open(stream));

        Assert.Contains("zip", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void An_empty_stream_is_not_a_workbook()
    {
        Assert.Throws<InvalidDataException>(() => SpreadsheetWorkbookReader.Open(new MemoryStream()));
    }

    [Fact]
    public void A_zip_that_is_not_a_workbook_is_refused_with_what_is_missing()
    {
        XlsxBuilder builder = new();
        builder.Add("readme.txt", "hello");

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => SpreadsheetWorkbookReader.Open(builder.ToStream()));

        Assert.Contains("not an .xlsx", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_workbook_that_names_a_sheet_part_the_package_does_not_have_is_refused()
    {
        XlsxBuilder builder = XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", "x")));
        builder.Remove("xl/worksheets/sheet1.xml");

        Assert.Throws<InvalidDataException>(() => SpreadsheetWorkbookReader.Open(builder.ToStream()));
    }

    [Fact]
    public void Two_parts_with_one_name_are_refused_because_either_could_be_the_one_that_is_read()
    {
        XlsxBuilder builder = XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", "x")));
        builder.Add("xl/worksheets/sheet1.xml", XlsxBuilder.Sheet(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", "other"))));

        Assert.Throws<InvalidDataException>(() => SpreadsheetWorkbookReader.Open(builder.ToStream()));
    }

    [Fact]
    public void Rows_that_are_out_of_order_or_cells_that_are_out_of_order_are_refused()
    {
        Assert.Throws<InvalidDataException>(() => Read(XlsxBuilder.Workbook("<row r=\"2\"><c r=\"A2\"><v>1</v></c></row><row r=\"1\"><c r=\"A1\"><v>1</v></c></row>")));
        Assert.Throws<InvalidDataException>(() => Read(XlsxBuilder.Workbook("<row r=\"1\"><c r=\"B1\"><v>1</v></c><c r=\"A1\"><v>1</v></c></row>")));
        Assert.Throws<InvalidDataException>(() => Read(XlsxBuilder.Workbook("<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"A1\"><v>2</v></c></row>")));
    }

    [Fact]
    public void A_row_numbered_past_the_last_row_a_spreadsheet_has_is_refused()
    {
        Assert.Throws<InvalidDataException>(() => Read(XlsxBuilder.Workbook("<row r=\"1048577\"><c r=\"A1048577\"><v>1</v></c></row>")));
        Assert.Throws<InvalidDataException>(() => Read(XlsxBuilder.Workbook("<row r=\"x\"><c r=\"A1\"><v>1</v></c></row>")));
    }

    [Fact]
    public void A_cell_with_no_column_in_its_reference_is_refused()
    {
        Assert.Throws<InvalidDataException>(() => Read(XlsxBuilder.Workbook("<row r=\"1\"><c r=\"7\"><v>1</v></c></row>")));
    }

    [Fact]
    public void A_part_that_is_not_well_formed_is_refused_as_such()
    {
        XlsxBuilder builder = XlsxBuilder.Workbook(string.Empty);
        builder.Replace("xl/worksheets/sheet1.xml", "<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><v>1</c></row></sheetData></worksheet>");

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => Read(builder));

        Assert.Contains("well-formed", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_part_with_a_document_type_declaration_is_refused_and_nothing_is_expanded()
    {
        const string bomb = "<?xml version=\"1.0\"?><!DOCTYPE lolz [<!ENTITY lol \"lol\"><!ENTITY lol2 \"&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;\">"
            + "<!ENTITY lol3 \"&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;\">]>"
            + "<worksheet><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>&lol3;</t></is></c></row></sheetData></worksheet>";
        XlsxBuilder builder = XlsxBuilder.Workbook(string.Empty);
        builder.Replace("xl/worksheets/sheet1.xml", bomb);

        Assert.Throws<InvalidDataException>(() => Read(builder));
    }

    [Fact]
    public void An_external_entity_is_never_read()
    {
        string file = Path.Combine(Path.GetTempPath(), "sheetwright-xxe-" + Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(file, "SECRET");
        try
        {
            string xxe = $"<?xml version=\"1.0\"?><!DOCTYPE x [<!ENTITY secret SYSTEM \"{new Uri(file).AbsoluteUri}\">]>"
                + "<worksheet><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>&secret;</t></is></c></row></sheetData></worksheet>";
            XlsxBuilder builder = XlsxBuilder.Workbook(string.Empty);
            builder.Replace("xl/worksheets/sheet1.xml", xxe);

            Assert.Throws<InvalidDataException>(() => Read(builder));
        }
        finally
        {
            File.Delete(file);
        }
    }

    [Fact]
    public void Nesting_deeper_than_a_workbook_ever_is_refused()
    {
        string deep = string.Concat(Enumerable.Repeat("<a>", 200)) + string.Concat(Enumerable.Repeat("</a>", 200));
        XlsxBuilder builder = XlsxBuilder.Workbook(string.Empty);
        builder.Replace("xl/worksheets/sheet1.xml", $"<worksheet><sheetData><row r=\"1\"><c r=\"A1\">{deep}</c></row></sheetData></worksheet>");

        Assert.Throws<InvalidDataException>(() => Read(builder));
    }

    [Fact]
    public void An_empty_sheet_has_no_rows()
    {
        Assert.Empty(Read(XlsxBuilder.Workbook(string.Empty)));

        XlsxBuilder empty = XlsxBuilder.Workbook(string.Empty);
        empty.Replace("xl/worksheets/sheet1.xml", $"<worksheet xmlns=\"{XlsxBuilder.Main}\"><sheetData/></worksheet>");
        Assert.Empty(Read(empty));
    }

    [Fact]
    public void A_workbook_stream_must_be_seekable()
    {
        using NonSeekableStream stream = new(new byte[] { 1, 2, 3 });

        Assert.Throws<ArgumentException>(() => SpreadsheetWorkbookReader.Open(stream));
    }

    [Fact]
    public void Disposing_closes_the_stream_unless_it_was_left_open_and_reading_after_it_is_refused()
    {
        MemoryStream closed = XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", "x"))).ToStream();
        SpreadsheetWorkbookReader first = SpreadsheetWorkbookReader.Open(closed);
        first.Dispose();
        Assert.False(closed.CanRead);
        Assert.Throws<ObjectDisposedException>(() => first.ReadRows(0));

        MemoryStream open = XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", "x"))).ToStream();
        SpreadsheetWorkbookReader.Open(open, leaveOpen: true).Dispose();
        Assert.True(open.CanRead);
    }

    [Fact]
    public void A_sheet_can_be_read_twice_from_one_reader()
    {
        using SpreadsheetWorkbookReader reader = SpreadsheetWorkbookReader.Open(
            XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", "x"))).ToStream());

        Assert.Equal("x", reader.ReadRows(0).Single()[0].Text);
        Assert.Equal("x", reader.ReadRows(0).Single()[0].Text);
    }

    [Fact]
    public void A_row_may_be_kept_after_the_next_is_read()
    {
        List<SpreadsheetRow> rows = Read(XlsxBuilder.Workbook(
            XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", "first")) + XlsxBuilder.Row(2, XlsxBuilder.Inline("A2", "second"))));

        Assert.Equal("first", rows[0][0].Text);
        Assert.Equal("second", rows[1][0].Text);
    }

    [Fact]
    public void A_workbook_read_from_a_file_stream_works_the_same()
    {
        string path = Path.Combine(Path.GetTempPath(), "sheetwright-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            using (FileStream write = File.Create(path))
            {
                XlsxBuilder.Workbook(XlsxBuilder.Row(1, XlsxBuilder.Inline("A1", "from disk"))).ToStream().CopyTo(write);
            }

            using FileStream read = File.OpenRead(path);
            using SpreadsheetWorkbookReader reader = SpreadsheetWorkbookReader.Open(read);

            Assert.Equal("from disk", reader.ReadRows(0).Single()[0].Text);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void The_cell_kinds_print_as_text_a_person_could_read()
    {
        Assert.Equal("12.5", SpreadsheetCellValue.FromNumber(12.5m).ToString());
        Assert.Equal("true", SpreadsheetCellValue.FromBoolean(true).ToString());
        Assert.Equal("2026-11-20T00:00:00.0000000", SpreadsheetCellValue.FromDate(new DateTime(2026, 11, 20)).ToString());
        Assert.Equal("#N/A", SpreadsheetCellValue.FromError("#N/A").ToString());
        Assert.Equal(string.Empty, SpreadsheetCellValue.Empty.ToString());
        Assert.Equal(SpreadsheetCellValue.FromText("x"), SpreadsheetCellValue.FromText("x"));
        Assert.NotEqual(SpreadsheetCellValue.FromText("1"), SpreadsheetCellValue.FromNumber(1m));
    }

    private sealed class NonSeekableStream(byte[] bytes) : Stream
    {
        private readonly MemoryStream _inner = new(bytes);

        public override bool CanRead => true;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

        public override void Flush()
        {
        }

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// Tables and defined names — the two things in a workbook that have to agree
/// with something else in it.
/// </summary>
/// <remarks>
/// A table part is its own part with its own relationship, and a table whose
/// column names do not match the cells of its header row is the classic cause of
/// Excel's repair dialogue. A defined name is a formula referring to a sheet by
/// name, so the quoting is load-bearing in exactly the way a sheet called
/// "Q1 Sales" makes obvious.
/// </remarks>
public sealed class SpreadsheetTableAndNameTests
{
    [Fact]
    public void A_table_takes_its_column_names_from_the_header_row()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "Reference";
            sheet.Cells[1, 2].Value = "Amount";
            sheet.Cells[2, 1].Value = "INV-1";
            sheet.Cells[2, 2].Value = 10D;
            sheet.Tables.Add(sheet.Cells["A1:B2"], "Invoices");
        });

        Table table = TableOf(written);

        Assert.Equal("Invoices", table.Name!.Value);
        Assert.Equal("Invoices", table.DisplayName!.Value);
        Assert.Equal("A1:B2", table.Reference!.Value);
        Assert.Equal<string[]>(
            ["Reference", "Amount"],
            [.. table.GetFirstChild<TableColumns>()!.Elements<TableColumn>().Select(x => x.Name!.Value!)]);
    }

    [Fact]
    public void A_table_column_is_numbered_from_one_whatever_column_it_starts_at()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 3].Value = "Reference";
            sheet.Cells[1, 4].Value = "Amount";
            sheet.Cells[2, 3].Value = "INV-1";
            sheet.Tables.Add(sheet.Cells["C1:D2"], "Invoices");
        });

        TableColumns columns = TableOf(written).GetFirstChild<TableColumns>()!;

        Assert.Equal(2U, columns.Count!.Value);
        Assert.Equal<uint[]>([1, 2], [.. columns.Elements<TableColumn>().Select(x => x.Id!.Value)]);
    }

    [Fact]
    public void A_repeated_heading_is_made_unique_because_Excel_refuses_duplicates()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "Amount";
            sheet.Cells[1, 2].Value = "Amount";
            sheet.Cells[2, 1].Value = 1D;
            sheet.Tables.Add(sheet.Cells["A1:B2"], "Invoices");
        });

        string[] names = [.. TableOf(written).GetFirstChild<TableColumns>()!
            .Elements<TableColumn>().Select(x => x.Name!.Value!)];

        Assert.Equal<string[]>(["Amount", "Amount2"], names);

        // And the header cell says so too. Excel checks the table's column names
        // against the cells of its header row and offers to repair the file when
        // they disagree, so renaming one without writing it back produced a table
        // that could not be opened.
        Assert.Equal("Amount2", written.Text("Data", "B1"));
    }

    [Fact]
    public void An_empty_heading_becomes_the_column_letter()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "Reference";
            sheet.Cells[2, 1].Value = "INV-1";
            sheet.Cells[2, 2].Value = 10D;
            sheet.Tables.Add(sheet.Cells["A1:B2"], "Invoices");
        });

        string[] names = [.. TableOf(written).GetFirstChild<TableColumns>()!
            .Elements<TableColumn>().Select(x => x.Name!.Value!)];

        Assert.Equal<string[]>(["Reference", "B"], names);

        // The invented heading is written into the empty cell for the same
        // reason a de-duplicated one is.
        Assert.Equal("B", written.Text("Data", "B1"));
    }

    [Fact]
    public void A_heading_that_needed_no_help_is_left_exactly_as_the_caller_wrote_it()
    {
        // The negative control for the two above: writing every heading back
        // unconditionally would churn the shared string table and, worse, retype
        // a header cell somebody had deliberately styled.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "Reference";
            sheet.Cells[1, 2].Value = "Amount";
            sheet.Cells[2, 1].Value = "INV-1";
            sheet.Tables.Add(sheet.Cells["A1:B2"], "Invoices");
        });

        Assert.Equal(["Reference", "Amount", "INV-1"], written.SharedStrings);
    }

    [Fact]
    public void A_table_over_a_header_row_that_does_not_exist_yet_still_gets_its_cells()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[2, 1].Value = "INV-1";
            sheet.Tables.Add(sheet.Cells["A1:A2"], "Invoices");
        });

        Assert.Equal("A", written.Text("Data", "A1"));
        Assert.Equal("A1:A2", written.Sheet("Data").GetFirstChild<SheetDimension>()!.Reference!.Value);
    }

    [Fact]
    public void A_table_is_its_own_part_and_the_sheet_points_at_it()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "Reference";
            sheet.Cells[2, 1].Value = "INV-1";
            sheet.Tables.Add(sheet.Cells["A1:A2"], "Invoices");
        });

        WorksheetPart worksheetPart = written.WorkbookPart.WorksheetParts.Single();
        TableParts parts = written.Sheet("Data").GetFirstChild<TableParts>()!;

        Assert.Equal(1U, parts.Count!.Value);
        Assert.Equal(
            worksheetPart.GetIdOfPart(worksheetPart.TableDefinitionParts.Single()),
            parts.Elements<TablePart>().Single().Id!.Value);
    }

    [Fact]
    public void Two_tables_get_two_identifiers()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "Reference";
            sheet.Cells[2, 1].Value = "INV-1";
            sheet.Cells[4, 1].Value = "Name";
            sheet.Cells[5, 1].Value = "Ada";
            sheet.Tables.Add(sheet.Cells["A1:A2"], "Invoices");
            sheet.Tables.Add(sheet.Cells["A4:A5"], "People");
        });

        WorksheetPart worksheetPart = written.WorkbookPart.WorksheetParts.Single();
        uint[] ids = [.. worksheetPart.TableDefinitionParts.Select(part => part.Table!.Id!.Value)];

        Assert.Equal(2, ids.Length);
        Assert.Equal(ids.Length, ids.Distinct().Count());
    }

    [Fact]
    public void A_table_style_can_be_changed_after_the_table_is_added()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "Reference";
            sheet.Cells[2, 1].Value = "INV-1";
            SpreadsheetTable table = sheet.Tables.Add(sheet.Cells["A1:A2"], "Invoices");
            table.TableStyle = "TableStyleLight9";
        });

        Assert.Equal("TableStyleLight9", TableOf(written).GetFirstChild<TableStyleInfo>()!.Name!.Value);
    }

    [Fact]
    public void The_default_table_style_is_the_one_Excel_offers_first()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        sheet.Cells[1, 1].Value = "Reference";

        Assert.Equal("TableStyleMedium2", sheet.Tables.Add(sheet.Cells["A1:A2"], "Invoices").TableStyle);
    }

    [Fact]
    public void Tables_are_counted_as_they_are_added()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        sheet.Cells[1, 1].Value = "Reference";

        Assert.Equal(0, sheet.Tables.Count);

        sheet.Tables.Add(sheet.Cells["A1:A2"], "Invoices");

        Assert.Equal(1, sheet.Tables.Count);
    }

    [Fact]
    public void A_null_table_name_or_range_is_refused()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        Assert.Throws<ArgumentNullException>(() => sheet.Tables.Add(sheet.Cells["A1:A2"], null!));
        Assert.Throws<ArgumentNullException>(() => sheet.Tables.Add(null!, "Invoices"));
    }

    [Fact]
    public void A_defined_name_points_at_an_absolute_range_on_its_sheet()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            package.Workbook.Names.Add("Statuses", sheet.Cells["B2:B9"]);
        });

        DefinedName name = written.Workbook.DefinedNames!.Elements<DefinedName>().Single();

        Assert.Equal("Statuses", name.Name!.Value);
        Assert.Equal("'Data'!$B$2:$B$9", name.Text);
    }

    [Fact]
    public void A_defined_name_for_one_cell_does_not_repeat_it()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            package.Workbook.Names.Add("Start", sheet.Cells["B2"]);
        });

        Assert.Equal("'Data'!$B$2", written.Workbook.DefinedNames!.Elements<DefinedName>().Single().Text);
    }

    [Fact]
    public void A_defined_formula_loses_its_leading_equals_sign()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = "x";
            package.Workbook.Names.AddFormula("Rate", "=0.2");
        });

        Assert.Equal("0.2", written.Workbook.DefinedNames!.Elements<DefinedName>().Single().Text);
    }

    [Fact]
    public void A_defined_name_is_found_whatever_the_case_and_absent_when_it_is_absent()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        package.Workbook.Names.Add("Statuses", sheet.Cells["B2:B9"]);

        Assert.True(package.Workbook.Names.ContainsKey("statuses"));
        Assert.False(package.Workbook.Names.ContainsKey("Nothing"));
        Assert.Equal("'Data'!$B$2:$B$9", package.Workbook.Names["STATUSES"].Formula);
    }

    [Fact]
    public void Re_pointing_a_name_drops_the_equals_sign_too()
    {
        // The gap this closes. Adding a name stripped the sign and re-pointing
        // one did not, so the same value written through the two doors produced
        // different files — and the door Excel would have complained about was
        // the one with no test on it.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            package.Workbook.Names.Add("Rate", sheet.Cells["A1"]);
            package.Workbook.Names["Rate"].Formula = "=0.2";
        });

        Assert.Equal("0.2", written.Workbook.DefinedNames!.Elements<DefinedName>().Single().Text);
    }

    [Fact]
    public void Both_ways_of_setting_a_formula_agree()
    {
        // The negative control for the two above: they are one code path now, and
        // this is what notices if somebody splits them again.
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        package.Workbook.Names.AddFormula("Added", "=SUM(A1:A2)");
        package.Workbook.Names.Add("Assigned", sheet.Cells["A1"]);
        package.Workbook.Names["Assigned"].Formula = "=SUM(A1:A2)";

        Assert.Equal(
            package.Workbook.Names["Added"].Formula,
            package.Workbook.Names["Assigned"].Formula);
    }

    [Fact]
    public void A_defined_name_can_be_repointed()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        package.Workbook.Names.Add("Statuses", sheet.Cells["B2:B9"]);
        package.Workbook.Names["Statuses"].Formula = "'Data'!$C$1:$C$5";

        Assert.Equal("'Data'!$C$1:$C$5", package.Workbook.Names["Statuses"].Formula);
    }

    [Fact]
    public void The_defined_names_can_be_enumerated()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        package.Workbook.Names.Add("First", sheet.Cells["A1"]);
        package.Workbook.Names.Add("Second", sheet.Cells["A2"]);

        Assert.Equal(["First", "Second"], package.Workbook.Names.Select(name => name.Name));
    }

    [Fact]
    public void A_null_name_or_range_is_refused()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        Assert.Throws<ArgumentNullException>(() => package.Workbook.Names.Add(null!, sheet.Cells["A1"]));
        Assert.Throws<ArgumentNullException>(() => package.Workbook.Names.Add("Name", null!));
        Assert.Throws<ArgumentNullException>(() => package.Workbook.Names.AddFormula("Name", null!));
        Assert.Throws<ArgumentNullException>(() => package.Workbook.Names.ContainsKey(null!));
    }

    [Fact]
    public void Asking_for_a_name_that_is_not_there_says_so()
    {
        using SpreadsheetPackage package = new();
        package.Workbook.Worksheets.Add("Data");

        KeyNotFoundException error = Assert.Throws<KeyNotFoundException>(
            () => package.Workbook.Names["Nothing"]);

        Assert.Contains("Nothing", error.Message, StringComparison.Ordinal);
    }

    private static Table TableOf(WrittenWorkbook written) =>
        written.WorkbookPart.WorksheetParts.Single().TableDefinitionParts.Single().Table!;
}

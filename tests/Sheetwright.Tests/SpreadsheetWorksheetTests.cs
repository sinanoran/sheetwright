using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// The sheet: its cells, its shape, and the order Excel insists they appear in.
/// </summary>
/// <remarks>
/// <para>
/// Two failures are being defended against here and neither is visible in the
/// object graph. The first is cell order — <c>sheetData</c> is a sequence of rows
/// in ascending index, each a sequence of cells in ascending column, and a caller
/// who fills C1 before A1 must still get a file in that order. The second is the
/// order of the worksheet's own children: <c>CT_Worksheet</c> is a sequence, so
/// putting <c>mergeCells</c> before <c>sheetData</c> produces a workbook Excel
/// offers to repair rather than open.
/// </para>
/// <para>
/// Both are the kind of thing that works for the sheet in front of you and fails
/// on the one a customer builds.
/// </para>
/// </remarks>
public sealed class SpreadsheetWorksheetTests
{
    [Fact]
    public void A_sheet_is_found_by_name_whatever_the_case()
    {
        using SpreadsheetPackage package = new();
        package.Workbook.Worksheets.Add("Summary");

        Assert.NotNull(package.Workbook.Worksheets["summary"]);
        Assert.NotNull(package.Workbook.Worksheets["SUMMARY"]);
    }

    [Fact]
    public void A_sheet_that_is_not_there_is_null_rather_than_an_exception()
    {
        using SpreadsheetPackage package = new();
        package.Workbook.Worksheets.Add("Summary");

        Assert.Null(package.Workbook.Worksheets["Detail"]);
    }

    [Fact]
    public void Sheets_are_numbered_from_one_because_that_is_what_Excel_shows()
    {
        using SpreadsheetPackage package = new();
        package.Workbook.Worksheets.Add("First");
        package.Workbook.Worksheets.Add("Second");

        Assert.Equal(2, package.Workbook.Worksheets.Count);
        Assert.Equal("First", package.Workbook.Worksheets[1].Name);
        Assert.Equal("Second", package.Workbook.Worksheets[2].Name);
    }

    [Fact]
    public void Each_sheet_gets_its_own_identifier()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            package.Workbook.Worksheets.Add("First");
            package.Workbook.Worksheets.Add("Second");
        });

        uint[] ids = [.. written.Workbook.Sheets!.Elements<Sheet>().Select(sheet => sheet.SheetId!.Value)];

        Assert.Equal(ids.Length, ids.Distinct().Count());
    }

    [Fact]
    public void A_null_sheet_name_is_refused()
    {
        using SpreadsheetPackage package = new();

        Assert.Throws<ArgumentNullException>(() => package.Workbook.Worksheets.Add(null!));
    }

    [Fact]
    public void Rows_are_written_in_ascending_order_however_they_were_filled()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[5, 1].Value = "five";
            sheet.Cells[1, 1].Value = "one";
            sheet.Cells[3, 1].Value = "three";
            sheet.Cells[2, 1].Value = "two";
        });

        uint[] rows = [.. written.Sheet("Data").GetFirstChild<SheetData>()!
            .Elements<Row>().Select(row => row.RowIndex!.Value)];

        Assert.Equal<uint[]>([1, 2, 3, 5], rows);
    }

    [Fact]
    public void Cells_are_written_in_ascending_column_order_however_they_were_filled()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 3].Value = "c";
            sheet.Cells[1, 1].Value = "a";
            sheet.Cells[1, 4].Value = "d";
            sheet.Cells[1, 2].Value = "b";
        });

        string[] references = [.. written.Sheet("Data").GetFirstChild<SheetData>()!
            .Elements<Row>().Single()
            .Elements<Cell>().Select(cell => cell.CellReference!.Value!)];

        Assert.Equal<string[]>(["A1", "B1", "C1", "D1"], references);
    }

    [Fact]
    public void Writing_the_same_cell_twice_keeps_one_cell()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "first";
            sheet.Cells[1, 1].Value = "second";
        });

        Assert.Single(written.Sheet("Data").GetFirstChild<SheetData>()!.Elements<Row>().Single().Elements<Cell>());
        Assert.Equal("second", written.Text("Data", "A1"));
    }

    [Fact]
    public void A_sheet_with_nothing_in_it_has_no_dimension()
    {
        using SpreadsheetPackage package = new();

        Assert.Null(package.Workbook.Worksheets.Add("Empty").Dimension);
    }

    [Fact]
    public void The_dimension_is_the_rectangle_that_was_actually_used()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        sheet.Cells[2, 3].Value = "here";
        sheet.Cells[7, 5].Value = "and here";

        SpreadsheetDimension dimension = sheet.Dimension!;

        Assert.Equal((2, 3), (dimension.Start.Row, dimension.Start.Column));
        Assert.Equal((7, 5), (dimension.End.Row, dimension.End.Column));
        Assert.Equal(6, dimension.Rows);
        Assert.Equal(3, dimension.Columns);
        Assert.Equal("C2:E7", dimension.Reference);
    }

    [Fact]
    public void The_dimension_is_written_into_the_file()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "a";
            sheet.Cells[4, 2].Value = "b";
        });

        Assert.Equal("A1:B4", written.Sheet("Data").GetFirstChild<SheetDimension>()!.Reference!.Value);
    }

    [Fact]
    public void A_merge_is_recorded_once_however_many_times_it_is_asked_for()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "title";
            sheet.Cells["A1:C1"].Merge = true;
            sheet.Cells["A1:C1"].Merge = true;
        });

        MergeCells merges = written.Sheet("Data").GetFirstChild<MergeCells>()!;

        Assert.Equal("A1:C1", merges.Elements<MergeCell>().Single().Reference!.Value);
        Assert.Equal(1U, merges.Count!.Value);
    }

    [Fact]
    public void A_merge_can_be_taken_back()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        sheet.Cells["A1:C1"].Merge = true;

        Assert.True(sheet.Cells["A1:C1"].Merge);

        sheet.Cells["A1:C1"].Merge = false;

        Assert.False(sheet.Cells["A1:C1"].Merge);
    }

    [Fact]
    public void Unmerging_everything_leaves_no_merge_element_behind()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            sheet.Cells["A1:C1"].Merge = true;
            sheet.Cells["A1:C1"].Merge = false;
        });

        Assert.Null(written.Sheet("Data").GetFirstChild<MergeCells>());
    }

    [Fact]
    public void Freezing_a_pane_splits_above_and_left_of_the_named_cell()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "header";
            sheet.View.FreezePanes(2, 3);
        });

        Pane pane = written.Sheet("Data").GetFirstChild<SheetViews>()!
            .Elements<SheetView>().Single().Pane!;

        Assert.Equal(1D, pane.VerticalSplit!.Value);
        Assert.Equal(2D, pane.HorizontalSplit!.Value);
        Assert.Equal("C2", pane.TopLeftCell!.Value);
        Assert.Equal(PaneStateValues.Frozen, pane.State!.Value);
        Assert.Equal(PaneValues.BottomRight, pane.ActivePane!.Value);
    }

    [Fact]
    public void Freezing_only_the_top_row_splits_horizontally_and_not_vertically()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "header";
            sheet.View.FreezePanes(2, 1);
        });

        Pane pane = written.Sheet("Data").GetFirstChild<SheetViews>()!
            .Elements<SheetView>().Single().Pane!;

        Assert.Equal(1D, pane.VerticalSplit!.Value);
        Assert.Equal(0D, pane.HorizontalSplit!.Value);
        Assert.Equal(PaneValues.BottomLeft, pane.ActivePane!.Value);
    }

    [Fact]
    public void A_frozen_pane_lands_after_the_sheet_properties_and_not_before_them()
    {
        // SetVbaProject writes a sheetPr for every sheet. sheetViews follows
        // sheetPr in the schema, so a workbook that attached its macros before
        // freezing a pane used to produce a file Excel would not open.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "header";
            sheet.SetCodeName("Data");
            sheet.View.FreezePanes(2, 1);
        });

        IReadOnlyList<string> order = written.ChildOrder("Data");

        Assert.True(
            order.ToList().IndexOf("sheetPr") < order.ToList().IndexOf("sheetViews"),
            "sheetPr must precede sheetViews: " + string.Join(", ", order));
    }

    [Fact]
    public void Everything_a_sheet_can_hold_comes_out_in_schema_order()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.SetCodeName("Data");
            sheet.DefaultColumnWidth = 12D;
            sheet.Cells[1, 1].Value = "Name";
            sheet.Cells[2, 1].Value = "Ada";
            sheet.Cells[1, 2].Value = "Count";
            sheet.Cells[2, 2].Value = 1;
            sheet.Column(1).SetWidth(20D);
            sheet.View.FreezePanes(2, 1);
            sheet.Cells["A4:B4"].Merge = true;
            sheet.DataValidations.AddAnyValidation("A2");
            sheet.Drawings.AddPicture("logo", new MemoryStream(TestImages.Png(120, 40)));
            sheet.Tables.Add(sheet.Cells["A1:B2"], "People");
            sheet.Protect(lockEverything: true, password: "not-a-secret");
        });

        Assert.Equal<string[]>(
            [
                "sheetPr",
                "dimension",
                "sheetViews",
                "sheetFormatPr",
                "cols",
                "sheetData",
                "sheetProtection",
                "mergeCells",
                "dataValidations",
                "drawing",
                "tableParts"
            ],
            [.. written.ChildOrder("Data")]);
    }

    [Fact]
    public void A_column_width_is_marked_custom_so_Excel_honours_it()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            sheet.Column(2).SetWidth(31.5D);
        });

        Column column = written.Sheet("Data").GetFirstChild<Columns>()!.Elements<Column>().Single();

        Assert.Equal(2U, column.Min!.Value);
        Assert.Equal(2U, column.Max!.Value);
        Assert.Equal(31.5D, column.Width!.Value);
        Assert.True(column.CustomWidth!.Value);
    }

    [Fact]
    public void A_hidden_column_is_hidden_and_has_no_custom_width()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            sheet.Column(3).SetHidden(true);
        });

        Column column = written.Sheet("Data").GetFirstChild<Columns>()!.Elements<Column>().Single();

        Assert.True(column.Hidden!.Value);
        Assert.False(column.CustomWidth!.Value);
    }

    [Fact]
    public void Auto_fitting_widens_a_column_to_its_longest_value()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "short";
            sheet.Cells[2, 1].Value = "a considerably longer value";
            sheet.Cells.AutoFitColumns();
        });

        Column column = written.Sheet("Data").GetFirstChild<Columns>()!.Elements<Column>().Single();

        Assert.Equal("a considerably longer value".Length + 2D, column.Width!.Value);
    }

    [Fact]
    public void Auto_fitting_never_goes_below_the_default_width()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.DefaultColumnWidth = 15D;
            sheet.Cells[1, 1].Value = "ab";
            sheet.Cells.AutoFitColumns();
        });

        Column column = written.Sheet("Data").GetFirstChild<Columns>()!.Elements<Column>().Single();

        Assert.Equal(15D, column.Width!.Value);
    }

    [Fact]
    public void Auto_fitting_stops_at_a_width_a_person_can_still_read()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = new string('x', 500);
            sheet.Cells.AutoFitColumns();
        });

        Column column = written.Sheet("Data").GetFirstChild<Columns>()!.Elements<Column>().Single();

        Assert.Equal(80D, column.Width!.Value);
    }

    [Fact]
    public void The_default_column_width_is_reported_before_it_is_set()
    {
        using SpreadsheetPackage package = new();

        Assert.Equal(8.43D, package.Workbook.Worksheets.Add("Data").DefaultColumnWidth);
    }

    [Fact]
    public void A_sheet_can_be_hidden_and_unhidden()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        Assert.Equal(SpreadsheetWorksheetVisibility.Visible, sheet.Visibility);

        sheet.Visibility = SpreadsheetWorksheetVisibility.VeryHidden;

        Assert.Equal(SpreadsheetWorksheetVisibility.VeryHidden, sheet.Visibility);

        sheet.Visibility = SpreadsheetWorksheetVisibility.Visible;

        Assert.Equal(SpreadsheetWorksheetVisibility.Visible, sheet.Visibility);
    }

    [Fact]
    public void Hiding_a_sheet_survives_the_round_trip()
    {
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Visibility = SpreadsheetWorksheetVisibility.Hidden;
        });

        Sheet sheet = written.Workbook.Sheets!.Elements<Sheet>().Single();

        Assert.Equal(SheetStateValues.Hidden, sheet.State!.Value);
    }

    [Fact]
    public void A_selection_is_remembered()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        Assert.Null(sheet.SelectedRange);

        sheet.Select("B2:C3");

        Assert.Equal("B2:C3", sheet.SelectedRange!.Address);
    }

    [Fact]
    public void Protection_stores_a_hash_and_never_the_password()
    {
        const string chosen = "correct-horse";

        using WrittenWorkbook written = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            sheet.Protect(lockEverything: true, password: chosen);
        });

        SheetProtection protection = written.Sheet("Data").GetFirstChild<SheetProtection>()!;

        Assert.True(protection.Sheet!.Value);
        Assert.NotNull(protection.Password);
        Assert.DoesNotContain(chosen, protection.Password!.Value!, StringComparison.Ordinal);
        Assert.Matches("^[0-9A-F]{4}$", protection.Password!.Value!);
    }

    [Fact]
    public void Locking_everything_locks_more_than_locking_some_of_it()
    {
        using WrittenWorkbook everything = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            sheet.Protect(lockEverything: true, password: "pw");
        });

        using WrittenWorkbook some = WrittenWorkbook.From(package =>
        {
            SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
            sheet.Cells[1, 1].Value = "x";
            sheet.Protect(lockEverything: false, password: "pw");
        });

        Assert.True(everything.Sheet("Data").GetFirstChild<SheetProtection>()!.FormatCells!.Value);
        Assert.False(some.Sheet("Data").GetFirstChild<SheetProtection>()!.FormatCells!.Value);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Protecting_a_sheet_without_a_password_is_refused(string password)
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        Assert.Throws<ArgumentException>(() => sheet.Protect(lockEverything: true, password));
    }

    [Fact]
    public void A_sheet_nobody_protected_has_no_protection_element()
    {
        // The negative control: sheetProtection with no password is worse than
        // none, because it looks protected.
        using WrittenWorkbook written = WrittenWorkbook.From(package =>
            package.Workbook.Worksheets.Add("Data").Cells[1, 1].Value = "x");

        Assert.Null(written.Sheet("Data").GetFirstChild<SheetProtection>());
    }
}

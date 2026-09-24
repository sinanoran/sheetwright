using System.Data;
using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// Ranges: addressing a rectangle, and filling one from a collection.
/// </summary>
/// <remarks>
/// <c>LoadFromCollection</c> is the method an export actually calls, so the
/// interesting questions are the ones a happy path never asks — what an empty
/// collection writes, where the first data row lands when headers were printed,
/// and what happens to a property holding null.
/// </remarks>
public sealed class SpreadsheetRangeTests
{
    private sealed record Person(string Name, int Age, DateTime Joined);

    [Fact]
    public void A_range_knows_its_own_address()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        Assert.Equal("B2:D5", sheet.Cells[2, 2, 5, 4].Address);
        Assert.Equal("B2", sheet.Cells[2, 2].Address);
    }

    [Fact]
    public void A_full_address_names_the_sheet_and_quotes_it()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Sales Data");

        Assert.Equal("'Sales Data'!A1", sheet.Cells[1, 1].FullAddress);
    }

    [Fact]
    public void A_sheet_name_containing_an_apostrophe_is_escaped_rather_than_breaking_the_reference()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Ada's Data");

        Assert.Equal("'Ada''s Data'!A1", sheet.Cells[1, 1].FullAddress);
    }

    [Fact]
    public void Reading_the_value_of_more_than_one_cell_is_refused()
    {
        // Silently returning the top-left is how a caller ships a report that
        // quietly drops every column but the first.
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        Assert.Throws<InvalidOperationException>(() => sheet.Cells["A1:B2"].Value);
        Assert.Throws<InvalidOperationException>(() => sheet.Cells["A1:B2"].Value = 1);
        Assert.Throws<InvalidOperationException>(() => sheet.Cells["A1:B2"].Formula);
    }

    [Fact]
    public void A_range_enumerates_every_cell_in_it_row_by_row()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        string[] addresses = [.. sheet.Cells["A1:B2"].Select(cell => cell.Address)];

        Assert.Equal<string[]>(["A1", "B1", "A2", "B2"], addresses);
    }

    [Fact]
    public void Enumerating_a_whole_column_stops_at_the_last_row_with_anything_in_it()
    {
        // Not at row 1,048,576, which is the number the address carries.
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        sheet.Cells[1, 1].Value = "a";
        sheet.Cells[3, 1].Value = "c";

        Assert.Equal(3, sheet.Cells["A:A"].Count());
    }

    [Fact]
    public void A_collection_becomes_rows_starting_where_the_range_starts()
    {
        Person[] people =
        [
            new("Ada", 36, new DateTime(2026, 1, 1)),
            new("Grace", 45, new DateTime(2026, 2, 1))
        ];

        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        sheet.Cells[2, 3].LoadFromCollection(people);

        Assert.Equal("Ada", ValueUnderHeader(sheet, people, "Name", 2, 3));
        Assert.Equal("Grace", ValueUnderHeader(sheet, people, "Name", 3, 3));
        Assert.Equal(36D, ValueUnderHeader(sheet, people, "Age", 2, 3));
    }

    [Fact]
    public void Printing_headers_pushes_the_data_down_one_row()
    {
        Person[] people = [new("Ada", 36, new DateTime(2026, 1, 1))];

        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        sheet.Cells[1, 1].LoadFromCollection(people, printHeaders: true);

        string[] headers = [.. Enumerable.Range(1, 3).Select(column => (string)sheet.Cells[1, column].Value!)];

        Assert.Equal<string[]>(["Age", "Joined", "Name"], [.. headers.Order(StringComparer.Ordinal)]);
        Assert.Equal("Ada", ValueUnderHeader(sheet, people, "Name", 2, 1));
    }

    [Fact]
    public void An_empty_collection_writes_headers_and_nothing_else()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        sheet.Cells[1, 1].LoadFromCollection(Array.Empty<Person>(), printHeaders: true);

        Assert.Equal(1, sheet.Dimension!.Rows);
        Assert.Equal(3, sheet.Dimension!.Columns);
    }

    [Fact]
    public void An_empty_collection_with_no_headers_writes_nothing_at_all()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        sheet.Cells[1, 1].LoadFromCollection(Array.Empty<Person>());

        Assert.Null(sheet.Dimension);
    }

    [Fact]
    public void A_null_collection_is_refused()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        Assert.Throws<ArgumentNullException>(() => sheet.Cells[1, 1].LoadFromCollection<Person>(null!));
    }

    [Fact]
    public void A_data_table_becomes_rows_with_its_column_names_as_headers()
    {
        using DataTable table = new();
        table.Columns.Add("Reference", typeof(string));
        table.Columns.Add("Amount", typeof(decimal));
        table.Rows.Add("INV-1", 10.5M);
        table.Rows.Add("INV-2", 20M);

        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        sheet.Cells[1, 1].LoadFromDataTable(table, printHeaders: true);

        Assert.Equal("Reference", sheet.Cells[1, 1].Value);
        Assert.Equal("Amount", sheet.Cells[1, 2].Value);
        Assert.Equal("INV-1", sheet.Cells[2, 1].Value);
        Assert.Equal(10.5D, sheet.Cells[2, 2].Value);
        Assert.Equal(20D, sheet.Cells[3, 2].Value);
    }

    [Fact]
    public void A_missing_value_in_a_data_table_leaves_the_cell_empty()
    {
        // DBNull is what a nullable database column arrives as, and writing the
        // words "System.DBNull" into a customer's report is the failure mode.
        using DataTable table = new();
        table.Columns.Add("Reference", typeof(string));
        table.Rows.Add(DBNull.Value);

        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");
        sheet.Cells[1, 1].LoadFromDataTable(table);

        Assert.Null(sheet.Cells[1, 1].Value);
    }

    [Fact]
    public void A_null_data_table_is_refused()
    {
        using SpreadsheetPackage package = new();
        SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Data");

        Assert.Throws<ArgumentNullException>(() => sheet.Cells[1, 1].LoadFromDataTable(null!));
    }

    /// <summary>
    /// The value beneath a named header, found rather than assumed.
    /// </summary>
    /// <remarks>
    /// <c>LoadFromCollection</c> lays columns out in the order
    /// <see cref="System.Type.GetProperties()"/> returns, which the runtime does
    /// not guarantee. Asserting a fixed column order would test the reflection
    /// implementation; what the caller is owed is that each value sits under its
    /// own name.
    /// </remarks>
    private static object? ValueUnderHeader(
        SpreadsheetWorksheet sheet,
        IReadOnlyCollection<Person> people,
        string header,
        int row,
        int firstColumn)
    {
        string[] names = [.. typeof(Person).GetProperties().Select(property => property.Name)];
        int offset = Array.IndexOf(names, header);

        Assert.True(offset >= 0, $"{typeof(Person).Name} has no property named {header}.");
        Assert.NotEmpty(people);

        return sheet.Cells[row, firstColumn + offset].Value;
    }
}

using Xunit;

namespace Sheetwright.Tests;

/// <summary>
/// The translation between "B7" and a pair of numbers.
/// </summary>
/// <remarks>
/// Every other part of the library is built on this one, and it is the part with
/// no obvious failure mode: an off-by-one in the base-26 conversion produces a
/// workbook that opens perfectly and has its data in the wrong column. The
/// boundaries are what these check — A, Z, AA, and XFD, which is where Excel's
/// grid stops.
/// </remarks>
public sealed class SpreadsheetAddressTests
{
    [Theory]
    [InlineData(1, "A")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(52, "AZ")]
    [InlineData(53, "BA")]
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    [InlineData(16_384, "XFD")]
    public void A_column_number_becomes_its_letters(int column, string expected)
    {
        Assert.Equal(expected, SpreadsheetAddress.GetColumnName(column));
        Assert.Equal(column, SpreadsheetAddress.GetColumnNumber(expected));
    }

    [Fact]
    public void The_first_column_is_one_and_not_zero()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => SpreadsheetAddress.GetColumnName(0));
    }

    [Fact]
    public void There_is_no_column_after_XFD()
    {
        Assert.Throws<ArgumentOutOfRangeException>(
            () => SpreadsheetAddress.GetColumnName(SpreadsheetAddress.MaximumColumn + 1));
        Assert.Throws<FormatException>(() => SpreadsheetAddress.GetColumnNumber("XFE"));
    }

    [Fact]
    public void A_column_name_far_past_the_end_is_refused_rather_than_overflowing()
    {
        // checked() would throw OverflowException somewhere around the seventh
        // letter; the answer a caller needs is that there is no such column.
        Assert.Throws<FormatException>(() => SpreadsheetAddress.GetColumnNumber("AAAAAAAAAA"));
    }

    [Theory]
    [InlineData("A0")]
    [InlineData("A1048577")]
    public void A_cell_outside_the_grid_is_refused_when_it_is_named(string address)
    {
        // It used to be accepted, and written out as <row r="0"><c r="A0">. The
        // package was produced without error and Excel would not open it.
        Assert.Throws<ArgumentOutOfRangeException>(() => SpreadsheetAddress.Parse(address));
    }

    [Theory]
    [InlineData("C1:A1")]
    [InlineData("A3:A1")]
    public void A_range_that_ends_before_it_begins_is_refused(string address)
    {
        // Every loop over an inverted range runs zero times, so styling or
        // merging one did nothing at all and said nothing about it.
        Assert.Throws<ArgumentOutOfRangeException>(() => SpreadsheetAddress.Parse(address));
    }

    [Fact]
    public void A_column_name_is_read_whatever_its_case()
    {
        Assert.Equal(27, SpreadsheetAddress.GetColumnNumber("aa"));
    }

    [Fact]
    public void Something_that_is_not_a_column_name_is_refused()
    {
        Assert.Throws<FormatException>(() => SpreadsheetAddress.GetColumnNumber("A1"));
    }

    [Fact]
    public void A_single_cell_parses_to_itself()
    {
        SpreadsheetAddress address = SpreadsheetAddress.Parse("B7");

        Assert.Equal(7, address.FromRow);
        Assert.Equal(2, address.FromColumn);
        Assert.Equal(7, address.ToRow);
        Assert.Equal(2, address.ToColumn);
        Assert.Equal("B7", address.Reference);
    }

    [Fact]
    public void A_range_keeps_both_ends()
    {
        SpreadsheetAddress address = SpreadsheetAddress.Parse("B2:D5");

        Assert.Equal((2, 2, 5, 4), (address.FromRow, address.FromColumn, address.ToRow, address.ToColumn));
        Assert.Equal("B2:D5", address.Reference);
    }

    [Fact]
    public void Absolute_markers_are_a_notation_and_not_a_difference()
    {
        Assert.Equal("B2:D5", SpreadsheetAddress.Parse("$B$2:$D$5").Reference);
    }

    [Fact]
    public void A_sheet_qualifier_is_dropped_because_the_worksheet_is_already_known()
    {
        // The exclamation mark is found from the right: a quoted sheet name may
        // contain one of its own.
        Assert.Equal("A1", SpreadsheetAddress.Parse("'Some!Sheet'!A1").Reference);
    }

    [Fact]
    public void A_whole_column_spans_every_row_Excel_has()
    {
        SpreadsheetAddress address = SpreadsheetAddress.Parse("B:D");

        Assert.True(address.IsWholeColumn);
        Assert.False(address.IsWholeRow);
        Assert.Equal(1, address.FromRow);
        Assert.Equal(1_048_576, address.ToRow);
        Assert.Equal("B:D", address.Reference);
    }

    [Fact]
    public void A_whole_row_spans_every_column_Excel_has()
    {
        SpreadsheetAddress address = SpreadsheetAddress.Parse("2:4");

        Assert.True(address.IsWholeRow);
        Assert.False(address.IsWholeColumn);
        Assert.Equal(1, address.FromColumn);
        Assert.Equal(16_384, address.ToColumn);
        Assert.Equal("2:4", address.Reference);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void An_empty_address_is_refused(string address)
    {
        Assert.Throws<ArgumentException>(() => SpreadsheetAddress.Parse(address));
    }

    [Theory]
    [InlineData("hello")]
    [InlineData("A1:")]
    [InlineData("1A")]
    [InlineData("A1:B")]
    public void An_address_that_is_not_one_is_refused(string address)
    {
        Assert.Throws<FormatException>(() => SpreadsheetAddress.Parse(address));
    }
}

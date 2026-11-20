namespace Sheetwright;

/// <summary>What a cell read by <see cref="SpreadsheetWorkbookReader"/> or <see cref="SpreadsheetCsvReader"/> holds.</summary>
public enum SpreadsheetCellKind
{
    /// <summary>Nothing: no cell, a cell with no value, or a formula with no cached value.</summary>
    Empty,

    /// <summary>Text. Every non-empty cell of a CSV file is text.</summary>
    Text,

    /// <summary>A number, as a <see cref="decimal"/>.</summary>
    Number,

    /// <summary>A true or false.</summary>
    Boolean,

    /// <summary>
    /// A date or a time of day: a number whose cell format is a date format, or a date stored as such. A
    /// time with no date has the date 30 December 1899 (1 January 1904 in a workbook that counts from 1904).
    /// The serial number the file stores is in <see cref="SpreadsheetCellValue.Number"/>.
    /// </summary>
    Date,

    /// <summary>
    /// An error value such as <c>#DIV/0!</c>; <see cref="SpreadsheetCellValue.Text"/> holds its spelling.
    /// </summary>
    Error,
}

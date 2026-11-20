namespace Sheetwright;

/// <summary>
/// One row of a file read by <see cref="SpreadsheetWorkbookReader"/> or <see cref="SpreadsheetCsvReader"/>.
/// </summary>
/// <remarks>
/// <para>
/// A row is a fresh object that nothing else holds, so it may be kept after the next one is read. Its cells
/// are numbered from 0. <b>A row ends at its last cell that has a value</b>, in both readers: a formatted but
/// empty cell, or a CSV record that ends in a separator, does not make it longer. Asking for a column past that
/// is an empty cell, not an error, because rows of one file have different lengths.
/// </para>
/// <para>
/// Only the cells that have a value are stored, so a row with one cell in column 16,000 costs one cell and not
/// sixteen thousand. <see cref="Cells"/> builds the whole run of cells up to <see cref="Count"/> the first time
/// it is asked for; code that reads a few columns of every row should use the indexer.
/// </para>
/// </remarks>
public sealed class SpreadsheetRow
{
    private readonly int[] _columns;
    private readonly SpreadsheetCellValue[] _values;
    private SpreadsheetCellValue[]? _dense;

    /// <summary>
    /// A row made by hand, for code that is tested with rows. Cells with no value are dropped, so the row
    /// ends at the last one that has a value.
    /// </summary>
    /// <param name="rowNumber">The row's number, from 1.</param>
    /// <param name="cells">The cells from the first column.</param>
    public SpreadsheetRow(int rowNumber, IReadOnlyList<SpreadsheetCellValue> cells)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(rowNumber, 1);
        ArgumentNullException.ThrowIfNull(cells);

        List<int> columns = [];
        List<SpreadsheetCellValue> values = [];
        for (int column = 0; column < cells.Count; column++)
        {
            if (!cells[column].IsEmpty)
            {
                columns.Add(column);
                values.Add(cells[column]);
            }
        }

        RowNumber = rowNumber;
        _columns = [.. columns];
        _values = [.. values];
    }

    internal SpreadsheetRow(int rowNumber, int[] columns, SpreadsheetCellValue[] values)
    {
        RowNumber = rowNumber;
        _columns = columns;
        _values = values;
    }

    /// <summary>
    /// The row's number in the file, counted from 1: the row number a spreadsheet shows, or the record number
    /// of a CSV (a quoted field with a line break in it is still one record). Rows skipped as empty leave gaps.
    /// </summary>
    public int RowNumber { get; }

    /// <summary>How many cells the row has: one more than the column of its last cell with a value.</summary>
    public int Count => _columns.Length == 0 ? 0 : _columns[^1] + 1;

    /// <summary>
    /// The cells, from the first column to the last one that has a value. The list is built when first asked
    /// for and holds <see cref="Count"/> cells.
    /// </summary>
    public IReadOnlyList<SpreadsheetCellValue> Cells => _dense ??= BuildDense();

    /// <summary>The cell in a column counted from 0, or an empty cell past the end of the row.</summary>
    public SpreadsheetCellValue this[int column]
    {
        get
        {
            ArgumentOutOfRangeException.ThrowIfNegative(column);
            int index = Array.BinarySearch(_columns, column);
            return index >= 0 ? _values[index] : SpreadsheetCellValue.Empty;
        }
    }

    /// <summary>Whether no cell of the row holds anything.</summary>
    public bool IsEmpty => _values.Length == 0;

    private SpreadsheetCellValue[] BuildDense()
    {
        SpreadsheetCellValue[] cells = new SpreadsheetCellValue[Count];
        for (int index = 0; index < _columns.Length; index++)
        {
            cells[_columns[index]] = _values[index];
        }

        return cells;
    }
}

namespace Sheetwright;

/// <summary>
/// One row of a file read by <see cref="SpreadsheetWorkbookReader"/> or <see cref="SpreadsheetCsvReader"/>.
/// </summary>
/// <remarks>
/// A row is a fresh object that nothing else holds, so it may be kept after the next one is read. Its cells
/// are numbered from 0 and run to the last cell the row has; asking for a column past that is an empty cell,
/// not an error, because rows of one file have different lengths.
/// </remarks>
public sealed class SpreadsheetRow
{
    internal SpreadsheetRow(int rowNumber, SpreadsheetCellValue[] cells)
    {
        RowNumber = rowNumber;
        Cells = cells;
    }

    /// <summary>
    /// The row's number in the file, counted from 1: the row number a spreadsheet shows, or the record number
    /// of a CSV (a quoted field with a line break in it is still one record). Rows skipped as empty leave gaps.
    /// </summary>
    public int RowNumber { get; }

    /// <summary>The cells, from the first column to the last one the row has.</summary>
    public IReadOnlyList<SpreadsheetCellValue> Cells { get; }

    /// <summary>How many cells the row has.</summary>
    public int Count => Cells.Count;

    /// <summary>The cell in a column counted from 0, or an empty cell past the end of the row.</summary>
    public SpreadsheetCellValue this[int column]
    {
        get
        {
            ArgumentOutOfRangeException.ThrowIfNegative(column);
            return column < Cells.Count ? Cells[column] : SpreadsheetCellValue.Empty;
        }
    }

    /// <summary>Whether no cell of the row holds anything.</summary>
    public bool IsEmpty
    {
        get
        {
            foreach (SpreadsheetCellValue cell in Cells)
            {
                if (!cell.IsEmpty)
                {
                    return false;
                }
            }

            return true;
        }
    }
}

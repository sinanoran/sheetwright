using System.Collections;
using System.Data;
using System.Reflection;

namespace Sheetwright;

public sealed class SpreadsheetRange : IEnumerable<SpreadsheetRange>
{
    private readonly SpreadsheetWorksheet _worksheet;
    private readonly SpreadsheetAddress _address;
    private SpreadsheetRangePosition? _start;
    private SpreadsheetRangePosition? _end;
    private SpreadsheetRangeStyle? _style;
    private SpreadsheetRangeDataValidation? _dataValidation;

    internal SpreadsheetRange(SpreadsheetWorksheet worksheet, SpreadsheetAddress address)
    {
        _worksheet = worksheet;
        _address = address;
    }

    public string Address => _address.Reference;

    public string FullAddress => $"'{_worksheet.Name.Replace("'", "''")}'!{_address.Reference}";

    public SpreadsheetRangePosition Start => _start ??= new SpreadsheetRangePosition(_address.FromRow, _address.FromColumn);

    public SpreadsheetRangePosition End => _end ??= new SpreadsheetRangePosition(_address.ToRow, _address.ToColumn);

    public SpreadsheetRangeStyle Style => _style ??= new SpreadsheetRangeStyle(this);

    public SpreadsheetRangeDataValidation DataValidation => _dataValidation ??= new SpreadsheetRangeDataValidation(this);

    public object? Value
    {
        get
        {
            EnsureSingleCell();
            return _worksheet.GetValue(_address.FromRow, _address.FromColumn);
        }
        set
        {
            EnsureSingleCell();
            _worksheet.SetValue(_address.FromRow, _address.FromColumn, value);
        }
    }

    public string? Formula
    {
        get
        {
            EnsureSingleCell();
            return _worksheet.GetFormula(_address.FromRow, _address.FromColumn);
        }
        set
        {
            EnsureSingleCell();
            _worksheet.SetFormula(_address.FromRow, _address.FromColumn, value);
        }
    }

    public bool Merge
    {
        get => _worksheet.IsMerged(_address);
        set => _worksheet.SetMerged(_address, value);
    }

    public void LoadFromCollection<T>(IReadOnlyCollection<T> source, bool printHeaders = false)
    {
        if (source is null)
        {
            throw new ArgumentNullException(nameof(source));
        }

        PropertyInfo[] properties = typeof(T).GetProperties();
        int row = _address.FromRow;

        if (printHeaders)
        {
            for (int column = 0; column < properties.Length; column++)
            {
                _worksheet.SetValue(row, _address.FromColumn + column, properties[column].Name);
            }
            row++;
        }

        foreach (T item in source)
        {
            for (int column = 0; column < properties.Length; column++)
            {
                _worksheet.SetValue(row, _address.FromColumn + column, properties[column].GetValue(item));
            }
            row++;
        }
    }

    public void LoadFromDataTable(DataTable table, bool printHeaders = false)
    {
        if (table is null)
        {
            throw new ArgumentNullException(nameof(table));
        }

        int row = _address.FromRow;

        if (printHeaders)
        {
            for (int column = 0; column < table.Columns.Count; column++)
            {
                _worksheet.SetValue(row, _address.FromColumn + column, table.Columns[column].ColumnName);
            }
            row++;
        }

        foreach (DataRow dataRow in table.Rows)
        {
            for (int column = 0; column < table.Columns.Count; column++)
            {
                _worksheet.SetValue(row, _address.FromColumn + column, dataRow[column]);
            }
            row++;
        }
    }

    internal SpreadsheetAddress InnerAddress => _address;

    internal SpreadsheetWorksheet Worksheet => _worksheet;

    internal void ApplyStyle(SpreadsheetStylePatch patch)
    {
        _worksheet.ApplyStyle(_address, patch);
    }

    internal void ApplyBorderAround(SpreadsheetBorderStyle borderStyle, string color)
    {
        _worksheet.ApplyBorderAround(_address, borderStyle, color);
    }

    public IEnumerator<SpreadsheetRange> GetEnumerator()
    {
        int maxRow = _address.IsWholeColumn ? _worksheet.Dimension?.End.Row ?? 0 : _address.ToRow;
        int maxColumn = _address.IsWholeRow ? _worksheet.Dimension?.End.Column ?? 0 : _address.ToColumn;

        for (int row = _address.FromRow; row <= maxRow; row++)
        {
            for (int column = _address.FromColumn; column <= maxColumn; column++)
            {
                yield return new SpreadsheetRange(_worksheet, new SpreadsheetAddress(row, column, row, column));
            }
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    private void EnsureSingleCell()
    {
        if (_address.FromRow != _address.ToRow || _address.FromColumn != _address.ToColumn)
        {
            throw new InvalidOperationException("This operation is only valid for a single cell.");
        }
    }
}

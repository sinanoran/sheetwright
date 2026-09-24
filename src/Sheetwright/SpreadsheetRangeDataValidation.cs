namespace Sheetwright;

public sealed class SpreadsheetRangeDataValidation
{
    private readonly SpreadsheetRange _range;

    internal SpreadsheetRangeDataValidation(SpreadsheetRange range)
    {
        _range = range;
    }

    public SpreadsheetAnyDataValidation AddAnyDataValidation()
    {
        return _range.InnerAddress.IsWholeColumn
            ? throw new InvalidOperationException("Whole-column data validation is not supported.")
            : _range.Worksheet.DataValidations.AddAnyValidation(_range.Address);
    }
}

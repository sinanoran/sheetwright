namespace Sheetwright;

/// <summary>
/// Conditional formats added through a range, which take the range's address.
/// </summary>
public sealed class SpreadsheetRangeConditionalFormatting
{
    private readonly SpreadsheetRange _range;

    internal SpreadsheetRangeConditionalFormatting(SpreadsheetRange range)
    {
        _range = range;
    }

    public SpreadsheetCellValueConditionalFormattingRule AddCellValue(SpreadsheetConditionalFormattingOperator @operator)
    {
        return Rules.AddCellValue(_range.Address, @operator);
    }

    public SpreadsheetFormulaConditionalFormattingRule AddFormula()
    {
        return Rules.AddFormula(_range.Address);
    }

    public SpreadsheetDuplicateValuesConditionalFormattingRule AddDuplicateValues()
    {
        return Rules.AddDuplicateValues(_range.Address);
    }

    public SpreadsheetUniqueValuesConditionalFormattingRule AddUniqueValues()
    {
        return Rules.AddUniqueValues(_range.Address);
    }

    public SpreadsheetColorScaleConditionalFormattingRule AddTwoColorScale()
    {
        return Rules.AddTwoColorScale(_range.Address);
    }

    public SpreadsheetColorScaleConditionalFormattingRule AddThreeColorScale()
    {
        return Rules.AddThreeColorScale(_range.Address);
    }

    private SpreadsheetConditionalFormattingCollection Rules => _range.Worksheet.ConditionalFormatting;
}

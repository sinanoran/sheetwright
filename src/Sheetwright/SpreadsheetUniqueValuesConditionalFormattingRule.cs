using DocumentFormat.OpenXml.Spreadsheet;

namespace Sheetwright;

/// <summary>
/// Formats every cell whose value appears exactly once in the range.
/// </summary>
public sealed class SpreadsheetUniqueValuesConditionalFormattingRule : SpreadsheetStyledConditionalFormattingRule
{
    internal SpreadsheetUniqueValuesConditionalFormattingRule(string address) : base(address)
    {
    }

    internal override ConditionalFormatValues RuleType => ConditionalFormatValues.UniqueValues;
}

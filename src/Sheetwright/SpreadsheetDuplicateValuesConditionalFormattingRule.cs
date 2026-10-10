using DocumentFormat.OpenXml.Spreadsheet;

namespace Sheetwright;

/// <summary>
/// Formats every cell whose value appears more than once in the range.
/// </summary>
public sealed class SpreadsheetDuplicateValuesConditionalFormattingRule : SpreadsheetStyledConditionalFormattingRule
{
    internal SpreadsheetDuplicateValuesConditionalFormattingRule(string address) : base(address)
    {
    }

    internal override ConditionalFormatValues RuleType => ConditionalFormatValues.DuplicateValues;
}

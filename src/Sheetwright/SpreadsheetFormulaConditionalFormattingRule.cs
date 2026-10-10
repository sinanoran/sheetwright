using DocumentFormat.OpenXml.Spreadsheet;

namespace Sheetwright;

/// <summary>
/// Formats a cell when a formula is true for it — Excel's "use a formula to determine which cells
/// to format".
/// </summary>
/// <remarks>
/// The formula is written as it would be for the top-left cell of the range, and Excel shifts its
/// relative references for every other cell. <c>$D2&lt;TODAY()</c> on <c>A2:F100</c> highlights each
/// whole row whose column D is in the past: the column is pinned, the row is not.
/// </remarks>
public sealed class SpreadsheetFormulaConditionalFormattingRule : SpreadsheetStyledConditionalFormattingRule
{
    internal SpreadsheetFormulaConditionalFormattingRule(string address) : base(address)
    {
    }

    /// <summary>The formula, with or without a leading <c>=</c>.</summary>
    public SpreadsheetFormula Formula { get; } = new();

    internal override ConditionalFormatValues RuleType => ConditionalFormatValues.Expression;

    private string FormulaText =>
        SpreadsheetTextSanitizer.Sanitize(Formula.ExcelFormula ?? string.Empty).Trim().TrimStart('=');

    internal override void Validate()
    {
        if (FormulaText.Length == 0)
        {
            throw new InvalidOperationException(
                $"The conditional format on {Address} has no formula. Set Formula.ExcelFormula.");
        }
    }

    internal override void Configure(ConditionalFormattingRule rule)
    {
        rule.Append(new Formula(FormulaText));
    }
}

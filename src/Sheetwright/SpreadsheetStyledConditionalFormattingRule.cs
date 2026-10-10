using DocumentFormat.OpenXml.Spreadsheet;

namespace Sheetwright;

/// <summary>
/// A conditional format that changes how matching cells look: font, fill, border, number format.
/// </summary>
/// <remarks>
/// Everything except a colour scale. A colour scale computes a colour per cell from where its value
/// falls in the range, so it has no fixed formatting to apply and no <see cref="Style"/>.
/// </remarks>
public abstract class SpreadsheetStyledConditionalFormattingRule : SpreadsheetConditionalFormattingRule
{
    internal SpreadsheetStyledConditionalFormattingRule(string address) : base(address)
    {
    }

    public SpreadsheetConditionalFormattingStyle Style { get; } = new();

    /// <summary>
    /// When the rule matches a cell, no lower-priority rule is applied to that cell.
    /// </summary>
    public bool StopIfTrue { get; set; }

    internal abstract ConditionalFormatValues RuleType { get; }

    /// <summary>Adds what is particular to the rule type: its operator, its formulas.</summary>
    internal virtual void Configure(ConditionalFormattingRule rule)
    {
    }

    internal override void Validate()
    {
    }

    internal override ConditionalFormattingRule ToOpenXml(int priority, SpreadsheetStyleRepository styles)
    {
        ConditionalFormattingRule rule = new()
        {
            Type = RuleType,
            Priority = priority
        };

        // A rule that formats nothing is written without a dxfId at all, which is how Excel
        // writes one. Pointing it at an empty dxf would cost an entry in the table and say the
        // same thing less plainly.
        if (!Style.Format.IsEmpty)
        {
            rule.FormatId = styles.GetDifferentialFormatIndex(Style.Format);
        }

        if (StopIfTrue)
        {
            rule.StopIfTrue = true;
        }

        Configure(rule);
        return rule;
    }
}

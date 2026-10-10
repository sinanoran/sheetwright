using DocumentFormat.OpenXml.Spreadsheet;

namespace Sheetwright;

/// <summary>
/// Formats a cell by comparing its value: greater than a number, between two dates, equal to a word.
/// </summary>
/// <remarks>
/// <see cref="Formula2"/> is the upper bound of <see cref="SpreadsheetConditionalFormattingOperator.Between"/>
/// and <see cref="SpreadsheetConditionalFormattingOperator.NotBetween"/>, and is refused for every
/// other operator rather than ignored: a second bound that was set and silently dropped is a rule
/// that does something other than what its author read it as doing.
/// </remarks>
public sealed class SpreadsheetCellValueConditionalFormattingRule : SpreadsheetStyledConditionalFormattingRule
{
    internal SpreadsheetCellValueConditionalFormattingRule(string address, SpreadsheetConditionalFormattingOperator @operator)
        : base(address)
    {
        if (!Enum.IsDefined(@operator))
        {
            throw new ArgumentOutOfRangeException(nameof(@operator), @operator, "The comparison is not one Excel has.");
        }

        Operator = @operator;
    }

    public SpreadsheetConditionalFormattingOperator Operator { get; }

    public SpreadsheetConditionalFormattingFormula Formula { get; } = new();

    public SpreadsheetConditionalFormattingFormula Formula2 { get; } = new();

    internal override ConditionalFormatValues RuleType => ConditionalFormatValues.CellIs;

    private bool TakesTwoBounds => Operator is SpreadsheetConditionalFormattingOperator.Between
        or SpreadsheetConditionalFormattingOperator.NotBetween;

    internal override void Validate()
    {
        if (!Formula.IsSet)
        {
            throw new InvalidOperationException(
                $"The conditional format on {Address} has nothing to compare with. Set Formula.Value or Formula.ExcelFormula.");
        }

        if (TakesTwoBounds && !Formula2.IsSet)
        {
            throw new InvalidOperationException(
                $"The conditional format on {Address} is {Operator} and needs an upper bound. Set Formula2.");
        }

        if (!TakesTwoBounds && Formula2.IsSet)
        {
            throw new InvalidOperationException(
                $"The conditional format on {Address} is {Operator}, which takes one value, but Formula2 was set.");
        }
    }

    internal override void Configure(ConditionalFormattingRule rule)
    {
        rule.Operator = Operator switch
        {
            SpreadsheetConditionalFormattingOperator.Equal => ConditionalFormattingOperatorValues.Equal,
            SpreadsheetConditionalFormattingOperator.NotEqual => ConditionalFormattingOperatorValues.NotEqual,
            SpreadsheetConditionalFormattingOperator.GreaterThan => ConditionalFormattingOperatorValues.GreaterThan,
            SpreadsheetConditionalFormattingOperator.GreaterThanOrEqual => ConditionalFormattingOperatorValues.GreaterThanOrEqual,
            SpreadsheetConditionalFormattingOperator.LessThan => ConditionalFormattingOperatorValues.LessThan,
            SpreadsheetConditionalFormattingOperator.LessThanOrEqual => ConditionalFormattingOperatorValues.LessThanOrEqual,
            SpreadsheetConditionalFormattingOperator.Between => ConditionalFormattingOperatorValues.Between,
            _ => ConditionalFormattingOperatorValues.NotBetween
        };

        rule.Append(new Formula(Formula.ToText()!));
        if (TakesTwoBounds)
        {
            rule.Append(new Formula(Formula2.ToText()!));
        }
    }
}

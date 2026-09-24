using DocumentFormat.OpenXml.Spreadsheet;

namespace Sheetwright;

public sealed class SpreadsheetDecimalDataValidation : SpreadsheetDataValidation
{
    internal SpreadsheetDecimalDataValidation(string address) : base(address)
    {
    }

    public SpreadsheetValidationFormula<decimal> Formula { get; } = new();

    public SpreadsheetValidationFormula<decimal> Formula2 { get; } = new();

    internal override DataValidationValues ValidationType => DataValidationValues.Decimal;

    internal override string? Formula1Text => Formula.ToText();

    internal override string? Formula2Text => Formula2.ToText();
}

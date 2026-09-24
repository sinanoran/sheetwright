using DocumentFormat.OpenXml.Spreadsheet;

namespace Sheetwright;

public sealed class SpreadsheetDateTimeDataValidation : SpreadsheetDataValidation
{
    internal SpreadsheetDateTimeDataValidation(string address) : base(address)
    {
    }

    public SpreadsheetValidationFormula<DateTime> Formula { get; } = new();

    public SpreadsheetValidationFormula<DateTime> Formula2 { get; } = new();

    internal override DataValidationValues ValidationType => DataValidationValues.Date;

    internal override string? Formula1Text => Formula.ToText();

    internal override string? Formula2Text => Formula2.ToText();
}

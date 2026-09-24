using DocumentFormat.OpenXml.Spreadsheet;

namespace Sheetwright;

public sealed class SpreadsheetTimeDataValidation : SpreadsheetDataValidation
{
    internal SpreadsheetTimeDataValidation(string address) : base(address)
    {
        Formula.Value = new SpreadsheetTime();
    }

    public SpreadsheetValidationFormula<SpreadsheetTime> Formula { get; } = new();

    internal override DataValidationValues ValidationType => DataValidationValues.Time;

    internal override string? Formula1Text => Formula.ToText();
}

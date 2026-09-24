using DocumentFormat.OpenXml.Spreadsheet;

namespace Sheetwright;

public sealed class SpreadsheetListDataValidation : SpreadsheetDataValidation
{
    internal SpreadsheetListDataValidation(string address) : base(address)
    {
    }

    public SpreadsheetFormula Formula { get; } = new();

    internal override DataValidationValues ValidationType => DataValidationValues.List;

    internal override string? Formula1Text => Formula.ExcelFormula;
}

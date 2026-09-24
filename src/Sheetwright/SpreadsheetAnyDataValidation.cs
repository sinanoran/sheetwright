using DocumentFormat.OpenXml.Spreadsheet;

namespace Sheetwright;

public sealed class SpreadsheetAnyDataValidation : SpreadsheetDataValidation
{
    internal SpreadsheetAnyDataValidation(string address) : base(address)
    {
    }

    internal override DataValidationValues ValidationType => DataValidationValues.None;

    internal override string? Formula1Text => null;
}

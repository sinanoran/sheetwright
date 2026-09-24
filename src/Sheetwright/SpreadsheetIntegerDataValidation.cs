using DocumentFormat.OpenXml.Spreadsheet;

namespace Sheetwright;

public sealed class SpreadsheetIntegerDataValidation : SpreadsheetDataValidation
{
    internal SpreadsheetIntegerDataValidation(string address, bool textLength) : base(address)
    {
        IsTextLength = textLength;
    }

    private bool IsTextLength { get; }

    public SpreadsheetValidationFormula<int> Formula { get; } = new();

    public SpreadsheetValidationFormula<int> Formula2 { get; } = new();

    internal override DataValidationValues ValidationType => IsTextLength ? DataValidationValues.TextLength : DataValidationValues.Whole;

    internal override string? Formula1Text => Formula.ToText();

    internal override string? Formula2Text => Formula2.ToText();
}

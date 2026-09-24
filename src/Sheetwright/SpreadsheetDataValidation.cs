using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Sheetwright;

public abstract class SpreadsheetDataValidation
{
    internal SpreadsheetDataValidation(string address)
    {
        if (address is null)
        {
            throw new ArgumentNullException(nameof(address));
        }

        Address = new SpreadsheetValidationAddress(address);
    }

    public SpreadsheetValidationAddress Address { get; }

    public SpreadsheetDataValidationWarningStyle ErrorStyle { get; set; } = SpreadsheetDataValidationWarningStyle.Stop;

    public string? PromptTitle { get; set; }

    public string? Prompt { get; set; }

    public bool ShowInputMessage { get; set; }

    public string? ErrorTitle { get; set; }

    public string? Error { get; set; }

    public bool ShowErrorMessage { get; set; }

    public SpreadsheetDataValidationOperator Operator { get; set; } = SpreadsheetDataValidationOperator.Between;

    internal abstract DataValidationValues ValidationType { get; }

    internal abstract string? Formula1Text { get; }

    internal virtual string? Formula2Text => null;

    internal DataValidation ToOpenXml()
    {
        DataValidation validation = new()
        {
            Type = ValidationType,
            SequenceOfReferences = new ListValue<StringValue> { InnerText = Address.Address },
            ShowInputMessage = ShowInputMessage,
            ShowErrorMessage = ShowErrorMessage,
            PromptTitle = SpreadsheetTextSanitizer.SanitizeNullable(PromptTitle),
            Prompt = SpreadsheetTextSanitizer.SanitizeNullable(Prompt),
            ErrorTitle = SpreadsheetTextSanitizer.SanitizeNullable(ErrorTitle),
            Error = SpreadsheetTextSanitizer.SanitizeNullable(Error),
            ErrorStyle = ErrorStyle switch
            {
                SpreadsheetDataValidationWarningStyle.Warning => DataValidationErrorStyleValues.Warning,
                SpreadsheetDataValidationWarningStyle.Information => DataValidationErrorStyleValues.Information,
                _ => DataValidationErrorStyleValues.Stop
            }
        };

        if (ValidationType != DataValidationValues.None && ValidationType != DataValidationValues.List)
        {
            validation.Operator = Operator switch
            {
                SpreadsheetDataValidationOperator.GreaterThan => DataValidationOperatorValues.GreaterThan,
                SpreadsheetDataValidationOperator.LessThanOrEqual => DataValidationOperatorValues.LessThanOrEqual,
                _ => DataValidationOperatorValues.Between
            };
        }

        string? formula1Text = Formula1Text;
        if (formula1Text is not null && formula1Text.Length > 0)
        {
            validation.Append(new Formula1(SpreadsheetTextSanitizer.Sanitize(formula1Text)));
        }

        string? formula2Text = Formula2Text;
        if (formula2Text is not null && formula2Text.Length > 0)
        {
            validation.Append(new Formula2(SpreadsheetTextSanitizer.Sanitize(formula2Text)));
        }

        return validation;
    }
}

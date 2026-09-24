using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;

namespace Sheetwright;

public sealed class SpreadsheetDataValidationCollection : IEnumerable<SpreadsheetDataValidation>
{
    private readonly List<SpreadsheetDataValidation> _validations = new();

    public SpreadsheetIntegerDataValidation AddIntegerValidation(string address)
    {
        return Add(new SpreadsheetIntegerDataValidation(address, textLength: false));
    }

    public SpreadsheetDecimalDataValidation AddDecimalValidation(string address)
    {
        return Add(new SpreadsheetDecimalDataValidation(address));
    }

    public SpreadsheetDateTimeDataValidation AddDateTimeValidation(string address)
    {
        return Add(new SpreadsheetDateTimeDataValidation(address));
    }

    public SpreadsheetTimeDataValidation AddTimeValidation(string address)
    {
        return Add(new SpreadsheetTimeDataValidation(address));
    }

    public SpreadsheetIntegerDataValidation AddTextLengthValidation(string address)
    {
        return Add(new SpreadsheetIntegerDataValidation(address, textLength: true));
    }

    public SpreadsheetAnyDataValidation AddAnyValidation(string address)
    {
        return Add(new SpreadsheetAnyDataValidation(address));
    }

    public SpreadsheetListDataValidation AddListValidation(string address)
    {
        return Add(new SpreadsheetListDataValidation(address));
    }

    internal void Commit(Worksheet worksheet)
    {
        worksheet.RemoveAllChildren<DataValidations>();
        if (_validations.Count == 0)
        {
            return;
        }

        DataValidations validations = new() { Count = (uint)_validations.Count };
        foreach (SpreadsheetDataValidation validation in _validations)
        {
            validations.Append(validation.ToOpenXml());
        }

        Drawing? drawing = worksheet.GetFirstChild<Drawing>();
        if (drawing is null)
        {
            worksheet.Append(validations);
        }
        else
        {
            worksheet.InsertBefore(validations, drawing);
        }
    }

    public IEnumerator<SpreadsheetDataValidation> GetEnumerator()
    {
        return _validations.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    private T Add<T>(T validation) where T : SpreadsheetDataValidation
    {
        _validations.Add(validation);
        return validation;
    }
}

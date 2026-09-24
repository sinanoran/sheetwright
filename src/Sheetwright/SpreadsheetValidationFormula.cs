using System.Globalization;

namespace Sheetwright;

public sealed class SpreadsheetValidationFormula<T>
{
    private T? _value;

    // T is unconstrained, so for a value type `T?` is an annotation only and _value can never be
    // null. Track assignment explicitly, otherwise an unset bound still writes out its default and
    // Excel receives a formula the caller never asked for.
    private bool _isSet;

    public T? Value
    {
        get => _value;
        set
        {
            _value = value;
            _isSet = value is not null;
        }
    }

    internal string? ToText()
    {
        if (!_isSet)
        {
            return null;
        }

        object? value = Value;
        if (value is null)
        {
            return null;
        }

        if (value is DateTime dateTime)
        {
            return dateTime.ToOADate().ToString(CultureInfo.InvariantCulture);
        }

        if (value is SpreadsheetTime time)
        {
            return time.ToSerialValue();
        }

        return Convert.ToString(value, CultureInfo.InvariantCulture);
    }
}

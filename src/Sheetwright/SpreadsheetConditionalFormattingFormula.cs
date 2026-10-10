using System.Globalization;

namespace Sheetwright;

/// <summary>
/// One side of a cell-value comparison: either a value, or a formula.
/// </summary>
/// <remarks>
/// <para>
/// A value is typed exactly as a cell is typed: a number is a number, a <see cref="DateTime"/> is
/// its serial number, a <see cref="bool"/> is <c>TRUE</c> or <c>FALSE</c>, and text — or anything
/// else, through <see cref="Convert.ToString(object, IFormatProvider)"/> — is a quoted string. All
/// of it under <see cref="CultureInfo.InvariantCulture"/>, because the file format spells
/// one-and-a-half <c>1.5</c> whoever wrote it. A Turkish machine writing <c>1,5</c> would hand
/// Excel a formula it cannot read.
/// </para>
/// <para>
/// A formula is written as given, which is how a rule compares against another cell:
/// <c>ExcelFormula = "$B$1"</c>. Text set as a <see cref="Value"/> is never treated as one — it is
/// quoted, so <c>"=A1"</c> compares with the three characters and not with the cell.
/// </para>
/// <para>
/// Setting either one clears the other; the comparison has one operand, not two.
/// </para>
/// </remarks>
public sealed class SpreadsheetConditionalFormattingFormula
{
    private object? _value;
    private string? _excelFormula;

    /// <summary>
    /// The value to compare with. <see langword="null"/> clears it.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The value is <c>NaN</c> or an infinity, which a formula has no way to spell.
    /// </exception>
    public object? Value
    {
        get => _value;
        set
        {
            if ((value is double number && (double.IsNaN(number) || double.IsInfinity(number)))
                || (value is float single && (float.IsNaN(single) || float.IsInfinity(single))))
            {
                // A cell holding NaN is written as text so that the workbook still opens. A
                // comparison against the text "NaN" would quietly never match, which is worse
                // than being told.
                throw new ArgumentOutOfRangeException(
                    nameof(value),
                    value,
                    "A conditional format cannot compare with NaN or an infinity; a formula has no spelling for either.");
            }

            _value = value == DBNull.Value ? null : value;
            _excelFormula = null;
        }
    }

    /// <summary>
    /// A formula to compare with, such as <c>$B$1</c> or <c>AVERAGE($C:$C)</c>. A leading
    /// <c>=</c> is optional. <see langword="null"/> or empty clears it.
    /// </summary>
    public string? ExcelFormula
    {
        get => _excelFormula;
        set
        {
            _excelFormula = value;
            _value = null;
        }
    }

    internal bool IsSet => ToText() is not null;

    internal string? ToText()
    {
        if (_excelFormula is not null)
        {
            string formula = SpreadsheetTextSanitizer.Sanitize(_excelFormula).Trim().TrimStart('=');
            return formula.Length == 0 ? null : formula;
        }

        return _value switch
        {
            null => null,
            string text => Quote(text),
            char character => Quote(character.ToString()),
            bool boolean => boolean ? "TRUE" : "FALSE",
            DateTime dateTime => dateTime.ToOADate().ToString(CultureInfo.InvariantCulture),
            byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal =>
                Convert.ToString(_value, CultureInfo.InvariantCulture),
            _ => Quote(Convert.ToString(_value, CultureInfo.InvariantCulture) ?? string.Empty)
        };
    }

    // A string constant in a formula is delimited by double quotes and escapes one by doubling it.
    private static string Quote(string text)
    {
        return "\"" + SpreadsheetTextSanitizer.Sanitize(text).Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";
    }
}

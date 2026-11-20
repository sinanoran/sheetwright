using System.Globalization;

namespace Sheetwright;

/// <summary>
/// One cell as the readers return it: its kind and its value, with no formatting.
/// </summary>
/// <remarks>
/// A formula cell holds the value Excel last calculated for it; the formula is never evaluated, and a formula
/// with no calculated value is <see cref="SpreadsheetCellKind.Empty"/>. Only the property that goes with
/// <see cref="Kind"/> means anything; the others are their defaults.
/// </remarks>
public readonly struct SpreadsheetCellValue : IEquatable<SpreadsheetCellValue>
{
    private readonly string? _text;
    private readonly decimal _number;
    private readonly DateTime _date;

    private SpreadsheetCellValue(SpreadsheetCellKind kind, string? text, decimal number, DateTime date, bool flag)
    {
        Kind = kind;
        _text = text;
        _number = number;
        _date = date;
        Boolean = flag;
    }

    /// <summary>The cell that has nothing in it.</summary>
    public static SpreadsheetCellValue Empty => default;

    /// <summary>What the cell holds.</summary>
    public SpreadsheetCellKind Kind { get; }

    /// <summary>Whether <see cref="Kind"/> is <see cref="SpreadsheetCellKind.Empty"/>.</summary>
    public bool IsEmpty => Kind == SpreadsheetCellKind.Empty;

    /// <summary>The text, for <see cref="SpreadsheetCellKind.Text"/> and <see cref="SpreadsheetCellKind.Error"/>.</summary>
    public string? Text => _text;

    /// <summary>
    /// The number, for <see cref="SpreadsheetCellKind.Number"/>. For a <see cref="SpreadsheetCellKind.Date"/> it is
    /// the serial number the file stores (days since the workbook's epoch), which is what to use for an elapsed
    /// time such as <c>[h]:mm</c>, where 36 hours is the serial 1.5 and the date part means nothing.
    /// </summary>
    public decimal Number => _number;

    /// <summary>The flag, for <see cref="SpreadsheetCellKind.Boolean"/>.</summary>
    public bool Boolean { get; }

    /// <summary>The date and time, for <see cref="SpreadsheetCellKind.Date"/>.</summary>
    public DateTime Date => _date;

    /// <summary>A text cell.</summary>
    public static SpreadsheetCellValue FromText(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        return new(SpreadsheetCellKind.Text, text, 0m, default, false);
    }

    /// <summary>A number cell.</summary>
    public static SpreadsheetCellValue FromNumber(decimal number) => new(SpreadsheetCellKind.Number, null, number, default, false);

    /// <summary>A true or false cell.</summary>
    public static SpreadsheetCellValue FromBoolean(bool value) => new(SpreadsheetCellKind.Boolean, null, 0m, default, value);

    /// <summary>A date cell.</summary>
    public static SpreadsheetCellValue FromDate(DateTime date) => new(SpreadsheetCellKind.Date, null, 0m, date, false);

    /// <summary>A date cell that also keeps the serial number the file stored it as.</summary>
    public static SpreadsheetCellValue FromDate(DateTime date, decimal serial) => new(SpreadsheetCellKind.Date, null, serial, date, false);

    /// <summary>An error cell.</summary>
    public static SpreadsheetCellValue FromError(string error)
    {
        ArgumentNullException.ThrowIfNull(error);
        return new(SpreadsheetCellKind.Error, error, 0m, default, false);
    }

    /// <summary>
    /// The value as text a person could read, in the invariant culture: numbers without a thousands
    /// separator, dates in ISO 8601, <c>true</c> and <c>false</c>. Empty is the empty string.
    /// </summary>
    public override string ToString()
    {
        return Kind switch
        {
            SpreadsheetCellKind.Text or SpreadsheetCellKind.Error => _text ?? string.Empty,
            SpreadsheetCellKind.Number => _number.ToString(CultureInfo.InvariantCulture),
            SpreadsheetCellKind.Boolean => Boolean ? "true" : "false",
            SpreadsheetCellKind.Date => _date.ToString("O", CultureInfo.InvariantCulture),
            _ => string.Empty,
        };
    }

    /// <inheritdoc />
    public bool Equals(SpreadsheetCellValue other)
    {
        return Kind == other.Kind
            && string.Equals(_text, other._text, StringComparison.Ordinal)
            && _number == other._number
            && _date == other._date
            && Boolean == other.Boolean;
    }

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is SpreadsheetCellValue other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => HashCode.Combine(Kind, _text, _number, _date, Boolean);

    /// <summary>Whether two cells hold the same.</summary>
    public static bool operator ==(SpreadsheetCellValue left, SpreadsheetCellValue right) => left.Equals(right);

    /// <summary>Whether two cells hold different values.</summary>
    public static bool operator !=(SpreadsheetCellValue left, SpreadsheetCellValue right) => !left.Equals(right);
}

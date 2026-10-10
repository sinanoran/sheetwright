using System.Globalization;

namespace Sheetwright;

/// <summary>
/// One point on a colour scale: what it is anchored to, and the colour a cell there gets.
/// </summary>
/// <remarks>
/// The anchor is set by one method per kind, each checking its own argument, so a percentile of
/// 150 fails on the line that wrote it. Separate <c>Type</c> and <c>Value</c> setters would leave
/// a window where the two disagree, and nothing to say which one the caller meant.
/// </remarks>
public sealed class SpreadsheetColorScaleValue
{
    internal SpreadsheetColorScaleValue(SpreadsheetColorScaleValueType type, string? value, string color)
    {
        Type = type;
        Value = value;
        RgbColor = color;
        Color = new SpreadsheetColor(rgb => RgbColor = rgb);
    }

    public SpreadsheetColorScaleValueType Type { get; private set; }

    /// <summary>
    /// The anchor as it is written into the file: a number in invariant form, or a formula.
    /// <see langword="null"/> for the lowest and highest value, which need none.
    /// </summary>
    public string? Value { get; private set; }

    public SpreadsheetColor Color { get; }

    internal string RgbColor { get; private set; }

    public void SetLowestValue()
    {
        Set(SpreadsheetColorScaleValueType.LowestValue, null);
    }

    public void SetHighestValue()
    {
        Set(SpreadsheetColorScaleValueType.HighestValue, null);
    }

    public void SetNumber(double number)
    {
        if (double.IsNaN(number) || double.IsInfinity(number))
        {
            throw new ArgumentOutOfRangeException(nameof(number), number, "A colour scale cannot be anchored to NaN or an infinity.");
        }

        Set(SpreadsheetColorScaleValueType.Number, number.ToString(CultureInfo.InvariantCulture));
    }

    public void SetPercent(double percent)
    {
        Set(SpreadsheetColorScaleValueType.Percent, FromZeroToOneHundred(percent, nameof(percent)));
    }

    public void SetPercentile(double percentile)
    {
        Set(SpreadsheetColorScaleValueType.Percentile, FromZeroToOneHundred(percentile, nameof(percentile)));
    }

    /// <summary>Anchors the point to a formula, with or without a leading <c>=</c>.</summary>
    public void SetFormula(string formula)
    {
        if (formula is null)
        {
            throw new ArgumentNullException(nameof(formula));
        }

        string text = SpreadsheetTextSanitizer.Sanitize(formula).Trim().TrimStart('=');
        if (text.Length == 0)
        {
            throw new ArgumentException("A colour scale formula cannot be empty.", nameof(formula));
        }

        Set(SpreadsheetColorScaleValueType.Formula, text);
    }

    private void Set(SpreadsheetColorScaleValueType type, string? value)
    {
        Type = type;
        Value = value;
    }

    private static string FromZeroToOneHundred(double value, string name)
    {
        // The negated comparison is what catches NaN, which is neither below 0 nor above 100.
        if (!(value >= 0D && value <= 100D))
        {
            throw new ArgumentOutOfRangeException(name, value, "A percentage or percentile runs from 0 to 100.");
        }

        return value.ToString(CultureInfo.InvariantCulture);
    }
}

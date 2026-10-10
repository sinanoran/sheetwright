using DocumentFormat.OpenXml.Spreadsheet;

namespace Sheetwright;

/// <summary>
/// Shades each cell by where its value falls in the range, between two or three colours.
/// </summary>
/// <remarks>
/// <para>
/// A new scale runs from red at the lowest value to green at the highest; a three-colour scale
/// passes through yellow at the 50th percentile. Each point can be re-anchored and recoloured.
/// </para>
/// <para>
/// There is no <c>Style</c> and no Stop If True: the colour is computed per cell rather than
/// applied, and Excel offers Stop If True for neither.
/// </para>
/// </remarks>
public sealed class SpreadsheetColorScaleConditionalFormattingRule : SpreadsheetConditionalFormattingRule
{
    internal const string LowColor = "FFF8696B";
    internal const string MiddleColor = "FFFFEB84";
    internal const string HighColor = "FF63BE7B";

    internal SpreadsheetColorScaleConditionalFormattingRule(string address, bool threeColors) : base(address)
    {
        LowValue = new SpreadsheetColorScaleValue(SpreadsheetColorScaleValueType.LowestValue, null, LowColor);
        MiddleValue = threeColors
            ? new SpreadsheetColorScaleValue(SpreadsheetColorScaleValueType.Percentile, "50", MiddleColor)
            : null;
        HighValue = new SpreadsheetColorScaleValue(SpreadsheetColorScaleValueType.HighestValue, null, HighColor);
    }

    public SpreadsheetColorScaleValue LowValue { get; }

    /// <summary>The middle point of a three-colour scale; <see langword="null"/> on a two-colour one.</summary>
    public SpreadsheetColorScaleValue? MiddleValue { get; }

    public SpreadsheetColorScaleValue HighValue { get; }

    internal override void Validate()
    {
        // Every anchor is checked as it is set, and every point starts with a colour.
    }

    internal override ConditionalFormattingRule ToOpenXml(int priority, SpreadsheetStyleRepository styles)
    {
        SpreadsheetColorScaleValue[] points = MiddleValue is null
            ? [LowValue, HighValue]
            : [LowValue, MiddleValue, HighValue];

        // CT_ColorScale is every cfvo first and then every colour, paired by position — not
        // cfvo, colour, cfvo, colour.
        ColorScale scale = new();
        foreach (SpreadsheetColorScaleValue point in points)
        {
            ConditionalFormatValueObject anchor = new()
            {
                Type = point.Type switch
                {
                    SpreadsheetColorScaleValueType.LowestValue => ConditionalFormatValueObjectValues.Min,
                    SpreadsheetColorScaleValueType.HighestValue => ConditionalFormatValueObjectValues.Max,
                    SpreadsheetColorScaleValueType.Number => ConditionalFormatValueObjectValues.Number,
                    SpreadsheetColorScaleValueType.Percent => ConditionalFormatValueObjectValues.Percent,
                    SpreadsheetColorScaleValueType.Percentile => ConditionalFormatValueObjectValues.Percentile,
                    _ => ConditionalFormatValueObjectValues.Formula
                }
            };

            // Assigned only when there is one. A null handed to Val is written as val="", which
            // the schema accepts and Excel never writes for the lowest or highest value.
            if (point.Value is not null)
            {
                anchor.Val = point.Value;
            }

            scale.Append(anchor);
        }

        foreach (SpreadsheetColorScaleValue point in points)
        {
            scale.Append(new Color { Rgb = point.RgbColor });
        }

        return new ConditionalFormattingRule(scale)
        {
            Type = ConditionalFormatValues.ColorScale,
            Priority = priority
        };
    }
}

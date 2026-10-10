namespace Sheetwright;

/// <summary>
/// What a point on a colour scale is anchored to.
/// </summary>
public enum SpreadsheetColorScaleValueType
{
    /// <summary>The smallest value in the range.</summary>
    LowestValue,

    /// <summary>The largest value in the range.</summary>
    HighestValue,

    /// <summary>A fixed number.</summary>
    Number,

    /// <summary>A percentage of the way from the smallest value to the largest, 0 to 100.</summary>
    Percent,

    /// <summary>A percentile of the values in the range, 0 to 100.</summary>
    Percentile,

    /// <summary>The result of a formula.</summary>
    Formula
}

namespace Sheetwright;

/// <summary>
/// How a cell-value rule compares each cell with the value or values it was given.
/// </summary>
/// <remarks>
/// <see cref="Between"/> and <see cref="NotBetween"/> take two bounds and are inclusive, as they
/// are in Excel; every other operator takes one.
/// </remarks>
public enum SpreadsheetConditionalFormattingOperator
{
    Equal,
    NotEqual,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
    Between,
    NotBetween
}

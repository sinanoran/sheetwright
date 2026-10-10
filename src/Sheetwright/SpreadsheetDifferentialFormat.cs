namespace Sheetwright;

/// <summary>
/// The formatting a conditional format applies when its rule matches: a <c>dxf</c>.
/// </summary>
/// <remarks>
/// <para>
/// A differential format is a patch, like <see cref="SpreadsheetStylePatch"/>: only what was set
/// is written, and everything else stays as the cell's own style has it. That is why every
/// property is nullable, and why <c>false</c> is a value rather than an absence — a rule that sets
/// bold to <c>false</c> takes the bold off a heading that has it.
/// </para>
/// <para>
/// There is no font name or size. Excel ignores both in a conditional format and offers neither in
/// its own dialog, so a property for either would be one that silently did nothing.
/// </para>
/// </remarks>
internal sealed class SpreadsheetDifferentialFormat
{
    internal bool? Bold { get; set; }

    internal bool? Italic { get; set; }

    internal bool? Underline { get; set; }

    internal bool? Strikethrough { get; set; }

    internal string? FontColor { get; set; }

    internal string? FillColor { get; set; }

    internal string? NumberFormat { get; set; }

    internal SpreadsheetBorderStyle? LeftBorder { get; set; }

    internal string? LeftBorderColor { get; set; }

    internal SpreadsheetBorderStyle? RightBorder { get; set; }

    internal string? RightBorderColor { get; set; }

    internal SpreadsheetBorderStyle? TopBorder { get; set; }

    internal string? TopBorderColor { get; set; }

    internal SpreadsheetBorderStyle? BottomBorder { get; set; }

    internal string? BottomBorderColor { get; set; }

    internal bool HasFont => Bold is not null || Italic is not null || Underline is not null
        || Strikethrough is not null || FontColor is not null;

    internal bool HasBorder => LeftBorder is not null || RightBorder is not null
        || TopBorder is not null || BottomBorder is not null;

    internal bool IsEmpty => !HasFont && !HasBorder && FillColor is null && NumberFormat is null;
}

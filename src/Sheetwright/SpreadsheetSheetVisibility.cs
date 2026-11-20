namespace Sheetwright;

/// <summary>Whether a worksheet is shown in Excel's tab bar.</summary>
public enum SpreadsheetSheetVisibility
{
    /// <summary>Shown.</summary>
    Visible,

    /// <summary>Hidden by the person who made the workbook, and listed under Unhide.</summary>
    Hidden,

    /// <summary>Hidden so that Excel's own interface cannot unhide it.</summary>
    VeryHidden,
}

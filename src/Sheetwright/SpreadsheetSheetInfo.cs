namespace Sheetwright;

/// <summary>A worksheet of a workbook, as <see cref="SpreadsheetWorkbookReader.Sheets"/> lists it.</summary>
public sealed class SpreadsheetSheetInfo
{
    internal SpreadsheetSheetInfo(string name, int index, SpreadsheetSheetVisibility visibility, string partName)
    {
        Name = name;
        Index = index;
        Visibility = visibility;
        PartName = partName;
    }

    /// <summary>The sheet's name, as the tab shows it.</summary>
    public string Name { get; }

    /// <summary>The sheet's position among the worksheets, counted from 0, in tab order.</summary>
    public int Index { get; }

    /// <summary>Whether the sheet is hidden.</summary>
    public SpreadsheetSheetVisibility Visibility { get; }

    internal string PartName { get; }
}

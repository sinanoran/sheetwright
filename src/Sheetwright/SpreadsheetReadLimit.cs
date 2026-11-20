namespace Sheetwright;

/// <summary>
/// The limit of <see cref="SpreadsheetReaderOptions"/> a file went past. Each member has the name of the option
/// that sets it, so a message can tell a person which one to raise (or why the file is refused).
/// </summary>
public enum SpreadsheetReadLimit
{
    /// <summary><see cref="SpreadsheetReaderOptions.MaxUncompressedBytes"/>.</summary>
    MaxUncompressedBytes,

    /// <summary><see cref="SpreadsheetReaderOptions.MaxCompressionRatio"/>.</summary>
    MaxCompressionRatio,

    /// <summary><see cref="SpreadsheetReaderOptions.MaxSharedStrings"/>.</summary>
    MaxSharedStrings,

    /// <summary><see cref="SpreadsheetReaderOptions.MaxSharedStringCharacters"/>.</summary>
    MaxSharedStringCharacters,

    /// <summary><see cref="SpreadsheetReaderOptions.MaxParts"/>.</summary>
    MaxParts,

    /// <summary><see cref="SpreadsheetReaderOptions.MaxRows"/>.</summary>
    MaxRows,

    /// <summary><see cref="SpreadsheetReaderOptions.MaxColumns"/>.</summary>
    MaxColumns,

    /// <summary><see cref="SpreadsheetReaderOptions.MaxCellCharacters"/>.</summary>
    MaxCellCharacters,
}

namespace Sheetwright;

internal sealed class SpreadsheetSheetProtection
{
    internal SpreadsheetSheetProtection(bool lockEverything, string password)
    {
        LockEverything = lockEverything;
        Password = password;
    }

    internal bool LockEverything { get; }

    internal string Password { get; }
}

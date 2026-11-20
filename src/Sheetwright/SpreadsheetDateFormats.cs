namespace Sheetwright;

/// <summary>
/// Which cell formats make a number a date, and what date a serial number is.
/// </summary>
/// <remarks>
/// <para>
/// A date in a workbook is a number, and nothing in the cell says so: the cell's <em>format</em> does. A
/// reader that returned the number would hand a caller 46,011 where a person sees 20 November 2026, so the
/// style table is read and a number with a date format is returned as a date. The built-in formats (14 to 22,
/// 27 to 36, 45 to 47 and 50 to 58, which include the East Asian era formats) are known; a custom format is a
/// date format when, outside quotes, escapes and bracketed sections, it uses a letter of a date or a time: y, m,
/// d, h or s. That is Excel's own test, and it errs the right way: a format that is a number and says so is
/// never mistaken for a date because it holds a unit in quotes.
/// </para>
/// <para>
/// A serial is days since 31 December 1899, with Excel's famous mistake: it counts 29 February 1900, a day that
/// never existed, so serials below 60 are one day behind what .NET's OLE Automation date gives. A workbook
/// made on a Mac in the 1904 date system counts from 1 January 1904 instead, and says so in its workbook
/// properties.
/// </para>
/// </remarks>
internal static class SpreadsheetDateFormats
{
    // The latest serial that is a date: 31 December 9999.
    private const double LastSerial = 2958465.999999999;
    private const double Offset1904 = 1462;

    internal static bool IsBuiltInDate(int id)
    {
        return id is >= 14 and <= 22
            or >= 27 and <= 36
            or >= 45 and <= 47
            or >= 50 and <= 58;
    }

    internal static bool IsDateFormatCode(string code)
    {
        int index = 0;
        while (index < code.Length)
        {
            char character = code[index];

            switch (character)
            {
                case '"':
                    index = code.IndexOf('"', index + 1);
                    if (index < 0)
                    {
                        return false;
                    }

                    index++;
                    continue;

                case '\\':
                case '_':
                case '*':
                    index += 2;
                    continue;

                case '[':
                    int close = code.IndexOf(']', index + 1);
                    if (close < 0)
                    {
                        return false;
                    }

                    // [h], [mm] and [ss] are elapsed time, which is a time; every other bracket is a colour, a
                    // condition or a locale and says nothing about the number.
                    string inside = code.Substring(index + 1, close - index - 1);
                    if (inside.Length > 0 && inside.All(letter => char.ToLowerInvariant(letter) is 'h' or 'm' or 's'))
                    {
                        return true;
                    }

                    index = close + 1;
                    continue;

                case 'y' or 'Y' or 'm' or 'M' or 'd' or 'D' or 'h' or 'H' or 's' or 'S':
                    return true;

                default:
                    index++;
                    continue;
            }
        }

        return false;
    }

    /// <summary>The date and time a serial number stands for, when it stands for one.</summary>
    internal static bool TryConvert(decimal serial, bool date1904, out DateTime date)
    {
        date = default;
        double days = (double)serial;

        if (days < 0 || days > LastSerial)
        {
            return false;
        }

        if (date1904)
        {
            if (days + Offset1904 > LastSerial)
            {
                return false;
            }

            date = new DateTime(1904, 1, 1).AddMilliseconds(Math.Round(days * 86_400_000d));
            return true;
        }

        if (days >= 60 && days < 61)
        {
            // The day Excel invented. There is no date for it.
            return false;
        }

        date = DateTime.FromOADate(days is >= 1 and < 60 ? days + 1 : days);
        return true;
    }
}

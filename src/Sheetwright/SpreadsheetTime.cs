using System.Globalization;

namespace Sheetwright;

public sealed class SpreadsheetTime
{
    public int Hour { get; set; }

    public int Minute { get; set; }

    public int Second { get; set; }

    internal string ToSerialValue()
    {
        return new TimeSpan(Hour, Minute, Second).TotalDays.ToString(CultureInfo.InvariantCulture);
    }
}

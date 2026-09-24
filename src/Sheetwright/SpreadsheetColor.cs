using System.Drawing;

namespace Sheetwright;

public sealed class SpreadsheetColor
{
    private readonly Action<string> _setter;

    internal SpreadsheetColor(Action<string> setter)
    {
        _setter = setter;
    }

    public void SetColor(Color color)
    {
        _setter($"{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}");
    }

    public void SetColor(string rgb)
    {
        if (rgb is null)
        {
            throw new ArgumentNullException(nameof(rgb));
        }

        string normalized = SpreadsheetTextSanitizer.Sanitize(rgb).TrimStart('#');
        _setter(normalized.Length == 6 ? "FF" + normalized : normalized);
    }
}

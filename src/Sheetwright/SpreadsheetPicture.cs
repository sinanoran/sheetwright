namespace Sheetwright;

public sealed class SpreadsheetPicture
{
    private readonly byte[] _data;

    internal SpreadsheetPicture(string name, string title, byte[] data, string format, int row)
    {
        Name = name;
        Title = title;
        _data = data;
        Format = format;
        From = new SpreadsheetPicturePosition(row);
    }

    public string Name { get; }

    public string Title { get; }

    public string Format { get; }

    public SpreadsheetPicturePosition From { get; }

    public byte[] GetData()
    {
        return (byte[])_data.Clone();
    }
}

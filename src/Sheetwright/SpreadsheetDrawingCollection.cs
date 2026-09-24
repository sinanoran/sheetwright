using System.Collections;

namespace Sheetwright;

public sealed class SpreadsheetDrawingCollection : IEnumerable<SpreadsheetPicture>
{
    private readonly SpreadsheetWorksheet _worksheet;

    internal SpreadsheetDrawingCollection(SpreadsheetWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    public SpreadsheetPictureWriter AddPicture(string name, Stream image)
    {
        if (name is null)
        {
            throw new ArgumentNullException(nameof(name));
        }

        if (image is null)
        {
            throw new ArgumentNullException(nameof(image));
        }

        if (image.CanSeek)
        {
            return new SpreadsheetPictureWriter(_worksheet, SpreadsheetTextSanitizer.Sanitize(name), image);
        }

        // The image header has to be re-read to size the drawing, so buffer anything non-seekable.
        using MemoryStream buffered = new();
        image.CopyTo(buffered);
        buffered.Position = 0;
        return new SpreadsheetPictureWriter(_worksheet, SpreadsheetTextSanitizer.Sanitize(name), buffered);
    }

    public IEnumerator<SpreadsheetPicture> GetEnumerator()
    {
        return _worksheet.ReadPictures().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

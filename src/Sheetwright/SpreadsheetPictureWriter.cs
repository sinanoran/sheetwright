using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Sheetwright;

public sealed class SpreadsheetPictureWriter
{
    private readonly SpreadsheetWorksheet _worksheet;
    private readonly Xdr.OneCellAnchor _anchor;

    internal SpreadsheetPictureWriter(SpreadsheetWorksheet worksheet, string name, Stream image)
    {
        _worksheet = worksheet;
        _anchor = worksheet.AddPicture(name, image, 1, 1, 0, 0);
    }

    public void SetPosition(int pixelTop, int pixelLeft)
    {
        _worksheet.MovePicture(_anchor, 1, 1, pixelTop, pixelLeft);
    }
}

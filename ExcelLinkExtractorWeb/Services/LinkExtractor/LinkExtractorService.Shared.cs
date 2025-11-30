using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Caching.Memory;

namespace ExcelLinkExtractorWeb.Services.LinkExtractor;

public partial class LinkExtractorService
{
    private static Stylesheet GetStylesheet()
    {
        return CreateStylesheet();
    }

    private static Stylesheet CreateStylesheet()
    {
        var stylesheet = new Stylesheet();

        var fonts = new Fonts();
        fonts.Append(new Font());
        fonts.Append(new Font(new Bold()));
        fonts.Append(new Font(new Bold(), new Color { Rgb = new HexBinaryValue { Value = "FF0000FF" } }, new Underline()));
        fonts.Append(new Font(new Color { Rgb = new HexBinaryValue { Value = "FF888888" } }));
        fonts.Count = (uint)fonts.ChildElements.Count;

        var fills = new Fills();
        fills.Append(new Fill(new PatternFill { PatternType = PatternValues.None }));
        fills.Append(new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
        fills.Append(new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue { Value = "FFD9EAF7" } }) { PatternType = PatternValues.Solid }));
        fills.Append(new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue { Value = "FFD9F7E8" } }) { PatternType = PatternValues.Solid }));
        fills.Count = (uint)fills.ChildElements.Count;

        var borders = new Borders();
        borders.Append(new Border());
        borders.Count = (uint)borders.ChildElements.Count;

        var cellFormats = new CellFormats();
        cellFormats.Append(new CellFormat());
        cellFormats.Append(new CellFormat { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true });
        cellFormats.Append(new CellFormat { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true });
        cellFormats.Append(new CellFormat { FontId = 1, FillId = 2, BorderId = 0, ApplyFont = true, ApplyFill = true });
        cellFormats.Append(new CellFormat { FontId = 3, FillId = 0, BorderId = 0, ApplyFont = true });
        cellFormats.Append(new CellFormat { FontId = 1, FillId = 3, BorderId = 0, ApplyFont = true, ApplyFill = true });
        cellFormats.Count = (uint)cellFormats.ChildElements.Count;

        stylesheet.Fonts = fonts;
        stylesheet.Fills = fills;
        stylesheet.Borders = borders;
        stylesheet.CellFormats = cellFormats;

        return stylesheet;
    }

    private sealed class ProcessContext
    {
        public long InputBytes { get; init; }
        public int Rows { get; set; }
        public TimeSpan Duration { get; set; }
    }
}

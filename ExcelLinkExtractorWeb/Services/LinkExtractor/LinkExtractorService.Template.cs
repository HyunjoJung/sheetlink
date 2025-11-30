using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Caching.Memory;

namespace ExcelLinkExtractorWeb.Services.LinkExtractor;

public partial class LinkExtractorService
{
    public byte[] CreateTemplate()
    {
        if (_cache.TryGetValue("template:extract", out var cachedExtract) && cachedExtract is byte[] cachedTemplate)
        {
            return cachedTemplate;
        }

        var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Data"
            };
            sheets.Append(sheet);

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = GetStylesheet();

            var headerRow = new Row { RowIndex = 1 };
            headerRow.Append(new Cell
            {
                CellReference = "A1",
                DataType = CellValues.String,
                CellValue = new CellValue("Title"),
                StyleIndex = 3
            });
            headerRow.Append(new Cell
            {
                CellReference = "B1",
                DataType = CellValues.String,
                CellValue = new CellValue("URL"),
                StyleIndex = 3
            });
            sheetData.Append(headerRow);

            var row2 = new Row { RowIndex = 2 };
            row2.Append(new Cell
            {
                CellReference = "A2",
                DataType = CellValues.String,
                CellValue = new CellValue("Example Link 1"),
                StyleIndex = 2
            });
            sheetData.Append(row2);

            var row3 = new Row { RowIndex = 3 };
            row3.Append(new Cell
            {
                CellReference = "A3",
                DataType = CellValues.String,
                CellValue = new CellValue("Example Link 2"),
                StyleIndex = 2
            });
            sheetData.Append(row3);

            var hyperlinks = new Hyperlinks();
            hyperlinks.Append(new Hyperlink { Reference = "A2", Id = AddHyperlinkRelationship(worksheetPart, "https://www.example.com") });
            hyperlinks.Append(new Hyperlink { Reference = "A3", Id = AddHyperlinkRelationship(worksheetPart, "https://www.google.com") });
            worksheetPart.Worksheet.Append(hyperlinks);

            var row5 = new Row { RowIndex = 5 };
            row5.Append(new Cell
            {
                CellReference = "A5",
                DataType = CellValues.String,
                CellValue = new CellValue("Add hyperlinks to Title column. URLs will be extracted automatically."),
                StyleIndex = 4
            });
            sheetData.Append(row5);

            var columns = new Columns();
            columns.Append(new Column { Min = 1, Max = 1, Width = 30, CustomWidth = true });
            columns.Append(new Column { Min = 2, Max = 2, Width = 50, CustomWidth = true });
            worksheetPart.Worksheet.InsertBefore(columns, sheetData);

            workbookPart.Workbook.Save();
        }

        var bytes = stream.ToArray();
        _cache.Set("template:extract", bytes, new MemoryCacheEntryOptions
        {
            SlidingExpiration = TimeSpan.FromHours(2)
        });

        return bytes;
    }

    public byte[] CreateMergeTemplate()
    {
        if (_cache.TryGetValue("template:merge", out var cachedMerge) && cachedMerge is byte[] cachedTemplate)
        {
            return cachedTemplate;
        }

        var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Data"
            };
            sheets.Append(sheet);

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = GetStylesheet();

            var headerRow = new Row { RowIndex = 1 };
            headerRow.Append(new Cell
            {
                CellReference = "A1",
                DataType = CellValues.String,
                CellValue = new CellValue("Title"),
                StyleIndex = 5
            });
            headerRow.Append(new Cell
            {
                CellReference = "B1",
                DataType = CellValues.String,
                CellValue = new CellValue("URL"),
                StyleIndex = 5
            });
            sheetData.Append(headerRow);

            var samples = new[] {
                ("Google", "https://www.google.com"),
                ("GitHub", "https://github.com"),
                ("Stack Overflow", "https://stackoverflow.com")
            };

            uint currentRow = 2;
            foreach (var (title, url) in samples)
            {
                var row = new Row { RowIndex = currentRow };
                row.Append(new Cell
                {
                    CellReference = $"A{currentRow}",
                    DataType = CellValues.String,
                    CellValue = new CellValue(title)
                });
                row.Append(new Cell
                {
                    CellReference = $"B{currentRow}",
                    DataType = CellValues.String,
                    CellValue = new CellValue(url)
                });
                sheetData.Append(row);
                currentRow++;
            }

            var infoRow = new Row { RowIndex = currentRow + 1 };
            infoRow.Append(new Cell
            {
                CellReference = $"A{currentRow + 1}",
                DataType = CellValues.String,
                CellValue = new CellValue("Add your Title and URL values. URLs will be converted to hyperlinks."),
                StyleIndex = 4
            });
            sheetData.Append(infoRow);

            var columns = new Columns();
            columns.Append(new Column { Min = 1, Max = 1, Width = 30, CustomWidth = true });
            columns.Append(new Column { Min = 2, Max = 2, Width = 50, CustomWidth = true });
            worksheetPart.Worksheet.InsertBefore(columns, sheetData);

            workbookPart.Workbook.Save();
        }

        var bytes = stream.ToArray();
        _cache.Set("template:merge", bytes, new MemoryCacheEntryOptions
        {
            SlidingExpiration = TimeSpan.FromHours(2)
        });

        return bytes;
    }
}

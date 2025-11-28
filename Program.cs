using ClosedXML.Excel;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("===== 엑셀 하이퍼링크 추출 프로그램 =====\n");

        string inputPath = @"C:\Users\USER\Downloads\새 폴더\새 Microsoft Excel 워크시트.xlsx";
        string columnName = "제목";

        if (!File.Exists(inputPath))
        {
            Console.WriteLine("오류: 파일을 찾을 수 없습니다.");
            return;
        }

        try
        {
            ExtractHyperlinks(inputPath, columnName);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\n오류 발생: {ex.Message}");
            Console.WriteLine($"상세 정보: {ex.StackTrace}");
        }
    }

    static void ExtractHyperlinks(string inputPath, string columnName)
    {
        using var workbook = new XLWorkbook(inputPath);

        // 모든 시트 확인
        Console.WriteLine($"총 {workbook.Worksheets.Count}개의 시트:");
        foreach (var ws in workbook.Worksheets)
        {
            Console.WriteLine($"  - {ws.Name}");
        }
        Console.WriteLine();

        // 첫 번째 시트만 간단히 확인
        var worksheet = workbook.Worksheet(1);

        Console.WriteLine($"\n[시트: {worksheet.Name}] 헤더 행:");
        var headerRow = worksheet.Row(1);
        foreach (var cell in headerRow.CellsUsed())
        {
            Console.WriteLine($"  - {cell.Address.ColumnLetter}열: {cell.GetValue<string>()}");
        }
        Console.WriteLine();

        // 헤더 행 찾기 (columnName이 있는 행)
        int? headerRowNumber = null;
        int? targetColumnIndex = null;

        for (int i = 1; i <= 10; i++)
        {
            var row = worksheet.Row(i);
            foreach (var cell in row.CellsUsed())
            {
                if (cell.GetValue<string>().Equals(columnName, StringComparison.OrdinalIgnoreCase))
                {
                    headerRowNumber = i;
                    targetColumnIndex = cell.Address.ColumnNumber;
                    break;
                }
            }
            if (headerRowNumber != null) break;
        }

        if (targetColumnIndex == null)
        {
            Console.WriteLine($"\n오류: '{columnName}' 열을 찾을 수 없습니다.");
            return;
        }

        Console.WriteLine($"\n'{columnName}' 열을 {headerRowNumber}행, {targetColumnIndex}번 열에서 찾았습니다.");

        Console.WriteLine($"모든 데이터를 복사합니다 (하이퍼링크 유지)...\n");

        // 새 엑셀 파일 생성
        var newWorkbook = new XLWorkbook();
        var newWorksheet = newWorkbook.Worksheets.Add("추출된 링크");

        // 원본 헤더 행 복사
        var originalHeaderRow = worksheet.Row(headerRowNumber.Value);
        int colCount = originalHeaderRow.CellsUsed().Count();

        foreach (var cell in originalHeaderRow.CellsUsed())
        {
            var newCell = newWorksheet.Cell(1, cell.Address.ColumnNumber);
            newCell.Value = cell.Value;
            newCell.Style.Font.Bold = true;
        }

        int linkCount = 0;

        // 데이터 행 순회 (헤더 다음 행부터) - 모든 행 복사
        int dataStartRow = headerRowNumber.Value + 1;
        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;

        for (int i = dataStartRow; i <= lastRow; i++)
        {
            var row = worksheet.Row(i);
            int newRowIndex = i - dataStartRow + 2; // 헤더 다음부터 시작

            var titleCell = row.Cell(targetColumnIndex.Value);
            string extractedUrl = "";

            // 제목 셀에서 하이퍼링크 URL 추출
            if (titleCell.HasHyperlink)
            {
                var hyperlink = titleCell.GetHyperlink();
                extractedUrl = hyperlink.ExternalAddress?.ToString() ?? hyperlink.InternalAddress ?? "";
                linkCount++;
                Console.WriteLine($"[{linkCount}] {titleCell.GetValue<string>()} -> {extractedUrl}");
            }

            // 모든 열 복사
            foreach (var cell in row.CellsUsed())
            {
                var newCell = newWorksheet.Cell(newRowIndex, cell.Address.ColumnNumber);
                newCell.Value = cell.Value;

                // 하이퍼링크가 있으면 그대로 복사
                if (cell.HasHyperlink)
                {
                    var hyperlink = cell.GetHyperlink();
                    string url = hyperlink.ExternalAddress?.ToString() ?? hyperlink.InternalAddress ?? "";

                    newCell.SetHyperlink(new XLHyperlink(url));
                    newCell.Style.Font.FontColor = XLColor.Blue;
                    newCell.Style.Font.Underline = XLFontUnderlineValues.Single;
                }
            }

            // B열(2번 열)에 추출한 URL을 텍스트로 추가
            if (!string.IsNullOrEmpty(extractedUrl))
            {
                var linkTextCell = newWorksheet.Cell(newRowIndex, 2);
                linkTextCell.Value = extractedUrl;
            }
        }

        if (linkCount == 0)
        {
            Console.WriteLine("추출된 하이퍼링크가 없습니다.");
            return;
        }

        // 자동 열 너비 조정
        newWorksheet.Columns().AdjustToContents();

        // 저장
        string outputPath = Path.Combine(
            Path.GetDirectoryName(inputPath) ?? "",
            $"{Path.GetFileNameWithoutExtension(inputPath)}_추출된링크.xlsx"
        );

        newWorkbook.SaveAs(outputPath);

        Console.WriteLine($"\n총 {linkCount}개의 링크가 추출되었습니다.");
        Console.WriteLine($"저장 위치: {outputPath}");
    }
}

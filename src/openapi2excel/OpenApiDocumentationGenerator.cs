using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.OpenApi.Readers;
using openapi2excel.core.Builders;
using openapi2excel.core.Common;
using System.Text;

namespace openapi2excel.core;

public static class OpenApiDocumentationGenerator
{
    public static async Task GenerateDocumentation(string openApiFile, string outputFile,
       OpenApiDocumentationOptions options)
    {
        if (!File.Exists(openApiFile))
            throw new FileNotFoundException($"Invalid input file path: {openApiFile}.");

        if (string.IsNullOrEmpty(outputFile))
            throw new ArgumentNullException(outputFile, "Invalid output file path.");

        await using var fileStream = File.OpenRead(openApiFile);
        await GenerateDocumentationImpl(fileStream, outputFile, options);
    }

    public static async Task GenerateDocumentation(Stream openApiFileStream, string outputFile,
       OpenApiDocumentationOptions options)
    {
        if (string.IsNullOrEmpty(outputFile))
            throw new ArgumentNullException(outputFile, "Invalid output file path.");

        await GenerateDocumentationImpl(openApiFileStream, outputFile, options);
    }

    private static async Task GenerateDocumentationImpl(Stream openApiFileStream, string outputFile,
       OpenApiDocumentationOptions options)
    {
        var readResult = await new OpenApiStreamReader().ReadAsync(openApiFileStream);
        AssertReadResult(readResult);

        using var workbook = new XLWorkbook();
        var infoWorksheetsBuilder = new InfoWorksheetBuilder(workbook, options);
        infoWorksheetsBuilder.Build(readResult.OpenApiDocument);

        var worksheetBuilder = new OperationWorksheetBuilder(workbook, options);
        readResult.OpenApiDocument.Paths.ForEach(path
           => path.Value.Operations.ForEach(operation
                 =>
                 {
                     var worksheet = worksheetBuilder.Build(path.Key, path.Value, operation.Key, operation.Value);
                     infoWorksheetsBuilder.AddLink(operation.Key, path.Key, worksheet);
                 }
           ));
       
        foreach (var worksheet in workbook.Worksheets)
        {
            // 워크시트 기본 스타일 설정
            worksheet.Style.Font.FontName = "맑은 고딕";
            worksheet.Style.Font.FontSize = 10;
            worksheet.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Top);
            
            var usedRange = worksheet.RangeUsed();
            if (usedRange != null)
            {
                // 컬럼 너비 자동 조정
                foreach (var rangeColumn in usedRange.Columns())
                {
                    var col = worksheet.Column(rangeColumn.ColumnNumber());
                    col.AdjustToContents();
                    
                    // 컬럼별 최소/최대 너비 설정
                    double minWidth = 12;  // 최소 너비
                    double maxWidth = 80;  // 최대 너비 (설명 컬럼 등을 위해)
                    
                    if (col.Width < minWidth)
                        col.Width = minWidth;
                    else if (col.Width > maxWidth)
                        col.Width = maxWidth;
                }
                
                // 행 높이는 자동 조정하지 않음 (성능상 이유)
                // foreach (var row in usedRange.Rows())
                // {
                //     row.AdjustToContents();
                // }
                
                // 워크시트에 보기 좋은 격자선 설정
                worksheet.ShowGridLines = true;
                
                // 첫 번째 행 고정 (헤더가 있는 경우)
                if (usedRange.RowCount() > 1)
                {
                    // "작업 요약" 행까지 찾기
                    var freezeRowIndex = 1; // 기본값은 첫 번째 행
                    for (int i = 1; i <= usedRange.RowCount(); i++)
                    {
                        var firstCell = worksheet.Cell(i, 1);
                        var cellValue = firstCell.Value.ToString();
                        
                        // "작업 요약" 텍스트가 포함된 행을 찾으면 그 행까지 고정
                        if (cellValue.Contains("작업 요약"))
                        {
                            freezeRowIndex = i;
                            break;
                        }
                        
                        // "작업 요약"이 없는 경우를 위한 기본 헤더 찾기 (기존 로직 유지)
                        if (firstCell.Style.Font.Bold && 
                            (firstCell.Style.Fill.BackgroundColor == XLColor.FromArgb(68, 114, 196) ||
                             firstCell.Value.ToString().Contains("파라미터") ||
                             firstCell.Value.ToString().Contains("Type")))
                        {
                            if (!cellValue.Contains("작업 요약"))
                            {
                                freezeRowIndex = i;
                            }
                        }
                    }
                    
                    if (freezeRowIndex <= usedRange.RowCount())
                    {
                        worksheet.SheetView.FreezeRows(freezeRowIndex);
                    }
                }
                
                // 인쇄 설정 개선
                worksheet.PageSetup.PageOrientation = XLPageOrientation.Portrait;
                worksheet.PageSetup.Margins.Top = 0.75;
                worksheet.PageSetup.Margins.Bottom = 0.75;
                worksheet.PageSetup.Margins.Left = 0.7;
                worksheet.PageSetup.Margins.Right = 0.7;
                
                // 페이지 번호 추가
                worksheet.PageSetup.Header.Right.AddText("페이지 ");
                worksheet.PageSetup.Header.Right.AddText(XLHFPredefinedText.PageNumber);
                worksheet.PageSetup.Header.Right.AddText(" / ");
                worksheet.PageSetup.Header.Right.AddText(XLHFPredefinedText.NumberOfPages);
            }
        }

        workbook.SaveAs(new FileInfo(outputFile).FullName);
    }

    private static void AssertReadResult(ReadResult readResult)
    {
        if (!readResult.OpenApiDiagnostic.Errors.Any())
            return;

        var errorMessageBuilder = new StringBuilder();
        errorMessageBuilder.AppendLine("Some errors occurred while processing input file.");
        readResult.OpenApiDiagnostic.Errors.ToList().ForEach(e => errorMessageBuilder.AppendLine($"{e.Message} ({e.Pointer})"));
        throw new InvalidOperationException(errorMessageBuilder.ToString());
    }
}
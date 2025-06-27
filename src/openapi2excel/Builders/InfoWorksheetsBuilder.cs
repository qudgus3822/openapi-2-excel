using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders;

internal class InfoWorksheetBuilder(IXLWorkbook workbook, OpenApiDocumentationOptions options)
   : WorksheetBuilder(options)
{
   private OpenApiDocument _readResultOpenApiDocument = null!;
   private IXLWorksheet _worksheet = null!;
   public static string Name => OpenApiDocumentationLanguageConst.Info;
   private int _actualRowIndex = 1;

   public IXLWorksheet Build(OpenApiDocument openApiDocument)
   {
      _readResultOpenApiDocument = openApiDocument;
      _worksheet = workbook.Worksheets.Add(Name);
      
      // 워크시트 기본 스타일 설정
      _worksheet.Style.Font.FontName = "맑은 고딕";
      _worksheet.Style.Font.FontSize = 10;
      
      // 컬럼 너비 설정
      _worksheet.Column(1).Width = 20;
      _worksheet.Column(2).Width = 60;
      _worksheet.Column(3).Width = 20;

      // 대형 제목 영역 (표지 스타일)
      var titleRange = _worksheet.Range("A1:C3");
      titleRange.Merge();
      var titleCell = titleRange.FirstCell();
      titleCell.SetValue($"📋 {_readResultOpenApiDocument.Info.Title ?? "API 명세서"}");
      titleCell.Style.Font.SetBold(true);
      titleCell.Style.Font.SetFontSize(24);
      titleCell.Style.Font.SetFontColor(XLColor.White);
      titleCell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(68, 114, 196));
      titleCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
      titleCell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
      titleCell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
      _actualRowIndex = 4;

      // 부제목 (버전 정보)
      if (!string.IsNullOrEmpty(_readResultOpenApiDocument.Info.Version))
      {
         _actualRowIndex++;
         var versionRange = _worksheet.Range($"A{_actualRowIndex}:C{_actualRowIndex}");
         versionRange.Merge();
         var versionCell = versionRange.FirstCell();
         versionCell.SetValue($"Version {_readResultOpenApiDocument.Info.Version}");
         versionCell.Style.Font.SetBold(true);
         versionCell.Style.Font.SetFontSize(16);
         versionCell.Style.Font.SetFontColor(XLColor.FromArgb(68, 114, 196));
         versionCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
         _actualRowIndex += 2;
      }

      // 설명 영역
      if (!string.IsNullOrEmpty(_readResultOpenApiDocument.Info.Description))
      {
         var descRange = _worksheet.Range($"A{_actualRowIndex}:C{_actualRowIndex + 2}");
         descRange.Merge();
         var descCell = descRange.FirstCell();
         descCell.SetValue(_readResultOpenApiDocument.Info.Description);
         descCell.Style.Font.SetFontSize(12);
         descCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
         descCell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
         descCell.Style.Alignment.SetWrapText(true);
         descCell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(248, 248, 248));
         descCell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
         _actualRowIndex += 4;
      }

      // 생성 정보
      _actualRowIndex++;
      var dateRange = _worksheet.Range($"A{_actualRowIndex}:C{_actualRowIndex}");
      dateRange.Merge();
      var dateCell = dateRange.FirstCell();
      dateCell.SetValue($"📅 생성일: {DateTime.Now:yyyy년 MM월 dd일 HH:mm}");
      dateCell.Style.Font.SetFontSize(11);
      dateCell.Style.Font.SetItalic(true);
      dateCell.Style.Font.SetFontColor(XLColor.Gray);
      dateCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
      _actualRowIndex += 3;

      // API 엔드포인트 목록 제목 추가
      var endpointsTitleCell = _worksheet.Cell(_actualRowIndex, 1);
      var endpointsTitleRange = _worksheet.Range($"A{_actualRowIndex}:C{_actualRowIndex}");
      endpointsTitleRange.Merge();
      endpointsTitleCell.SetValue("🔗 API 엔드포인트 목록");
      endpointsTitleCell.Style.Font.SetBold(true);
      endpointsTitleCell.Style.Font.SetFontSize(16);
      endpointsTitleCell.Style.Font.SetFontColor(XLColor.FromArgb(68, 114, 196));
      endpointsTitleCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
      endpointsTitleCell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(217, 237, 247));
      endpointsTitleCell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
      _actualRowIndex += 2;
      
      // 헤더 행 추가
      var methodHeaderCell = _worksheet.Cell(_actualRowIndex, 1);
      var pathHeaderCell = _worksheet.Cell(_actualRowIndex, 2);
      var summaryHeaderCell = _worksheet.Cell(_actualRowIndex, 3);
      
      methodHeaderCell.SetTextHeader("METHOD");
      pathHeaderCell.SetTextHeader("PATH");
      summaryHeaderCell.SetTextHeader("설명");
      
      // 헤더 스타일 강화
      var headerRange = _worksheet.Range(_actualRowIndex, 1, _actualRowIndex, 3);
      foreach (var cell in headerRange.Cells())
      {
         cell.Style.Font.SetBold(true);
         cell.Style.Font.SetFontColor(XLColor.White);
         cell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(68, 114, 196));
         cell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
         cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
      }
      
      _actualRowIndex++;

      return _worksheet;
   }

   public void AddLink(OperationType operation, string path, IXLWorksheet worksheet)
   {
      var methodCell = _worksheet.Cell(_actualRowIndex, 1);
      var pathCell = _worksheet.Cell(_actualRowIndex, 2);
      var summaryCell = _worksheet.Cell(_actualRowIndex, 3);
      
      // HTTP 메서드 셀 스타일링
      methodCell.SetValue(operation.ToString().ToUpper());
      methodCell.SetHyperlink(new XLHyperlink($"'{worksheet.Name}'!A1"));
      methodCell.SetMethodStyle(operation.ToString());
      
      // 경로 셀 스타일링
      pathCell.SetValue(path);
      pathCell.SetHyperlink(new XLHyperlink($"'{worksheet.Name}'!A1"));
      pathCell.SetDataStyle();
      pathCell.Style.Font.SetFontName("Consolas");
      pathCell.Style.Font.SetFontColor(XLColor.FromArgb(68, 114, 196));
      pathCell.Style.Font.SetUnderline(XLFontUnderlineValues.Single);

      // 설명 셀 추가 (operation의 summary나 description 사용)
      var operationInfo = _readResultOpenApiDocument.Paths[path].Operations[operation];
      var summary = !string.IsNullOrEmpty(operationInfo.Summary) ? operationInfo.Summary : 
                   !string.IsNullOrEmpty(operationInfo.Description) ? operationInfo.Description : "-";
      
      summaryCell.SetValue(summary);
      summaryCell.SetDataStyle();
      summaryCell.Style.Alignment.SetWrapText(true);
      if (summary != "-")
      {
         summaryCell.SetHyperlink(new XLHyperlink($"'{worksheet.Name}'!A1"));
         summaryCell.Style.Font.SetFontColor(XLColor.FromArgb(68, 114, 196));
      }

      // 전체 행에 교대로 배경색 적용
      if (_actualRowIndex % 2 == 0)
      {
         var rowRange = _worksheet.Range(_actualRowIndex, 1, _actualRowIndex, 3);
         foreach (var cell in rowRange.Cells())
         {
            if (cell.Style.Fill.BackgroundColor == XLColor.NoColor)
            {
               cell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(248, 248, 248));
            }
         }
      }

      _actualRowIndex++;
   }
}
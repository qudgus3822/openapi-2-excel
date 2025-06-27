using ClosedXML.Excel;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class HomePageLinkBuilder(RowPointer actualRow, IXLWorksheet worksheet, OpenApiDocumentationOptions options)
   : WorksheetPartBuilder(actualRow, worksheet, options)
{
   public void AddHomePageLinkSection()
   {
      var linkCell = Cell(1);
      linkCell.SetValue("🏠 목차로").SetHyperlink(new XLHyperlink($"'{InfoWorksheetBuilder.Name}'!A1"));
      
      // 홈 링크 스타일 적용
      linkCell.Style.Font.SetBold(true);
      linkCell.Style.Font.SetFontColor(XLColor.FromArgb(68, 114, 196)); // 파란색
      linkCell.Style.Font.SetUnderline(XLFontUnderlineValues.Single);
      linkCell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(245, 245, 245)); // 연한 회색 배경
      linkCell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
      linkCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
      linkCell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
      
      ActualRow.MoveNext(2);
   }
}
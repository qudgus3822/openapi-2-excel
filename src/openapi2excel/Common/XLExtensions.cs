using ClosedXML.Excel;

namespace openapi2excel.core.Common;

// ReSharper disable once InconsistentNaming
internal static class XLExtensions
{
   public static IXLCell SetBoldStyle(this IXLCell cell)
   {
      cell.Style.Font.SetBold(true);
      return cell;
   }

   public static IXLCell SetText(this IXLCell cell, string? value)
      => cell.SetValue(value?.Trim());

   public static IXLCell SetTextBold(this IXLCell cell, string? value)
      => cell.SetBoldStyle().SetValue(value?.Trim());

   public static IXLCell NextRow(this IXLCell cell)
      => cell.Worksheet.Cell(cell.Address.RowNumber + 1, 1);

   public static int GetColumnNumber(this IXLCell cell)
      => cell.Address.ColumnNumber;

   public static IXLCell If(this IXLCell cell, bool condition, Func<IXLCell, IXLCell> func)
      => condition ? func(cell) : cell;

   public static IXLCell IfNotEmpty(this IXLCell cell, string text, Func<IXLCell, IXLCell> func)
      => string.IsNullOrEmpty(text) ? cell : func(cell);

   public static IXLCell SetHorizontalAlignment(this IXLCell cell, XLAlignmentHorizontalValues alignment)
   {
      cell.Style.Alignment.SetHorizontal(alignment);
      return cell;
   }

   public static IXLCell SetVerticalAlignment(this IXLCell cell, XLAlignmentVerticalValues alignment)
   {
      cell.Style.Alignment.SetVertical(alignment);
      return cell;
   }

   public static IXLCell SetBackground(this IXLCell cell, XLColor color)
   {
      cell.Style.Fill.SetBackgroundColor(color);
      return cell;
   }

   public static IXLCell SetBackground(this IXLCell cell, int toColumn, XLColor color)
   {
      var tmpCell = cell;
      while (tmpCell.Address.ColumnNumber <= toColumn)
         tmpCell = tmpCell.SetBackground(color).CellRight();

      return cell;
   }

   public static IXLCell SetBottomBorder(this IXLCell cell)
   {
      cell.Style.Border.SetBottomBorder(XLBorderStyleValues.Medium);
      return cell;
   }

   public static IXLCell SetBottomBorder(this IXLCell cell, int toColumn)
   {
      var tmpCell = cell;
      while (tmpCell.Address.ColumnNumber <= toColumn)
         tmpCell = tmpCell.SetBottomBorder().CellRight();

      return cell;
   }

   /// <summary>
   /// 헤더 스타일 적용 (진한 파란색 배경, 흰색 글씨, 볼드)
   /// </summary>
   public static IXLCell SetHeaderStyle(this IXLCell cell)
   {
      cell.Style.Font.SetBold(true);
      cell.Style.Font.SetFontColor(XLColor.White);
      cell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(68, 114, 196)); // 진한 파란색
      cell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
      cell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
      cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
      return cell;
   }

   /// <summary>
   /// 서브헤더 스타일 적용 (연한 파란색 배경, 검은색 글씨, 볼드)
   /// </summary>
   public static IXLCell SetSubHeaderStyle(this IXLCell cell)
   {
      cell.Style.Font.SetBold(true);
      cell.Style.Font.SetFontColor(XLColor.Black);
      cell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(180, 198, 231)); // 연한 파란색
      cell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
      cell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
      cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
      return cell;
   }

   /// <summary>
   /// 데이터 셀 스타일 적용 (테두리, 정렬)
   /// </summary>
   public static IXLCell SetDataStyle(this IXLCell cell)
   {
      cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
      cell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Top);
      cell.Style.Alignment.SetWrapText(true);
      return cell;
   }

   /// <summary>
   /// 필수 필드 스타일 적용 (빨간색 텍스트)
   /// </summary>
   public static IXLCell SetRequiredStyle(this IXLCell cell)
   {
      cell.Style.Font.SetFontColor(XLColor.Red);
      cell.Style.Font.SetBold(true);
      return cell;
   }

   /// <summary>
   /// 선택적 필드 스타일 적용 (회색 텍스트)
   /// </summary>
   public static IXLCell SetOptionalStyle(this IXLCell cell)
   {
      cell.Style.Font.SetFontColor(XLColor.Gray);
      return cell;
   }

   /// <summary>
   /// HTTP 메서드 스타일 적용 (메서드에 따른 색상)
   /// </summary>
   public static IXLCell SetMethodStyle(this IXLCell cell, string method)
   {
      var color = method.ToUpper() switch
      {
         "GET" => XLColor.FromArgb(92, 184, 92),    // 초록색
         "POST" => XLColor.FromArgb(91, 192, 222),   // 하늘색
         "PUT" => XLColor.FromArgb(240, 173, 78),    // 주황색
         "DELETE" => XLColor.FromArgb(217, 83, 79),  // 빨간색
         "PATCH" => XLColor.FromArgb(138, 109, 59),  // 갈색
         _ => XLColor.FromArgb(119, 119, 119)        // 회색
      };
      
      cell.Style.Font.SetBold(true);
      cell.Style.Font.SetFontColor(XLColor.White);
      cell.Style.Fill.SetBackgroundColor(color);
      cell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
      cell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
      cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
      return cell;
   }

   /// <summary>
   /// 성공 응답 스타일 (초록색 배경)
   /// </summary>
   public static IXLCell SetSuccessResponseStyle(this IXLCell cell)
   {
      cell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(217, 237, 247)); // 연한 초록색
      cell.Style.Font.SetBold(true);
      cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
      return cell;
   }

   /// <summary>
   /// 오류 응답 스타일 (빨간색 배경)
   /// </summary>
   public static IXLCell SetErrorResponseStyle(this IXLCell cell)
   {
      cell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(242, 222, 222)); // 연한 빨간색
      cell.Style.Font.SetBold(true);
      cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
      return cell;
   }

   /// <summary>
   /// 제목 스타일 적용 (큰 글씨, 볼드)
   /// </summary>
   public static IXLCell SetTitleStyle(this IXLCell cell)
   {
      cell.Style.Font.SetBold(true);
      cell.Style.Font.SetFontSize(14);
      cell.Style.Font.SetFontColor(XLColor.FromArgb(68, 114, 196));
      cell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
      return cell;
   }

   /// <summary>
   /// 범위에 테이블 스타일 적용
   /// </summary>
   public static IXLRange SetTableStyle(this IXLRange range)
   {
      var table = range.CreateTable();
      table.Theme = XLTableTheme.TableStyleMedium9;
      return range;
   }

   /// <summary>
   /// 범위에 얼터네이팅 행 색상 적용
   /// </summary>
   public static IXLRange SetAlternatingRowColors(this IXLRange range)
   {
      for (int i = 1; i <= range.RowCount(); i++)
      {
         if (i % 2 == 0)
         {
            range.Row(i).Style.Fill.SetBackgroundColor(XLColor.FromArgb(248, 248, 248)); // 연한 회색
         }
      }
      return range;
   }

   /// <summary>
   /// 셀에 텍스트와 함께 헤더 스타일 적용
   /// </summary>
   public static IXLCell SetTextHeader(this IXLCell cell, string? value)
      => cell.SetHeaderStyle().SetValue(value?.Trim());

   /// <summary>
   /// 셀에 텍스트와 함께 서브헤더 스타일 적용
   /// </summary>
   public static IXLCell SetTextSubHeader(this IXLCell cell, string? value)
      => cell.SetSubHeaderStyle().SetValue(value?.Trim());

   /// <summary>
   /// 셀에 텍스트와 함께 데이터 스타일 적용
   /// </summary>
   public static IXLCell SetTextData(this IXLCell cell, string? value)
      => cell.SetDataStyle().SetValue(value?.Trim());

   /// <summary>
   /// 셀에 텍스트와 함께 제목 스타일 적용
   /// </summary>
   public static IXLCell SetTextTitle(this IXLCell cell, string? value)
      => cell.SetTitleStyle().SetValue(value?.Trim());
}
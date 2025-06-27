using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders.Common;

internal class OpenApiSchemaDescriptor(IXLWorksheet worksheet, OpenApiDocumentationOptions options)
{
   public int AddNameHeader(RowPointer actualRow, int startColumn)
      => worksheet.Cell(actualRow, startColumn).SetTextHeader("파라미터명").GetColumnNumber();

   public int AddNameValue(string name, int actualRow, int startColumn)
      => worksheet.Cell(actualRow, startColumn).SetTextData(name).GetColumnNumber();

   public int AddSchemaDescriptionHeader(RowPointer actualRow, int startColumn)
   {
      var cell = worksheet.Cell(actualRow, startColumn).SetTextHeader("Type")
         //.CellRight().SetTextHeader("Type detail")
         .CellRight().SetTextHeader("Format")
         //.CellRight().SetTextHeader("Length")
         .CellRight().SetTextHeader("필수여부")
         //.CellRight().SetTextHeader("Nullable")
         //.CellRight().SetTextHeader("Range")
         //.CellRight().SetTextHeader("Pattern")
         //.CellRight().SetTextHeader("Enum")
         //.CellRight().SetTextHeader("Deprecated")
         //.CellRight().SetTextHeader("Default")
         //.CellRight().SetTextHeader("Example")
         .CellRight().SetTextHeader("설명");

      return cell.GetColumnNumber();
   }

   public int AddSchemaDescriptionValues(OpenApiSchema schema, bool required, RowPointer actualRow, int startColumn, string? description = null, bool includeArrayItemType = false )
   {
      if (schema.Items != null && includeArrayItemType)
      {
         var typeCell = worksheet.Cell(actualRow, startColumn);
         typeCell.SetTextData($"array({schema.Items.Type})");
         // 배열 타입은 특별한 색상으로 강조
         typeCell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(230, 247, 255));
         typeCell.Style.Font.SetFontName("Consolas");

         var formatCell = typeCell.CellRight();
         //.CellRight().SetTextData(schema.GetObjectDescription())
         //.CellRight().SetTextData(schema.Items.Type)
         //.CellRight().SetTextData(schema.Items.Format)
         //.CellRight().SetTextData(schema.GetPropertyLengthDescription()).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
         formatCell.SetTextData(schema.Items.Format);

         var requiredCell = formatCell.CellRight();
         requiredCell.SetTextData(options.Language.Get(required));
         requiredCell.SetHorizontalAlignment(XLAlignmentHorizontalValues.Center);
         //.CellRight().SetTextData(options.Language.Get(schema.Nullable)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
         //.CellRight().SetTextData(schema.GetPropertyRangeDescription()).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
         //.CellRight().SetTextData(schema.Items.Pattern)
         //.CellRight().SetTextData(schema.Items.GetEnumDescription())
         //.CellRight().SetTextData(options.Language.Get(schema.Deprecated)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
         //.CellRight().SetTextData(schema.GetExampleDescription())
         
         // 필수/선택 여부에 따른 스타일 적용
         if (required)
         {
            requiredCell.SetRequiredStyle();
         }
         else
         {
            requiredCell.SetOptionalStyle();
         }

         var descCell = requiredCell.CellRight();
         var descText = string.IsNullOrEmpty(schema.Description) ? description : schema.Description;
         descCell.SetTextData(descText?.StripHtmlTags());
         descCell.Style.Alignment.SetWrapText(true);

         // 모든 셀에 데이터 스타일 적용
         typeCell.SetDataStyle();
         formatCell.SetDataStyle();
         requiredCell.SetDataStyle();
         descCell.SetDataStyle();

         return descCell.GetColumnNumber();
      }
      else
      {
         var typeCell = worksheet.Cell(actualRow, startColumn);
         var typeDescription = schema.GetTypeDescription();
         typeCell.SetTextData(typeDescription);
         
         // 타입별 색상 적용
         var typeColor = typeDescription?.ToLower() switch
         {
            "string" => XLColor.FromArgb(255, 248, 220),    // 연한 베이지
            "integer" or "number" => XLColor.FromArgb(230, 255, 230),  // 연한 초록
            "boolean" => XLColor.FromArgb(255, 230, 230),   // 연한 빨강
            "object" => XLColor.FromArgb(240, 240, 255),    // 연한 파랑
            _ => XLColor.White
         };
         typeCell.Style.Fill.SetBackgroundColor(typeColor);
         typeCell.Style.Font.SetFontName("Consolas");

         var formatCell = typeCell.CellRight();
         //.CellRight().SetTextData(schema.GetObjectDescription())
         formatCell.SetTextData(schema.Format);
         //.CellRight().SetTextData(schema.GetPropertyLengthDescription()).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)

         var requiredCell = formatCell.CellRight();
         requiredCell.SetTextData(options.Language.Get(required));
         requiredCell.SetHorizontalAlignment(XLAlignmentHorizontalValues.Center);
         //.CellRight().SetTextData(options.Language.Get(schema.Nullable)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
         //.CellRight().SetTextData(schema.GetPropertyRangeDescription()).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
         //.CellRight().SetTextData(schema.Pattern)
         //.CellRight().SetTextData(schema.GetEnumDescription())
         //.CellRight().SetTextData(options.Language.Get(schema.Deprecated)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
         //.CellRight().SetTextData(schema.GetDefaultDescription())
         //.CellRight().SetTextData(schema.GetExampleDescription())
         
         // 필수/선택 여부에 따른 스타일 적용
         if (required)
         {
            requiredCell.SetRequiredStyle();
         }
         else
         {
            requiredCell.SetOptionalStyle();
         }

         var descCell = requiredCell.CellRight();
         var descText = string.IsNullOrEmpty(schema.Description) ? description : schema.Description;
         descCell.SetTextData(descText?.StripHtmlTags());
         descCell.Style.Alignment.SetWrapText(true);

         // 모든 셀에 데이터 스타일 적용
         typeCell.SetDataStyle();
         formatCell.SetDataStyle();
         requiredCell.SetDataStyle();
         descCell.SetDataStyle();

         return descCell.GetColumnNumber();
      }
   }

   public int AddSchemaDescription(RowPointer actualRow, int startColumn, OpenApiSchema schema, bool isRequired = false)
   {
      var type = schema.GetTypeDescription();

      var cell = worksheet.Cell(actualRow, startColumn).SetTextData(type)
         //.CellRight().SetTextData(GetTypeDetails(schema))
         .CellRight().SetTextData(schema.Format ?? "")
         //.CellRight().SetTextData(GetLength(schema))
         .CellRight().SetTextData(isRequired ? "✓" : "")
         //.CellRight().SetTextData(schema.Nullable ? "Yes" : "No")
         //.CellRight().SetTextData(GetRange(schema))
         //.CellRight().SetTextData(schema.Pattern ?? "")
         //.CellRight().SetTextData(GetEnum(schema))
         //.CellRight().SetTextData(schema.Deprecated ? "Yes" : "No")
         //.CellRight().SetTextData(GetDefault(schema))
         //.CellRight().SetTextData(GetExample(schema))
         .CellRight().SetTextData(schema.Description ?? "");

      return cell.GetColumnNumber();
   }
}
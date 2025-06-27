using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class RequestParametersBuilder(
   RowPointer actualRow,
   int attributesColumnIndex,
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options)
   : WorksheetPartBuilder(actualRow, worksheet, options)
{
   private readonly OpenApiSchemaDescriptor _schemaDescriptor = new(worksheet, options);

   public void AddRequestParametersPart(OpenApiOperation operation)
   {
      attributesColumnIndex = attributesColumnIndex > 1 ? attributesColumnIndex : 2;
      if (!operation.Parameters.Any())
         return;

      Cell(1).SetTextTitle("파라미터");
      ActualRow.MoveNext();
      
      using (var _ = new Section(Worksheet, ActualRow))
      {
         var nameHeaderCell = Cell(1).SetTextHeader("파라미터명");
         var locationHeaderCell = nameHeaderCell.CellRight().SetTextHeader("위치");
         var nextCell = locationHeaderCell.CellRight();

         var lastUsedColumn = _schemaDescriptor.AddSchemaDescriptionHeader(ActualRow, nextCell.Address.ColumnNumber);

         var headerRange = Worksheet.Range(ActualRow.Get(), 1, ActualRow.Get(), lastUsedColumn);
         foreach (var cell in headerRange.Cells())
         {
            if (string.IsNullOrEmpty(cell.Value.ToString()))
            {
               cell.SetHeaderStyle();
            }
         }

         ActualRow.MoveNext();
         
         foreach (var operationParameter in operation.Parameters)
         {
            AddPropertyRow(operationParameter);
         }
         ActualRow.MovePrev();
      }

      ActualRow.MoveNext(2);
   }

   private void AddPropertyRow(OpenApiParameter parameter)
   {
      var nameCell = Cell(1).SetTextData(parameter.Name);
      
      var locationCell = nameCell.CellRight();
      locationCell.SetTextData(parameter.In.ToString()?.ToUpper());
      
      var locationColor = parameter.In.ToString()?.ToUpper() switch
      {
         "PATH" => XLColor.FromArgb(255, 235, 156),
         "QUERY" => XLColor.FromArgb(201, 242, 155),
         "HEADER" => XLColor.FromArgb(174, 203, 250),
         "COOKIE" => XLColor.FromArgb(255, 204, 204),
         _ => XLColor.White
      };
      locationCell.Style.Fill.SetBackgroundColor(locationColor);
      locationCell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
      locationCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
      locationCell.Style.Font.SetBold(true);

      var nextCell = locationCell.CellRight();

      if (parameter.Schema != null)
      {
         _schemaDescriptor.AddSchemaDescriptionValues(parameter.Schema, parameter.Required, ActualRow, nextCell.Address.ColumnNumber, parameter.Description, true);
      }
      
      var lastColumn = Worksheet.LastColumnUsed();
      if (lastColumn != null)
      {
         var rowRange = Worksheet.Range(ActualRow.Get(), 1, ActualRow.Get(), lastColumn.ColumnNumber());
         foreach (var cell in rowRange.Cells())
         {
            if (cell.Address.ColumnNumber != locationCell.Address.ColumnNumber)
            {
               cell.SetDataStyle();
            }
         }
      }

      ActualRow.MoveNext();
   }
}
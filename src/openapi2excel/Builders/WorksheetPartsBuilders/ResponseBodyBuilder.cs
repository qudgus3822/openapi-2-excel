using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class ResponseBodyBuilder(
   RowPointer actualRow,
   int attributesColumnIndex,
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options) : WorksheetPartBuilder(actualRow, worksheet, options)
{
   public void AddResponseBodyPart(OpenApiOperation operation)
   {
      if (!operation.Responses.Any())
         return;

      Cell(1).SetTextTitle("응답");
      ActualRow.MoveNext();
      
      using (var _ = new Section(Worksheet, ActualRow))
      {
         var builder = new PropertiesTreeBuilder(attributesColumnIndex, Worksheet, Options);
         foreach (var response in operation.Responses)
         {
            AddResponseHttpCode(response.Key, response.Value.Description);
            AddReponseHeaders(response.Value.Headers);
            builder.AddPropertiesTreeForMediaTypes(ActualRow, response.Value.Content, Options);
         }
      }
      ActualRow.MoveNext();
   }

   private void AddReponseHeaders(IDictionary<string, OpenApiHeader> valueHeaders)
   {
      if (!valueHeaders.Any())
         return;

      ActualRow.MoveNext();

      var responseHeadertRowPointer = ActualRow.Copy();
      Cell(1).SetTextSubHeader("Response Headers");
      ActualRow.MoveNext();

      using (var _ = new Section(Worksheet, ActualRow))
      {
         var schemaDescriptor = new OpenApiSchemaDescriptor(Worksheet, Options);

         InsertHeader(schemaDescriptor);
         ActualRow.MoveNext();

         foreach (var openApiHeader in valueHeaders)
         {
            InsertProperty(openApiHeader, schemaDescriptor);
            ActualRow.MoveNext();
         }
      }
      ActualRow.MoveNext();

      void InsertHeader(OpenApiSchemaDescriptor schemaDescriptor)
      {
         var nameHeaderCell = Cell(1).SetTextHeader("헤더명");
         var nextCell = nameHeaderCell.CellRight(attributesColumnIndex).GetColumnNumber();

         var lastUsedColumn = schemaDescriptor.AddSchemaDescriptionHeader(ActualRow, nextCell);

         var headerRange = Worksheet.Range(ActualRow.Get(), 1, ActualRow.Get(), lastUsedColumn);
         foreach (var cell in headerRange.Cells())
         {
            if (string.IsNullOrEmpty(cell.Value.ToString()))
            {
               cell.SetHeaderStyle();
            }
         }

         Worksheet.Cell(responseHeadertRowPointer, 1).SetSubHeaderStyle();
      }

      void InsertProperty(KeyValuePair<string, OpenApiHeader> openApiHeader, OpenApiSchemaDescriptor schemaDescriptor)
      {
         var nameCell = Cell(1).SetTextData(openApiHeader.Key);
         nameCell.Style.Font.SetFontName("Consolas");
         nameCell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(245, 245, 245));

         var nextCellNumber = nameCell.CellRight(attributesColumnIndex + 1).GetColumnNumber();

         nextCellNumber = schemaDescriptor.AddSchemaDescriptionValues(openApiHeader.Value.Schema, openApiHeader.Value.Required, ActualRow, nextCellNumber);

         var descCell = Cell(nextCellNumber);
         descCell.SetTextData(openApiHeader.Value.Description);
         descCell.Style.Alignment.SetWrapText(true);
      }
   }

   private void AddResponseHttpCode(string httpCode, string? description)
   {
      var isSuccessCode = httpCode.StartsWith("2") || httpCode.Equals("default");
      var responseCode = httpCode.Equals("default") ? "기본 응답" : $"HTTP {httpCode}";
      
      if (!string.IsNullOrEmpty(description) && !description.Equals("default response"))
      {
         responseCode += $": {description}";
      }

      var codeCell = Cell(1);
      codeCell.SetTextData(responseCode);
      
      if (isSuccessCode)
      {
         codeCell.SetSuccessResponseStyle();
      }
      else if (httpCode.StartsWith("4") || httpCode.StartsWith("5"))
      {
         codeCell.SetErrorResponseStyle();
      }
      else
      {
         codeCell.SetSubHeaderStyle();
      }

      ActualRow.MoveNext();
   }
}
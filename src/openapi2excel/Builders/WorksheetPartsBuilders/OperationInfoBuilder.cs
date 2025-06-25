using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class OperationInfoBuilder(
   RowPointer actualRow,
   int attributesColumnIndex,
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options)
   : WorksheetPartBuilder(actualRow, worksheet, options)
{
   public void AddOperationInfoSection(string path, OpenApiPathItem pathItem, OperationType operationType,
      OpenApiOperation operation)
   {
      Cell(1).SetTextBold("API ����");
      ActualRow.MoveNext();

      using (var _ = new Section(Worksheet, ActualRow))
      {
            var cell = Cell(1).SetTextBold("METHOD").CellRight(attributesColumnIndex).SetText(operationType.ToString().ToUpper())
               .IfNotEmpty(operation.OperationId, c => c.NextRow().SetTextBold("Id").CellRight(attributesColumnIndex).SetText(operation.OperationId))
               .NextRow().SetTextBold("���").CellRight(attributesColumnIndex).SetText(path)
               .IfNotEmpty(pathItem.Description, c => c.NextRow().SetTextBold("��� ����").CellRight(attributesColumnIndex).SetText(pathItem.Description))
               .IfNotEmpty(pathItem.Summary, c => c.NextRow().SetTextBold("��� ���").CellRight(attributesColumnIndex).SetText(pathItem.Summary));
            //.IfNotEmpty(operation.Description, c => c.NextRow().SetTextBold("Operation description").CellRight(attributesColumnIndex).SetText(operation.Description))
            //.IfNotEmpty(operation.Summary, c => c.NextRow().SetTextBold("Operation summary").CellRight(attributesColumnIndex).SetText(operation.Summary))
            //.NextRow().SetTextBold("Deprecated").CellRight(attributesColumnIndex).SetText(Options.Language.Get(operation.Deprecated));

         ActualRow.GoTo(cell.Address.RowNumber);
      }

      ActualRow.MoveNext(2);
   }
}
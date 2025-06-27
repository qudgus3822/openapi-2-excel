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
      // API 정보 제목 스타일 적용
      Cell(1).SetTextTitle("API 정보");
      ActualRow.MoveNext();

      using (var _ = new Section(Worksheet, ActualRow))
      {
         // HTTP 메서드 정보 - 메서드별 색상으로 강조
         var methodCell = Cell(1).SetTextSubHeader("METHOD");
         var methodValueCell = methodCell.CellRight(attributesColumnIndex);
         methodValueCell.SetMethodStyle(operationType.ToString()).SetText(operationType.ToString().ToUpper());

         // 다음 행으로 이동
         var cell = methodValueCell.NextRow();

         // Operation ID가 있는 경우 추가
         if (!string.IsNullOrEmpty(operation.OperationId))
         {
            cell.SetTextSubHeader("ID");
            cell.CellRight(attributesColumnIndex).SetTextData(operation.OperationId);
            cell = cell.NextRow();
         }

         // URL 경로 정보
         cell.SetTextSubHeader("경로");
         var pathCell = cell.CellRight(attributesColumnIndex);
         pathCell.SetTextData(path);
         pathCell.Style.Font.SetFontName("Consolas"); // 경로는 고정폭 폰트로
         pathCell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(245, 245, 245)); // 연한 회색 배경
         cell = cell.NextRow();

         // 경로 설명이 있는 경우 추가
         if (!string.IsNullOrEmpty(pathItem.Description))
         {
            cell.SetTextSubHeader("경로 설명");
            var descCell = cell.CellRight(attributesColumnIndex);
            descCell.SetTextData(pathItem.Description);
            descCell.Style.Alignment.SetWrapText(true);
            cell = cell.NextRow();
         }

         // 경로 요약이 있는 경우 추가
         if (!string.IsNullOrEmpty(pathItem.Summary))
         {
            cell.SetTextSubHeader("경로 요약");
            var summaryCell = cell.CellRight(attributesColumnIndex);
            summaryCell.SetTextData(pathItem.Summary);
            summaryCell.Style.Alignment.SetWrapText(true);
            cell = cell.NextRow();
         }

         // Operation 설명이 있는 경우 추가
         if (!string.IsNullOrEmpty(operation.Description))
         {
            cell.SetTextSubHeader("작업 설명");
            var opDescCell = cell.CellRight(attributesColumnIndex);
            opDescCell.SetTextData(operation.Description);
            opDescCell.Style.Alignment.SetWrapText(true);
            cell = cell.NextRow();
         }

         // Operation 요약이 있는 경우 추가
         if (!string.IsNullOrEmpty(operation.Summary))
         {
            cell.SetTextSubHeader("작업 요약");
            var opSummaryCell = cell.CellRight(attributesColumnIndex);
            opSummaryCell.SetTextData(operation.Summary);
            opSummaryCell.Style.Alignment.SetWrapText(true);
            cell = cell.NextRow();
         }

         // Deprecated 정보
         //if (operation.Deprecated)
         //{
         //   cell.SetTextSubHeader("사용 중단");
         //   var deprecatedCell = cell.CellRight(attributesColumnIndex);
         //   deprecatedCell.SetTextData("예");
         //   deprecatedCell.Style.Font.SetFontColor(XLColor.Red);
         //   deprecatedCell.Style.Font.SetBold(true);
         //   cell = cell.NextRow();
         //}

         ActualRow.GoTo(cell.Address.RowNumber - 1);
      }

      ActualRow.MoveNext(2);
   }
}
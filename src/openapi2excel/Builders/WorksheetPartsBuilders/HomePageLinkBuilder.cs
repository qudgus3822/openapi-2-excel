using ClosedXML.Excel;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class HomePageLinkBuilder(RowPointer actualRow, IXLWorksheet worksheet, OpenApiDocumentationOptions options)
   : WorksheetPartBuilder(actualRow, worksheet, options)
{
   public void AddHomePageLinkSection()
   {
      Cell(1).SetValue("Ç¥Áö·Î").SetHyperlink(new XLHyperlink($"'{InfoWorksheetBuilder.Name}'!A1"));
      ActualRow.MoveNext(2);
   }
}
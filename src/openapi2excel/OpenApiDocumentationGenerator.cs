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
            // C#
            var usedRange = worksheet.RangeUsed();
            if (usedRange != null)
            {
                foreach (var rangeColumn in usedRange.Columns())
                {
                    // rangeColumn.ColumnNumber()�� 1-based �ε���
                    worksheet.Column(rangeColumn.ColumnNumber()).AdjustToContents();
                }

                foreach (var rangeColumn in usedRange.Columns())
                {
                    double minWidth = 15; // ���ϴ� �ּ� �ʺ�(����)
                    var col = worksheet.Column(rangeColumn.ColumnNumber());
                    col.AdjustToContents();
                    if (col.Width < minWidth)
                        col.Width = minWidth;
                }
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
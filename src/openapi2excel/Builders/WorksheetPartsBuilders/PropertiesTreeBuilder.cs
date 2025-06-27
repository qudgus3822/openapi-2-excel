using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class PropertiesTreeBuilder(
   int attributesColumnIndex,
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options)
{
    private readonly int _attributesColumnIndex = attributesColumnIndex + 1;
    protected OpenApiDocumentationOptions Options { get; } = options;
    protected IXLWorksheet Worksheet { get; } = worksheet;
    private RowPointer ActualRow { get; set; } = null!;
    protected XLColor HeaderBackgroundColor => XLColor.LightGray;

    public void AddPropertiesTreeForMediaTypes(RowPointer actualRow, IDictionary<string, OpenApiMediaType> mediaTypes, OpenApiDocumentationOptions options)
    {
        ActualRow = actualRow;
        foreach (var mediaType in mediaTypes)
        {
            var bodyFormatRowPointer = ActualRow.Copy();
            var formatCell = Worksheet.Cell(ActualRow, 1);
            
            // Content-Type별 스타일 적용
            formatCell.SetTextSubHeader($"Content-Type: {mediaType.Key}");
            
            // Content-Type별 색상 적용
            var contentTypeColor = mediaType.Key.ToLower() switch
            {
                var ct when ct.Contains("json") => XLColor.FromArgb(217, 237, 247),      // 연한 파란색
                var ct when ct.Contains("xml") => XLColor.FromArgb(242, 222, 222),       // 연한 빨간색
                var ct when ct.Contains("form") => XLColor.FromArgb(223, 240, 216),      // 연한 초록색
                var ct when ct.Contains("text") => XLColor.FromArgb(252, 248, 227),      // 연한 노란색
                _ => XLColor.FromArgb(248, 248, 248)                                      // 연한 회색
            };
            formatCell.Style.Fill.SetBackgroundColor(contentTypeColor);
            
            ActualRow.MoveNext();

            if (mediaType.Value.Schema != null)
            {
                using (var _ = new Section(Worksheet, ActualRow))
                {
                    var columnCount = AddPropertiesTree(ActualRow, mediaType.Value.Schema, options);
                    ActualRow.MovePrev();
                }
                ActualRow.MoveNext(2);
            }
            else
            {
                ActualRow.MoveNext();
            }
        }
    }

    public int AddPropertiesTree(RowPointer actualRow, OpenApiSchema schema, OpenApiDocumentationOptions options)
    {
        ActualRow = actualRow;
        var columnCount = AddSchemaDescriptionHeader();
        var startColumn = CorrectRootElementIfArray(schema) ? 2 : 1;
        AddProperties(schema, startColumn, options);
        return columnCount;
    }

    protected bool CorrectRootElementIfArray(OpenApiSchema schema)
    {
        if (schema.Items == null)
            return false;

        AddPropertyRow("<array>", schema, false, 1);
        return true;
    }

    protected void AddProperties(OpenApiSchema? schema, int level, OpenApiDocumentationOptions options)
    {
        if (schema == null)
            return;

        if (schema.Items != null)
        {
            AddPropertiesForArray(schema, level, options);
        }
        if (schema.AllOf.Count == 1)
        {
            AddProperties(schema.AllOf[0], level, options);
        }
        if (schema.AnyOf.Count == 1)
        {
            AddProperties(schema.AnyOf[0], level, options);
        }
        foreach (var property in schema.Properties)
        {
            AddProperty(property.Key, property.Value, schema.Required.Contains(property.Key), level, options);
        }
    }

    private void AddPropertiesForArray(OpenApiSchema schema, int level, OpenApiDocumentationOptions options)
    {
        if (schema.Items.Properties.Any())
        {
            // array of object properties
            AddProperties(schema.Items, level, options);
        }
        else
        {
            // if array contains simple type items
            AddProperty("<value>", schema.Items, false, level, options);
        }
    }

    protected void AddProperty(string name, OpenApiSchema? schema, bool required, int level, OpenApiDocumentationOptions options)
    {
        if (schema == null || level >= options.MaxDepth)
        {
            return;
        }

        AddPropertyRow(name, schema, required, level++);
        AddProperties(schema, level, options);
    }

    private void AddPropertyRow(string propertyName, OpenApiSchema propertySchema, bool required, int propertyLevel)
    {
        var schemaDescriptor = new OpenApiSchemaDescriptor(Worksheet, Options);
        
        // 계층 구조 표시를 위한 들여쓰기 및 스타일링
        var indentedName = string.Concat(Enumerable.Repeat("  ", propertyLevel - 1)) + propertyName;
        
        // 특별한 속성명들에 대한 스타일링
        var nameCell = Worksheet.Cell(ActualRow, 1);
        nameCell.SetTextData(indentedName);
        
        if (propertyName.StartsWith("<") && propertyName.EndsWith(">"))
        {
            // 특수 속성 (<array>, <value> 등)은 이탤릭체로
            nameCell.Style.Font.SetItalic(true);
            nameCell.Style.Font.SetFontColor(XLColor.FromArgb(128, 128, 128));
        }
        else
        {
            // 일반 속성명
            nameCell.Style.Font.SetFontName("Consolas");
        }
        
        // 계층 레벨에 따른 배경색 적용 (들여쓰기 효과)
        if (propertyLevel > 1)
        {
            var indentColor = XLColor.FromArgb(250 - (propertyLevel * 10), 250 - (propertyLevel * 10), 250 - (propertyLevel * 10));
            nameCell.Style.Fill.SetBackgroundColor(indentColor);
        }
        
        nameCell.SetDataStyle();
        
        schemaDescriptor.AddSchemaDescriptionValues(propertySchema, required, ActualRow, _attributesColumnIndex);
        ActualRow.MoveNext();
    }

    protected int AddSchemaDescriptionHeader()
    {
        const int startColumn = 1;

        var schemaDescriptor = new OpenApiSchemaDescriptor(Worksheet, Options);
        schemaDescriptor.AddNameHeader(ActualRow, startColumn);
        var lastUsedColumn = schemaDescriptor.AddSchemaDescriptionHeader(ActualRow, _attributesColumnIndex);

        // 헤더 행 전체에 헤더 스타일 적용
        var headerRange = Worksheet.Range(ActualRow.Get(), startColumn, ActualRow.Get(), lastUsedColumn);
        foreach (var cell in headerRange.Cells())
        {
            if (string.IsNullOrEmpty(cell.Value.ToString()))
            {
                cell.SetHeaderStyle();
            }
        }

        ActualRow.MoveNext();
        return lastUsedColumn;
    }
}
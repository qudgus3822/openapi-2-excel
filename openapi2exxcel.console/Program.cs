// See https://aka.ms/new-console-template for more information
using openapi2exxcel.console;

Console.WriteLine(Console.Title = "OpenAPI to Excel Converter");
Console.WriteLine("Swagger yaml 주소를 입력하세요.");
var swaggerAddress = Console.ReadLine();
if (string.IsNullOrEmpty(swaggerAddress))
{
    swaggerAddress = "http://localhost:8080/swagger/v1/swagger.yaml";
    Console.WriteLine($"Swagger 주소가 비어 있습니다. 기본 주소로 설정합니다: {swaggerAddress}");
}

Console.WriteLine("출력 파일명을 입력하세요");
Console.WriteLine("출력 파일명은 확장자를 포함한 전체 파일명이어야 합니다.");

var documentationFileName = Console.ReadLine();


if (string.IsNullOrEmpty(documentationFileName))
{
    documentationFileName = "D://API 명세서.xlsx";
    Console.WriteLine($"출력 파일명이 비어 있습니다. 기본 파일명으로 설정합니다: {documentationFileName}");
}

ExcuteOpenApi2ExcelHelper helper = new ExcuteOpenApi2ExcelHelper();
try
{
    Console.WriteLine("문서 변환을 시작합니다...");
    await helper.ExecuteAsync(swaggerAddress, documentationFileName);
    Console.WriteLine("문서 변환이 완료되었습니다.");
}
catch (Exception ex)
{
    Console.WriteLine($"문서 변환에 실패했습니다. {ex.Message}");
}





using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using openapi2excel.core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace openapi2exxcel.console
{
    internal class ExcuteOpenApi2ExcelHelper
    {
        public async Task ExecuteAsync(string swaggerAddress, string outputFile, int depth = 10, bool noLogo = false, bool debug = false)
        {
            try
            {
                if (string.IsNullOrEmpty(swaggerAddress))
                {
                    throw new ArgumentNullException(nameof(swaggerAddress), "Swagger 주소가 비어 있습니다.");
                }

                var options = new OpenApiDocumentationOptions
                {
                    Language = OpenApiDocumentationLanguage.Ko,
                    MaxDepth = depth
                };

                await DownloadSwaggerYamlAsync(swaggerAddress);

                await OpenApiDocumentationGenerator.GenerateDocumentation("swagger.yaml", outputFile, options);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"문서 변환에 실패했습니다. {ex.Message}");
            }
        }

        private async Task DownloadSwaggerYamlAsync(string swaggerAddress, string documentationFileName = "swagger.yaml")
        {
            try
            {
                using var client = new WebClient();
                await client.DownloadFileTaskAsync(new Uri(swaggerAddress), documentationFileName);
            }
            catch
            {
                throw;
            }
        }
    }
}
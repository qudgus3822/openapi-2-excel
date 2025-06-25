namespace openapi2excel.core;

public class OpenApiDocumentationOptions
{
   private OpenApiDocumentationLanguage? _language;

   public OpenApiDocumentationLanguage Language
   {
      get => _language ?? OpenApiDocumentationLanguage.Default;
      set => _language = value;
   }

    public int MaxDepth { get; set; } = 10;
}

// TODO: Create language helper and refactor all text usage
public class OpenApiDocumentationLanguage : Dictionary<string, string>
{
   public string Get(string key)
      => TryGetValue(key, out var result) ? result : "<LANG_ERROR>";

   public string Get(bool value)
   {
      if (value)
      {
         return TryGetValue(OpenApiDocumentationLanguageConst.Yes, out var yesResult) ? yesResult : "<LANG_ERROR>";
      }
      return TryGetValue(OpenApiDocumentationLanguageConst.No, out var noResult) ? noResult : "<LANG_ERROR>";
   }

   public static OpenApiDocumentationLanguage Default
      => new()
      {
         { OpenApiDocumentationLanguageConst.Info, "Info" },
         { OpenApiDocumentationLanguageConst.Title, "Title" },
         { OpenApiDocumentationLanguageConst.Version, "Version" },
         { OpenApiDocumentationLanguageConst.Description, "Description" },

         { OpenApiDocumentationLanguageConst.Path, "Path" },
         { OpenApiDocumentationLanguageConst.PathSummary, "Path summary" },
         { OpenApiDocumentationLanguageConst.PathDescription, "Path description" },

         { OpenApiDocumentationLanguageConst.OperationType, "Operation type" },
         { OpenApiDocumentationLanguageConst.OperationSummary, "Operation summary" },
         { OpenApiDocumentationLanguageConst.OperationDescription, "Operation description" },
         { OpenApiDocumentationLanguageConst.Deprecated, "Deprecated" },

         { OpenApiDocumentationLanguageConst.Yes, "Yes" },
         { OpenApiDocumentationLanguageConst.No, "No" }
      };


    public static OpenApiDocumentationLanguage Ko
      => new()
      {
         { OpenApiDocumentationLanguageConst.Info, "표지" },
         { OpenApiDocumentationLanguageConst.Title, "제목" },
         { OpenApiDocumentationLanguageConst.Version, "버전" },
         { OpenApiDocumentationLanguageConst.Description, "설명" },

         { OpenApiDocumentationLanguageConst.Path, "경로" },
         { OpenApiDocumentationLanguageConst.PathSummary, "경로 요약" },
         { OpenApiDocumentationLanguageConst.PathDescription, "경로 설명" },

         { OpenApiDocumentationLanguageConst.OperationType, "Operation type" },
         { OpenApiDocumentationLanguageConst.OperationSummary, "Operation summary" },
         { OpenApiDocumentationLanguageConst.OperationDescription, "Operation description" },
         { OpenApiDocumentationLanguageConst.Deprecated, "Deprecated" },

         { OpenApiDocumentationLanguageConst.Yes, "Yes" },
         { OpenApiDocumentationLanguageConst.No, "No" }
      };
}

public static class OpenApiDocumentationLanguageConst
{
   public const string Info = "Info";
   public const string Title = "Title";
   public const string Version = "Version";
   public const string Description = "Description";

   public const string Path = "Path";
   public const string PathDescription = "PathDescription";
   public const string PathSummary = "PathSummary";

   public const string OperationType = "OperationType";
   public const string OperationDescription = "OperationDescription";
   public const string OperationSummary = "OperationSummary";
   public const string Deprecated = "Deprecated";

   public const string Yes = "Yes";
   public const string No = "No";
}
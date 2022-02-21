using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;

namespace MyNamespace;

public class Program
{
    static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
    static readonly string SpreedSheetId = "SpreedSheetId";
    static readonly string sheet = "Sheet1";
    public static SheetsService service;
    public static void Main()
    {
        GoogleCredential credential;

        using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
        {  
            credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
        }

        service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = "ForProgramming"
        });
        
        var list = new List<object>(){"shahboz", "khurramov", "some data"};
        SpreedSheetService.SaveData(SpreedSheetId, sheet, service, list);
        
    }
}

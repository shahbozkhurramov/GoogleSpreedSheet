using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

public class SpreedSheetService
{
    public static void SaveData(string SpreedSheetId, string sheet, SheetsService service, List<object> obj)
    {
        var valueRange = new ValueRange();
            
        valueRange.Values = new List<IList<object>>(){obj};

        var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreedSheetId, $"{sheet}!A:F");

        appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
        appendRequest.Execute();
    }

    public static void UpdateData(string SpreedSheetId, string sheet, SheetsService service, List<object> obj)
    {
        var valueRange = new ValueRange();

        valueRange.Values = new List<IList<object>>(){obj};

        var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreedSheetId, $"{sheet}!A:F");

        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
        updateRequest.Execute();
    }

    public static void ReadData(string SpreedSheetId, string sheet, SheetsService service)
    {
        var request = service.Spreadsheets.Values.Get(SpreedSheetId, $"{sheet}!A1:F10");

        var response = request.Execute();
        var values = response.Values;
        if(values != null && values.Count > 0)
        {
            foreach (var row in values)
            {
                //Do something
            }
        }
        else
        {
            Console.WriteLine("No data found");  
        }
    }

    public static void DeleteData(string SpreedSheetId, string sheet, SheetsService service)
    {
        var requestBody = new ClearValuesRequest();

        var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, SpreedSheetId, $"{sheet}!A:F");
        deleteRequest.Execute();
    }
}
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Newtonsoft.Json;

namespace GoogleSheetsAPI
{
    internal class Program
    {
        /* Global instance of the scopes required by this quickstart.
         If modifying these scopes, delete your previously saved token.json/ folder. */
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Google Sheets API .NET Quickstart";

        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            try
            {
                UserCredential credential;
                // Load client secrets.
                using (var stream =
                       new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    /* The file token.json stores the user's access and refresh tokens, and is created
                     automatically when the authorization flow completes for the first time. */
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.FromStream(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                // Create Google Sheets API service.
                var service = new SheetsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName
                });

                #region CreateSpreadsheet
                Spreadsheet spreadsheet = new Spreadsheet();
                SpreadsheetProperties properties = new SpreadsheetProperties();
                properties.Title = ApplicationName;
                spreadsheet.Properties = properties;

                SpreadsheetsResource.CreateRequest request = service.Spreadsheets.Create(spreadsheet);
                Spreadsheet response = request.Execute();

                string spreadsheetId = response.SpreadsheetId;
                Console.WriteLine("表格ID：" + spreadsheetId);                
                #endregion

                #region Modify spreadsheet cell's value.
                spreadsheetId = "1kfh4fshPRPkPdrtDs2uHYelB0eLtm3-AjpDpuRacnw8";

                ValueRange requestBody = new ValueRange();
                requestBody.Values = new List<IList<object>>()
{
    new List<object>() { "Value A1+", "Value B1+" },
    new List<object>() { "Value A2+", "Value B2+" }
};
                string range = "工作表1!A1:B2";
                SpreadsheetsResource.ValuesResource.UpdateRequest updateRequest = service.Spreadsheets.Values.Update(requestBody, spreadsheetId, range);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                updateRequest.Execute();


                BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();

                AddSheetRequest addSheetRequest = new AddSheetRequest();
                SheetProperties sheetProperties = new SheetProperties();
                sheetProperties.Title = "工作簿 2";
                addSheetRequest.Properties = sheetProperties;

                batchUpdateSpreadsheetRequest.Requests = new List<Request> { new Request { AddSheet = addSheetRequest } };
                var batchUpdateResponse = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadsheetId).Execute();
                #endregion

                #region Read spreadsheet cell's value
                // Define request parameters.
                spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
                 range = "Class Data!A2:E";
                SpreadsheetsResource.ValuesResource.GetRequest ValuesResourcerequest =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

                // Prints the names and majors of students in a sample spreadsheet:
                // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
                ValueRange ValueRangeresponse = ValuesResourcerequest.Execute();
                IList<IList<Object>> values = ValueRangeresponse.Values;
                if (values == null || values.Count == 0)
                {
                    Console.WriteLine("No data found.");
                    return;
                }
                Console.WriteLine("Name, Major");
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    Console.WriteLine("{0}, {1}", row[0], row[4]);
                }
                #endregion


            }
            catch (FileNotFoundException e)
            {
                Console.WriteLine(e.Message);
            }
        }

    }
}

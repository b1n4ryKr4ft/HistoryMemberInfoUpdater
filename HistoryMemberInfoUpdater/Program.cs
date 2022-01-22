using OfficeOpenXml;
using System.Text;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using HistoryMemberInfoUpdater.Models;

ReadDataFromExcelFile();



static void ReadDataFromExcelFile()
{
    var filePath = @"C:\Users\SibongeleniM\OneDrive - Club Leisure Management\Documents\HMI_Update\2022 HMI Update La Rochelle.xlsx";
    FileInfo file = new FileInfo(filePath);

    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

    using (var package = new ExcelPackage(file))
    {
        var workbookSheets = package.Workbook.Worksheets;

        var sbData = new StringBuilder();

        sbData.Append("[");

        for (var s = 0; s < workbookSheets.Count; s++)
        {
            ConvertSheetDataToJson(workbookSheets[s], ref sbData);

            if (s <= (workbookSheets.Count - 1))
            {
                sbData.Append(",");
            }
        }
        sbData.Append("]");

        var json = sbData.ToString();

        var hmiImportList = JsonConvert.DeserializeObject<List<HistoryMemberInfo>>(json);

        ProcessCompanyHistoryMemberInfo(hmiImportList);
    }

    static void ConvertSheetDataToJson(ExcelWorksheet worksheet, ref StringBuilder sbData)
    {
        var rowCount = worksheet.Dimension.Rows;
        var ColCount = worksheet.Dimension.Columns;

        var colNames = new Dictionary<int, string>();

        for (int col = 1; col <= ColCount; col++)
        {
            var cells = worksheet.Cells[1, col];
            var value = cells.Value?.ToString() ?? "";
            if (value != "")
            {
                value = value.Replace(" ", "");
                value = value.Replace("#", "");
                value = value.Replace("<", "");
                value = value.Replace(">", "");
                value = value.Replace("(", "");
                value = value.Replace(")", "");
                value = value.Replace("_", "");
                value = value.ToLower();

                colNames.Add(col, value);
            }
        }

        for (int row = 2; row <= rowCount; row++)
        {
            var end = false;
            var rawText = new StringBuilder();
            for (int colindex = 1; colindex <= colNames.Count; colindex++)
            {
                var x = colNames.SingleOrDefault(p => p.Key == colindex);

                var col = x.Key;

                var cells = worksheet.Cells[row, col];

                if (colindex > 1)
                {
                    rawText.Append(",");
                }

                var value = (cells.Value?.ToString() ?? "");

                int z;
                double y;
                bool w;
                bool isInt = Int32.TryParse(value, out z);
                bool isDouble = Double.TryParse(value, out y);
                bool isBoolean = Boolean.TryParse(value, out w);

                value = $"\"{value}\"";
                if (value.Trim() != "")
                {
                    var valueConcat = x.Value.Replace("company$$", "CompanyID").Replace("member$$", "MemberID");

                    rawText.Append($"\"{valueConcat}\":{value}");
                }
                else
                {
                    Console.WriteLine(value);
                }
            }
            if (row <= rowCount)
            {
                rawText.Append(",");
            }
            rawText.Append($"\"sheetname\":\"{worksheet.Name}\"");

            if (end == true)
            {
                break;
            }
            if (row > 2) { sbData.Append(","); }

            sbData.Append($"{{{rawText}}}{Environment.NewLine}");
        }
    }

    static void ProcessCompanyHistoryMemberInfo(List<HistoryMemberInfo> hmiList)
    {

        var membersForLevy = hmiList.Select(x => x.MemberID).ToArray();

        var url = "https://localhost:44343/";

        var existingMembers = GetExistingMembersInfo(url, membersForLevy); // "call to API using <membersForLevy>";

        // new members to be added
        var newMembers = hmiList.Where(x => !existingMembers.Contains(x.MemberID)).ToList();

        if (newMembers.Count() > 0)
        {
            InsertNewHistoryMemberInfos(url, newMembers);
        }


        // existing members to be updated
        var existing = hmiList.Where(x => existingMembers.Contains(x.MemberID)).ToList();

        if(existing.Count() > 0)
        {
            UpdateExistingMembers(url, existing);
        }  
    }
}

static int[] GetExistingMembersInfo(string path, int[] members)
{

    var existingMembers = new List<int>();

    try
    {
        var bodyObject = new MemberCancellationViewModel
        {
            TaskID = 99999,
            MemberIDs = members
        };

        var json = JsonConvert.SerializeObject(bodyObject);
        var data = new StringContent(json, Encoding.UTF32, "application/json");

        HttpClient client = new HttpClient();
        client.BaseAddress = new Uri(path);
        // Add an Accept header for JSON format.  
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        var inputMessage = new HttpRequestMessage
        {
            Content = new StringContent(json, Encoding.UTF8, "application/json")
        };

        inputMessage.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        HttpResponseMessage message = client.PostAsync("api/HistoryMemberInfoExisting/Post", inputMessage.Content).Result;

        if (message.IsSuccessStatusCode)
        {
            var content = message.Content.ReadAsStream();

            using (StreamReader sr = new StreamReader(content))
            {
                var existingMembersStream = sr.ReadToEnd();

                existingMembers = JsonConvert.DeserializeObject<List<int>>(existingMembersStream);
            }
        }

        if (!message.IsSuccessStatusCode)
            throw new ArgumentException(message.ToString());

    }

    catch (Exception ex)
    {
        Console.WriteLine(ex.Message);
    }
    
    return existingMembers.ToArray();
}

static void InsertNewHistoryMemberInfos(string path, List<HistoryMemberInfo> newMembers)
{
    try
    {
        MakeHttpRequest(newMembers, path, "api/HistoryMemberInfoNew/Post");
    }
    catch (Exception ex) { }
}

static void UpdateExistingMembers(string path, List<HistoryMemberInfo> existingMembers)
{
    try
    {
        MakeHttpRequest(existingMembers, path, "api/HistoryMemberInfoUpdate/Post");
    }
    catch (Exception ex) { }
}
static void MakeHttpRequest(List<HistoryMemberInfo> membersList, string path, string apiName)
{
    var listChucks = ListExtensions.ChunkBy(membersList, 100);

    HttpClient client = new HttpClient();
    client.BaseAddress = new Uri(path);
    // Add an Accept header for JSON format.  
    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

    Console.WriteLine("API being called => {0}", apiName);
    Console.WriteLine("========================================================\n\n");
    Console.WriteLine("Numbers of Chunks {0}\n", listChucks.Count);

    for(var i=0; i<listChucks.Count; i++)
    {

        Console.WriteLine("\n...Processing Chunk Number => {0}\n", (i + 1));

        var currentChunk = JsonConvert.SerializeObject(listChucks[i]);

        var inputMessage = new HttpRequestMessage
        {
            Content = new StringContent(currentChunk, Encoding.UTF8, "application/json")
        };

        inputMessage.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));


        HttpResponseMessage message = client.PostAsync(apiName, inputMessage.Content).Result;

        if (message.IsSuccessStatusCode)
        {
            var content = message.Content.ReadAsStream();

            using (StreamReader sr = new StreamReader(content))
            {
                Console.WriteLine(sr.ReadToEnd());
            }
        }
    }  
}

public static class ListExtensions
{
    public static List<List<T>> ChunkBy<T>(this List<T> source, int chunkSize)
    {
        if(chunkSize%2 == 0)
        {
            return source
            .Select((x, i) => new { Index = i, Value = x })
            .GroupBy(x => x.Index / chunkSize)
            .Select(x => x.Select(v => v.Value).ToList())
            .ToList();
        }
        else
        {
            return null;
        }
    }
}
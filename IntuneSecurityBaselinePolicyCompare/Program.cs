using CompareIntuneBaselineCompare.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

if (args.Length != 3)
{
    Console.WriteLine("Usage: Program.exe input1.json input2.json output.json");
    return;
}

string inputFilePath1 = args[0];
string inputFilePath2 = args[1];
string outputFilePath = EnsurePathEndsWithSlash(args[2]);

Console.WriteLine(inputFilePath1);
Console.WriteLine(inputFilePath2);
Console.WriteLine(outputFilePath);

if (!File.Exists(inputFilePath1) || !File.Exists(inputFilePath2))
{
    Console.WriteLine("One or both input files do not exist.");
    return;
}

Dictionary<string, CompareResult> compareResult = new Dictionary<string, CompareResult>();

//string json1 = File.ReadAllText("c:\\intune\\Pilot-Windows10-Device-SecurityBaseline_01-11-2023-15-53-56.743.json");
//string json2 = File.ReadAllText("c:\\intune\\Pilot-Windows-10-Device-SecurityBaselineACSCWindowsHardening.json");
//string json1 = File.ReadAllText("c:\\intune\\1.json");
//string json2 = File.ReadAllText("c:\\intune\\2.json");
//string outputPath = "c:\\intune\\";

string json1 = File.ReadAllText(inputFilePath1);
string json2 = File.ReadAllText(inputFilePath2);
string outputPath = (outputFilePath);

string fileName = $"PolicyCompare_{DateTime.Now:ddMMyyyyhhmmss}.xlsx";

// Deserialize the JSON files into Root objects.
var policy1 = JsonConvert.DeserializeObject<Policy>(json1);
var policy2 = JsonConvert.DeserializeObject<Policy>(json2);
var flatPolicy1 = new Dictionary<string, object>();
var flatPolicy2 = new Dictionary<string, object>();

flatPolicy1["displayName"] = policy1.DisplayName;
flatPolicy1["description"] = policy1.Description;
flatPolicy1["roleScopeTagIds"] = string.Join("|", policy1.RoleScopeTagIds);
flatPolicy1["TemplateDisplayName"] = policy1.TemplateDisplayName;
flatPolicy1["TemplateId"] = policy1.TemplateId;
flatPolicy1["versionInfo"] = policy1.VersionInfo;


flatPolicy2["displayName"] = policy2.DisplayName;
flatPolicy2["description"] = policy2.Description;
flatPolicy2["roleScopeTagIds"] = string.Join("|", policy2.RoleScopeTagIds);
flatPolicy2["TemplateDisplayName"] = policy2.TemplateDisplayName;
flatPolicy2["TemplateId"] = policy2.TemplateId;
flatPolicy2["versionInfo"] = policy2.VersionInfo;

flatPolicy1 = flattenPolicy(flatPolicy1, policy1.SettingsDelta);
flatPolicy2 = flattenPolicy(flatPolicy2, policy2.SettingsDelta);

var distinctKeys = flatPolicy1.Keys.Union(flatPolicy2.Keys);

foreach (var key in distinctKeys)
{
    var value1 = flatPolicy1.ContainsKey(key) ? flatPolicy1[key] : "MISSING";
    var value2 = flatPolicy2.ContainsKey(key) ? flatPolicy2[key] : "MISSING";
    var values = new List<object>() { value1, value2 };
    var result = GetResult(value1, value2);

    compareResult.Add(key, new CompareResult() { Values = values, Result = result });
}
GenerateOutputInExcel(distinctKeys, compareResult, outputPath + fileName);
Console.WriteLine($"Complete:{outputPath + fileName}");

Dictionary<string, object> flattenPolicy(Dictionary<string, object> flatPolicy, List<SettingsDeltum> settingsDelta)
{
    if (settingsDelta != null || settingsDelta.Any())
    {
        foreach (var settings in settingsDelta)
        {
            if (!IsComplex(settings) && !IsAbstract(settings))
            {
                flatPolicy[settings.DefinitionId] = settings.Value;
            }
            else if (IsComplex(settings) && !IsAbstract(settings))
            {
                flatPolicy[settings.DefinitionId] = "ARRAY";
                var nestedSettings = ((JArray)settings.Value).ToObject<List<SettingsDeltum>>();
                flatPolicy = flattenPolicy(flatPolicy, nestedSettings);
            }
            else if (IsAbstract(settings))
            {
                flatPolicy = flattenAbstractSettings(flatPolicy, settings);
            }
        }
    }
    return flatPolicy;
}

static string EnsurePathEndsWithSlash(string path)
{
    if (!path.EndsWith("\\"))
    {
        path += "\\";
    }
    return path;
}
Dictionary<string, object> flattenAbstractSettings(Dictionary<string, object> flatPolicy, SettingsDeltum settings)
{
    flatPolicy[settings.DefinitionId] = settings.ImplementationId;

    if (settings.ImplementationId != null)
    {
        Dictionary<string, string> keyValue = GetKVFromJson(settings.ValueJson);
        foreach (var kv in keyValue)
        {
            if (kv.Key.Equals("$implementationId")) continue;
            flatPolicy[settings.DefinitionId + "_" + kv.Key] = kv.Value;
        }
    }

    return flatPolicy;
}

Dictionary<string, string> GetKVFromJson(string valueJson)
{
    JObject jsonObject = JObject.Parse(valueJson);

    // Create a dictionary to store key-value pairs
    Dictionary<string, string> keyValues = new Dictionary<string, string>();

    // Iterate over the properties of the JObject
    foreach (var property in jsonObject.Properties())
    {
        string key = property.Name;
        string value = property.Value.ToString(Formatting.None); // Convert the value to a string

        keyValues[key] = value;
    }

    return keyValues;
}

bool IsAbstract(SettingsDeltum settings)
{
    return settings.OdataType.Contains("AbstractComplexSettingInstance");
}

bool IsComplex(SettingsDeltum settings)
{
    return settings.OdataType.Contains("Complex") && !settings.OdataType.Contains("Abstract");
}
string GetResult(object value1, object value2)
{

    if (value1 == null && value2 == null) 
        return "MATCH";
    else if (value1 == null || value2 == null || value1.Equals("MISSING") || value2.Equals("MISSING") || value1.Equals("notConfigured") || value2.Equals("notConfigured"))
        return "DIFF";
    else if (value1.Equals(value2))
        return "MATCH";
    else
        return "CONFLICT";
}
void GenerateOutputInExcel(IEnumerable<string> distinctKeys, Dictionary<string, CompareResult> compareResult, string filePath)
{
    IWorkbook workbook = new XSSFWorkbook();

    // Create a worksheet
    ISheet worksheet = workbook.CreateSheet("SampleSheet");

    // Sample data
    string[] headers = { "Setting", compareResult["displayName"].Values[0]?.ToString(), compareResult["displayName"].Values[1].ToString(), "Status" };


    // Create the header row
    IRow headerRow = worksheet.CreateRow(0);
    for (int i = 0; i < headers.Length; i++)
    {
        headerRow.CreateCell(i).SetCellValue(headers[i]);
    }
    int rowIndex = 0;
    foreach (string key in distinctKeys)
    {

        IRow dataRow = worksheet.CreateRow(rowIndex + 1);
        dataRow.CreateCell(0).SetCellValue(key);
        for (int columnIndex = 1; columnIndex < 3; columnIndex++)
        {
            dataRow.CreateCell(columnIndex).SetCellValue(compareResult[key].Values[columnIndex - 1]?.ToString());
        }
        dataRow.CreateCell(3).SetCellValue(compareResult[key].Result);
        rowIndex++;
    }


    // Save the Excel file
    using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
    {
        workbook.Write(fileStream);
    }
}

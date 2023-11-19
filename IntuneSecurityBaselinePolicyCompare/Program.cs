using CompareIntuneBaselineCompare.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Net.Http.Json;

// Check if the correct number of command-line arguments (3) is provided.
if (args.Length != 3)
{
    Console.WriteLine($"Usage: Program.exe \"input1.json\" \"input2.json\" \"output.json\"");
    return;
}

// Extract the command-line arguments for input and output file paths.
string inputFilePath1 = args[0];
string inputFilePath2 = args[1];
string outputFilePath = EnsurePathEndsWithSlash(args[2]);
const string MISSING = "MISSING";
const string NOTCONFIGURED = "notConfigured";
const string ARRAY = "ARRAY";
string[] metaData = new string[] { "displayName", "description", "roleScopeTagIds", "TemplateDisplayName", "TemplateId", "versionInfo" };
string[] ignoreProperties = new string[] { "@odata.type", "id", "lastModifiedDateTime", "createdDateTime" };
// Display the input and output file paths.
Console.WriteLine("Input File 1: " + inputFilePath1);
Console.WriteLine("Input File 2: " + inputFilePath2);
Console.WriteLine("Output File: " + outputFilePath);

// Check if input files exist.
if (!File.Exists(inputFilePath1) || !File.Exists(inputFilePath2))
{
    Console.WriteLine("One or both input files do not exist.");
    return;
}

// Create a dictionary to store comparison results.
Dictionary<string, CompareResultRow> compareResult = new Dictionary<string, CompareResultRow>();

// Read JSON content from the input files.
string json1 = File.ReadAllText(inputFilePath1);
string json2 = File.ReadAllText(inputFilePath2);
string outputPath = outputFilePath;
var flatPolicy1 = new Dictionary<string, object>();
var flatPolicy2 = new Dictionary<string, object>();

// Generate a unique output file name.
string fileName = $"BaselinePolicyCompare_{DateTime.Now:ddMMyyyyhhmmss}.xlsx";

if (IsTemplatedPolicy(json1))
{
    flatPolicy1 = FlatternBaselinePolicy(json1, flatPolicy1);
}
else
{
    flatPolicy1 = FlatternNonBaseLinePolicy(json1, flatPolicy1);
}


if (IsTemplatedPolicy(json2))
{
    flatPolicy2 = FlatternBaselinePolicy(json2, flatPolicy2);
}
else
{
    flatPolicy2 = FlatternNonBaseLinePolicy(json2, flatPolicy2);
}



// Get distinct keys from both flattened policies.
var distinctKeys = flatPolicy1.Keys.Union(flatPolicy2.Keys);

// Compare values and generate results for each key.
foreach (var key in distinctKeys)
{
    if(ignoreProperties.Contains(key)) continue;

    var value1 = flatPolicy1.ContainsKey(key) ? flatPolicy1[key] : MISSING;
    var value2 = flatPolicy2.ContainsKey(key) ? flatPolicy2[key] : MISSING;
    var values = new List<object>() { value1, value2 };
    var result = GetResult(key, value1, value2);

    // Add comparison results to the dictionary.
    compareResult.Add(key, new CompareResultRow() { Values = values, Result = result });
}

// Generate the comparison result in an Excel file.
GenerateOutputInExcel(distinctKeys, compareResult, outputPath + fileName);

// Display completion message with the path to the generated Excel file.
Console.WriteLine($"Complete: {outputPath + fileName}");

Dictionary<string, object> FlatternNonBaseLinePolicy(string json, Dictionary<string, object> flatPolicy)
{
    JObject jsonObject = JObject.Parse(json);
    flatPolicy = Flatten(jsonObject);
    return flatPolicy;
}

Dictionary<string, object> Flatten(JToken token, string prefix = "")
{
    var dict = new Dictionary<string, object>();

    if (token.Type == JTokenType.Object)
    {
        foreach (JProperty prop in token.Children<JProperty>())
        {
            string propPath = $"{prefix}_{prop.Name}";
            var innerDict = Flatten(prop.Value, propPath);

            foreach (var kvp in innerDict)
            {
                dict.Add(kvp.Key, kvp.Value);
            }
        }
    }
    else if (token.Type == JTokenType.Array)
    {
        dict.Add(prefix.Trim('_'), JTokenArrayToString(token));
    }
    else
    {
        dict.Add(prefix.Trim('_'), ((JValue)token).Value);
    }

    return dict;
}

string GetContentAfterLastUnderscore(string input)
{
    int lastUnderscoreIndex = input.LastIndexOf('_');

    if (lastUnderscoreIndex != -1 && lastUnderscoreIndex < input.Length - 1)
    {
        // Get the content after the last underscore
        return input.Substring(lastUnderscoreIndex + 1);
    }

    // If no underscore found or it's the last character, return an empty string or handle it as needed
    return string.Empty;
}

// Function to flatten policy data by processing settings deltas.
Dictionary<string, object> FlattenPolicySettings(Dictionary<string, object> flatPolicy, List<SettingsDeltum> settingsDelta, string prefix = "")
{
    string definitationId = string.Empty;

    if (settingsDelta != null || settingsDelta.Any())
    {
        foreach (var settings in settingsDelta)
        {
            definitationId = $"{prefix}_{GetContentAfterLastUnderscore(settings.DefinitionId)}";
            definitationId = definitationId.Trim('_');

            if (!IsComplex(settings) && !IsAbstract(settings))
            {
                flatPolicy[definitationId] = settings.Value;
            }
            else if (IsComplex(settings) && !IsAbstract(settings))
            {
                flatPolicy[definitationId] = ARRAY;
                var nestedSettings = ((JArray)settings.Value).ToObject<List<SettingsDeltum>>();
                flatPolicy = FlattenPolicySettings(flatPolicy, nestedSettings, definitationId);
            }
            else if (IsAbstract(settings))
            {
                flatPolicy = FlattenAbstractSettings(flatPolicy, settings);
            }
        }
    }
    return flatPolicy;
}

string JTokenArrayToString(JToken arrayToken)
{
    if (arrayToken != null && arrayToken.Type == JTokenType.Array)
    {
        var array = (JArray)arrayToken;

        // Convert each JToken element to its string representation
        string[] stringArray = Array.ConvertAll(array.ToArray(), x => x.ToString());

        // Join the string representations using a delimiter (comma and space in this case)
        return string.Join(", ", stringArray);
    }

    return string.Empty;
}
bool IsTemplatedPolicy(string json)
{
    if (!string.IsNullOrWhiteSpace(json) && json.Contains("\"TemplateId\"")) return true;
    return false;
}

// Function to ensure the path ends with a slash (for directories).
static string EnsurePathEndsWithSlash(string path)
{
    if (!path.EndsWith("\\"))
    {
        path += "\\";
    }
    return path;
}
// Function to flatten abstract settings and add them to the flatPolicy dictionary.
Dictionary<string, object> FlattenAbstractSettings(Dictionary<string, object> flatPolicy, SettingsDeltum settings)
{
    string definitationId = string.Empty;
    definitationId = GetContentAfterLastUnderscore(settings.DefinitionId);
    // Store the implementation ID in the flatPolicy using the definition ID as the key.
    flatPolicy[definitationId] = settings.ImplementationId;

    if (settings.ImplementationId != null)
    {
        // Parse the JSON value and extract key-value pairs.
        Dictionary<string, string> keyValue = GetKVFromJson(settings.ValueJson);

        foreach (var kv in keyValue)
        {
            // Skip the special key "$implementationId" and construct the key for the flatPolicy.
            if (kv.Key.Equals("$implementationId"))
                continue;

            // Add key-value pairs to the flatPolicy with a combined key.
            flatPolicy[definitationId + "_" + kv.Key] = kv.Value;
        }
    }

    return flatPolicy;
}

// Function to parse a JSON string and extract key-value pairs into a dictionary.
Dictionary<string, string> GetKVFromJson(string valueJson)
{
    // Parse the JSON content into a JObject.
    JObject jsonObject = JObject.Parse(valueJson);

    // Create a dictionary to store key-value pairs.
    Dictionary<string, string> keyValues = new Dictionary<string, string>();

    // Iterate over the properties of the JObject.
    foreach (var property in jsonObject.Properties())
    {
        string key = property.Name;
        string value = property.Value.ToString(Formatting.None); // Convert the value to a string

        // Add the key-value pair to the dictionary.
        keyValues[key] = value;
    }

    return keyValues;
}

// Function to check if the given settings are abstract.
bool IsAbstract(SettingsDeltum settings)
{
    return settings.OdataType.Contains("AbstractComplexSettingInstance");
}

// Function to check if the given settings are complex and not abstract.
bool IsComplex(SettingsDeltum settings)
{
    return settings.OdataType.Contains("Complex") && !settings.OdataType.Contains("Abstract");
}

// Function to determine the result based on two values for comparison.

//DIFF: When two values are different e.g. missing setting in another file or not configured in another in another file.
//MATCH : When two values are exact match
//CONFLICT:When two values are set by user and different in other policy.
string GetResult(string key, object value1, object value2)
{
    if (metaData.Contains(key))
        return string.Empty;

    if (value1 == null && value2 != null || value1 != null && value2 == null)
        return "DIFF";
    else if (value1 == value2)
        return "MATCH";
    else if (value1 != null && value1.Equals(value2))
        return "MATCH";
    else if (
        value1 != value2
        && (value1.Equals(NOTCONFIGURED)
        || value2.Equals(NOTCONFIGURED)
        || value1.Equals(MISSING)
        || value2.Equals(MISSING)))
        return "DIFF";
    else
        return "CONFLICT";
}

// Function to generate an Excel output with comparison results.
void GenerateOutputInExcel(IEnumerable<string> distinctKeys, Dictionary<string, CompareResultRow> compareResult, string filePath)
{
    // Create a new Excel workbook.
    IWorkbook workbook = new XSSFWorkbook();

    // Create a worksheet within the workbook.
    ISheet worksheet = workbook.CreateSheet("Settings Comparison");

    // Define the headers for the Excel file.
    string[] headers = { "Setting", compareResult["displayName"].Values[0]?.ToString(), compareResult["displayName"].Values[1].ToString(), "Status" };

    // Create the header row in the worksheet.
    IRow headerRow = worksheet.CreateRow(0);
    for (int i = 0; i < headers.Length; i++)
    {
        headerRow.CreateCell(i).SetCellValue(headers[i]);
    }

    int rowIndex = 0;

    // Fill the worksheet with comparison results.
    foreach (string key in distinctKeys)
    {
        if(ignoreProperties.Contains(key)) continue;

        IRow dataRow = worksheet.CreateRow(rowIndex + 1);
        dataRow.CreateCell(0).SetCellValue(key);
        for (int columnIndex = 1; columnIndex < 3; columnIndex++)
        {
            dataRow.CreateCell(columnIndex).SetCellValue(compareResult[key].Values[columnIndex - 1]?.ToString());
        }
        dataRow.CreateCell(3).SetCellValue(compareResult[key].Result);
        rowIndex++;
    }

    // Save the Excel file to the specified filePath.
    using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
    {
        workbook.Write(fileStream);
    }
}

Dictionary<string, object> FlatternBaselinePolicy(string json, Dictionary<string, object> flatPolicy)
{
    // Deserialize the JSON files into Policy objects.
    var policy = JsonConvert.DeserializeObject<BaselinePolicy>(json);
    // Populate the flatPolicy dictionaries with basic policy information.
    flatPolicy["displayName"] = policy.DisplayName;
    flatPolicy["description"] = policy.Description;
    flatPolicy["roleScopeTagIds"] = string.Join("|", policy.RoleScopeTagIds);
    flatPolicy["TemplateDisplayName"] = policy.TemplateDisplayName;
    flatPolicy["TemplateId"] = policy.TemplateId;
    flatPolicy["version"] = policy.VersionInfo;
    flatPolicy = FlattenPolicySettings(flatPolicy, policy.SettingsDelta);
    return flatPolicy;
}
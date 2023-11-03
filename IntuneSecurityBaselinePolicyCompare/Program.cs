using CompareIntuneBaselineCompare.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

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
string[] metaData = new string[] { "displayName", "description" , "roleScopeTagIds", "TemplateDisplayName" , "TemplateId", "versionInfo" };

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
Dictionary<string, CompareResult> compareResult = new Dictionary<string, CompareResult>();

// Read JSON content from the input files.
string json1 = File.ReadAllText(inputFilePath1);
string json2 = File.ReadAllText(inputFilePath2);
string outputPath = outputFilePath;

// Generate a unique output file name.
string fileName = $"PolicyCompare_{DateTime.Now:ddMMyyyyhhmmss}.xlsx";

// Deserialize the JSON files into Policy objects.
var policy1 = JsonConvert.DeserializeObject<Policy>(json1);
var policy2 = JsonConvert.DeserializeObject<Policy>(json2);
var flatPolicy1 = new Dictionary<string, object>();
var flatPolicy2 = new Dictionary<string, object>();

// Populate the flatPolicy dictionaries with basic policy information.
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

// Flatten policy data by processing settings deltas.
flatPolicy1 = FlattenPolicy(flatPolicy1, policy1.SettingsDelta);
flatPolicy2 = FlattenPolicy(flatPolicy2, policy2.SettingsDelta);

// Get distinct keys from both flattened policies.
var distinctKeys = flatPolicy1.Keys.Union(flatPolicy2.Keys);

// Compare values and generate results for each key.
foreach (var key in distinctKeys)
{
    var value1 = flatPolicy1.ContainsKey(key) ? flatPolicy1[key] :MISSING;
    var value2 = flatPolicy2.ContainsKey(key) ? flatPolicy2[key] : MISSING;
    var values = new List<object>() { value1, value2 };
    var result = GetResult(key,value1, value2);

    // Add comparison results to the dictionary.
    compareResult.Add(key, new CompareResult() { Values = values, Result = result });
}

// Generate the comparison result in an Excel file.
GenerateOutputInExcel(distinctKeys, compareResult, outputPath + fileName);

// Display completion message with the path to the generated Excel file.
Console.WriteLine($"Complete: {outputPath + fileName}");

// Function to flatten policy data by processing settings deltas.
Dictionary<string, object> FlattenPolicy(Dictionary<string, object> flatPolicy, List<SettingsDeltum> settingsDelta)
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
                flatPolicy[settings.DefinitionId] = ARRAY;
                var nestedSettings = ((JArray)settings.Value).ToObject<List<SettingsDeltum>>();
                flatPolicy = FlattenPolicy(flatPolicy, nestedSettings);
            }
            else if (IsAbstract(settings))
            {
                flatPolicy = FlattenAbstractSettings(flatPolicy, settings);
            }
        }
    }
    return flatPolicy;
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
    // Store the implementation ID in the flatPolicy using the definition ID as the key.
    flatPolicy[settings.DefinitionId] = settings.ImplementationId;

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
            flatPolicy[settings.DefinitionId + "_" + kv.Key] = kv.Value;
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
void GenerateOutputInExcel(IEnumerable<string> distinctKeys, Dictionary<string, CompareResult> compareResult, string filePath)
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


using PPTCreatorApp;
using System.Reflection;

string GetJsonPath()
{
    string exePath = Assembly.GetExecutingAssembly().Location;
    // Get the directory of the .exe file
    return Path.GetDirectoryName(exePath);
}
string jsonDirectoryPath = GetJsonPath();
string jsonPath = $@"{jsonDirectoryPath}" + "/presentationData.json";
if (!File.Exists(jsonPath))
{
    File.Create(jsonPath).Dispose(); // Create and close the file
    Console.WriteLine($"File created successfully at {jsonPath}");
}

Console.WriteLine($"Using JSON file at {jsonPath}");

FileDataHandler handler = new($@"{jsonDirectoryPath}", "presentationData");

DynamicJsonDataHandler jsonReader = new DynamicJsonDataHandler();
PowerPointData pData = await jsonReader.ReadJson(handler.GetRawJson());
if (pData != null)
{
    PPTCreator pptCreator = new PPTCreator(handler, pData);
    pptCreator.CreatePowerPointFile();
}
else
{
    Console.WriteLine("No Data Found");
}

Console.ReadLine();
using System;
using System.IO;
using System.Text.Json; // For JSON serialization and deserialization

namespace PPTCreatorApp
{
    internal class FileDataHandler
    {
        private string _dataDirPath = string.Empty;
        private string _dataFileName = string.Empty;

        public FileDataHandler(string dirPath, string fileName)
        {
            _dataDirPath = dirPath;
            _dataFileName = fileName;
        }

        public string GetRawJson()
        {
            string fullPath = Path.Combine(_dataDirPath, _dataFileName + ".json");
            PowerPointData loadedData = null;

            string dataToLoad = string.Empty;
            if (File.Exists(fullPath))
            {
                try
                {
                    dataToLoad = File.ReadAllText(fullPath);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Loading data has failed: " + e);
                }
            }
            return dataToLoad;
        }

        public PowerPointData Load()
        {
            string fullPath = Path.Combine(_dataDirPath, _dataFileName + ".json");
            PowerPointData loadedData = null;

            if (File.Exists(fullPath))
            {
                try
                {
                    string dataToLoad = File.ReadAllText(fullPath);
                    loadedData = JsonSerializer.Deserialize<PowerPointData>(dataToLoad);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Loading data has failed: " + e);
                }
            }
            return loadedData;
        }

        public void Save(PowerPointData pptData)
        {
            string fullPath = Path.Combine(_dataDirPath, _dataFileName + ".json");

            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(fullPath));
                string dataToStore = JsonSerializer.Serialize(pptData, new JsonSerializerOptions { WriteIndented = true });

                File.WriteAllText(fullPath, dataToStore);
                Console.WriteLine(dataToStore);
            }
            catch (Exception e)
            {
                Console.WriteLine("Saving data has failed: " + e);
            }
        }
    }
}

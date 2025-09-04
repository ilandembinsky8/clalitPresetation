using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Collections;

namespace PPTCreatorApp
{
    internal class DynamicJsonDataHandler
    {
        public async Task<PowerPointData> ReadJson(string data)
        {
            var dynamicJsonData = JsonConvert.DeserializeObject<dynamic>(data); // Parse the JSON

            PowerPointData pData = new PowerPointData
            {
                SlideCount = ReadInt(dynamicJsonData["SlideCount"]),
                Sadna = ReadString(dynamicJsonData["Sadna"]),
                Link = ReadString(dynamicJsonData["Link"]),
                Date = ReadString(dynamicJsonData["Date"]),
                Hour = ReadString(dynamicJsonData["Hour"]),
                TextsData = ListConverter<TextItem>(dynamicJsonData, "TextsData"),
                ImagesData = ListConverter<ImageItem>(dynamicJsonData, "ImagesData"),
                PostersData = ListConverter<PosterItem>(dynamicJsonData, "PostersData")
            }; // safely setup pData

            return pData;
        }

        public int ReadInt(dynamic? dynamicInt)
        {
            if (dynamicInt == null)
                return -1;

            return dynamicInt;
        }
        public string ReadString(dynamic? dynamicString)
        {
            if (dynamicString == null)
                return "";

            return dynamicString;
        }
        public bool ReadBool(dynamic? dynamicBool)
        {
            if (dynamicBool == null)
                return false;

            return dynamicBool;
        }
        public List<T> ListConverter<T>(dynamic dynamiC, string jsonProperty)
        {
            var array = dynamiC[jsonProperty] as JArray; // Get the JArray from the dynamic object
            if (array == null) // Check if the array is null
            {
                Console.WriteLine($"Property '{jsonProperty}' is null or not a valid JArray.");
                return new List<T>(); // Return an empty list as fallback
            }
            return array.ToObject<List<T>>(); // Deserialize the JArray into a List<T>
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTCreatorApp
{
    internal class PosterSubItem
    {
        public string Name { get; set; }
        public string ItemType { get; set; }
        public int ItemNumber { get; set; }
        public string File { get; set; }
        public float X { get; set; }
        public float Y { get; set; }
        public float W { get; set; }
        public float H { get; set; }
        public float RotatedAngle { get; set; }

        public PosterSubItem(string name, string itemType, int itemNumber, string file, float x, float y, float w, float h, float rotatedAngle)
        {
            Name = name;
            ItemType = itemType;
            ItemNumber = itemNumber;
            File = file;
            X = x;
            Y = y;
            //W = 200.0f; 
            W = w; 
            //H = 200.0f;
            H = h;
            RotatedAngle = rotatedAngle;
        }
    }
}

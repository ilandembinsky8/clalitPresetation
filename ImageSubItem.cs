using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTCreatorApp
{
    internal class ImageSubItem
    {
        public string Text { get; set; }
        public string FeelingText { get; set; }
        public string TheContent { get; set; }
        public string TheFile { get; set; }
        
        public ImageSubItem(string text, string feelingText, string theContent, string theFile)
        {
            Text = text;
            FeelingText = feelingText;
            TheContent = theContent;
            TheFile = theFile;
        }
    }
}

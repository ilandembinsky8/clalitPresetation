using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTCreatorApp
{
    internal class TextSubItem
    {
        public int ItemOrder { get; set; }
        public string Paragraph { get; set; }

        public TextSubItem(int itemOrder, string paragraph)
        {
            ItemOrder = itemOrder;
            Paragraph = paragraph;
        }
    }
}

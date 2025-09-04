using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTCreatorApp
{
    internal class TextItem : BaseItem
    {
        public List<TextSubItem> SubItems { get; set; }

        public TextItem(int orderInWorkshop, string itemName, string itemType, List<TextSubItem> subItems) : base(orderInWorkshop, itemName, itemType)
        {
            SubItems = subItems;
        }
    }
}

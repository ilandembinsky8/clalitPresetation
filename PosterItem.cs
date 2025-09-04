using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTCreatorApp
{
    internal class PosterItem : BaseItem
    {
        public List<PosterSubItem> SubItems { get; set; }

        public PosterItem(int orderInWorkshop, string itemName, string itemType, List<PosterSubItem> subItems) : base(orderInWorkshop, itemName, itemType)
        {
            SubItems = subItems;
        }
    }
}

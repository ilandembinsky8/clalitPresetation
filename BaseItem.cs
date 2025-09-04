namespace PPTCreatorApp
{
    internal class BaseItem
    {
        public int OrderInWorkshop {  get; set; }
        public string ItemType { get; set; }
        public string ItemName { get; set; }

        public BaseItem(int orderInWorkshop, string itemName, string itemType)
        {
            ItemName = itemName;
            ItemType = itemType;
        }
    }
}

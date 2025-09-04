namespace PPTCreatorApp
{
    internal class ImageItem : BaseItem
    {
        public string OriginalImage { get; set; } // file.jpg etc..
        public List<ImageSubItem> SubItems { get; set; }

        public ImageItem(int orderInWorkshop, string itemName, string itemType, string originalImage, List<ImageSubItem> subItems) : base(orderInWorkshop, itemName, itemType)
        {
            OriginalImage = originalImage;
            SubItems = subItems;
        }
    }
}
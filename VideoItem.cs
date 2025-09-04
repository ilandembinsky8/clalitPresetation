namespace PPTCreatorApp
{
    internal class VideoItem : BaseItem
    {
        public string VideoLink { get; set; }

        public VideoItem(int orderInWorkshop, string itemName, string itemType, string videoLink) : base(orderInWorkshop, itemName, itemType)
        {
            VideoLink = videoLink;
        }
    }
}
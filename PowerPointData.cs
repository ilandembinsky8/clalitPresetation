namespace PPTCreatorApp
{
    internal class PowerPointData
    {
        public int SlideCount { get; set; }

        public string Sadna { get; set; }
        public string Link { get; set; }
        public string Date { get; set; }
        public string Hour { get; set; }

        public List<TextItem> TextsData { get; set; }
        public List<ImageItem> ImagesData { get; set; }
        public List<PosterItem> PostersData { get; set; }

        public PowerPointData()
        {
            SlideCount = 0;

            Sadna = string.Empty;
            Link = string.Empty;
            Date = string.Empty;
            Hour = string.Empty;

            TextsData = new ();
            ImagesData = new();
            PostersData = new();
        }

        public PowerPointData(int slideCount, string sadna, string link, string date, string hour, List<TextItem> textsData, List<ImageItem> imgsData, List<PosterItem> postersData)
        {
            SlideCount = slideCount;

            Sadna = sadna;
            Link = link;
            Date = date;
            Hour = hour;

            TextsData = textsData;
            ImagesData = imgsData;
            PostersData = postersData;
        }
    }
}

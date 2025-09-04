using Microsoft.Office.Core;

namespace PPTCreatorApp
{
    public static class SlideContentDataFactory
    {
        private const string PloniFont = "Ploni ML v2 AAA";
        private const string PloniFontBold = "Ploni ML Bold AAA";
        private static System.Drawing.Color ClalitDarkBlue = System.Drawing.Color.FromArgb(0, 32, 96);

        // Main title slide vars
        private const float DefaultTitleBoxWidth = 600;
        public const float DefaultTitleBoxHeight = 100;
        public const float DefaultMainCameraImgWidth = 500.0f;
        public const float TitleBoxYOffsetMultiplier = 2.2f;
        public const int TitleFontSize = 48;
        public static TextBoxData CreateTitle(string text, float slideWidth, float slideHeight, float boxWidth = DefaultTitleBoxWidth, float boxHeight = DefaultTitleBoxHeight)
        {
            // Center the title box on the slide
            float x = slideWidth - boxWidth;
            float y = (float)(slideHeight - (boxHeight * TitleBoxYOffsetMultiplier));

            return new TextBoxData
            {
                Text = text,
                X = x,
                Y = y,
                Width = boxWidth,
                Height = boxHeight,
                FontName = PloniFont,
                FontSize = TitleFontSize,
                Bold = MsoTriState.msoFalse,
                Italic = MsoTriState.msoFalse,
                Underline = MsoTriState.msoFalse,
                Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignRight,
                TextColor = System.Drawing.Color.White
            };
        }

        // Content Slide title vars
        public const float DefaultSlideTitleBoxWidth = 400;
        public const float DefaultSlideTitleBoxHeight = 50;
        public const int SlideTitleFontSize = 28;
        public const float SlideTitleBoxYOffset = 13;
        public static TextBoxData CreateContentSlideTitle(string text, float slideWidth, float slideHeight, float boxWidth = DefaultSlideTitleBoxWidth, float boxHeight = DefaultSlideTitleBoxHeight)
        {
            // Center the title box on the slide
            float x = slideWidth - boxWidth;
            float y = boxHeight - SlideTitleBoxYOffset;

            return new TextBoxData
            {
                Text = text,
                X = x,
                Y = y,
                Width = boxWidth,
                Height = boxHeight,
                FontName = PloniFont,
                FontSize = SlideTitleFontSize,
                Bold = MsoTriState.msoFalse,
                Italic = MsoTriState.msoFalse,
                Underline = MsoTriState.msoFalse,
                Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignRight,
                TextColor = ClalitDarkBlue
            };
        }

        public static readonly string MainCameraImgPath = Path.Combine(AppContext.BaseDirectory, "Templates", "MainCameraImg.png");
        public static ImageData CreateMainCameraImg(float height, float width = 500.0f)
        {
            var imagePath = MainCameraImgPath;
            var imageBytes = System.IO.File.ReadAllBytes(imagePath);

            return new ImageData(
                -width / 20,
                0,
                width,
                height,
                imageBytes,
                imagePath,
                "image/png"
            );
        }
        public const float GoodluckTextBoxWidth = 250;
        public const float GoodluckTextBoxHeight = 100;
        public const float GoodluckTextBoxXOffset = 70;
        public const float GoodluckTextBoxFontSize = 60;
        public const float GoodluckTextBoxPadding = 70;

        public static TextBoxData GoodluckTextBox(float slideWidth, float slideHeight)
        {
            return new TextBoxData
            {
                Text = "בהצלחה!",
                X = slideWidth - GoodluckTextBoxWidth - GoodluckTextBoxPadding,
                Y = slideHeight - (GoodluckTextBoxHeight * 2),
                Width = GoodluckTextBoxWidth,
                Height = GoodluckTextBoxHeight,
                FontName = PloniFontBold,
                FontSize = GoodluckTextBoxFontSize,
                Bold = MsoTriState.msoFalse,
                Italic = MsoTriState.msoFalse,
                Underline = MsoTriState.msoFalse,
                Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignRight,
                TextColor = System.Drawing.Color.FromArgb(201, 255, 77)
            };
        }


        public static readonly string YearImgPath = Path.Combine(AppContext.BaseDirectory, "Templates", "Year.png");
        private static ImageData YearImg()
        {
            var imagePath = YearImgPath;
            var imageBytes = System.IO.File.ReadAllBytes(imagePath);

            return new ImageData(
                0,
                0,
                0,
                0,
                imageBytes,
                imagePath,
                "image/png"
            );
        }

        public const float YearTitleImgWidth = 120f;
        public const float YearTitleImgHeight = 70f;
        public const float YearTitleImgXMultiplier = 4f;
        public const float YearTitleImgYOffset = 7f;
        public static ImageData YearTitleSlideImg(float slideWidth, float slideHeight)
        {
            ImageData yearImg = YearImg();
            yearImg.Width = YearTitleImgWidth;
            yearImg.Height = YearTitleImgHeight;
            yearImg.X = (float)(slideWidth - (yearImg.Width * YearTitleImgXMultiplier));
            yearImg.Y = (float)(slideHeight - yearImg.Height) - YearTitleImgYOffset;
            return yearImg;
        }
        public const float YearClosingImgWidth = 375f;
        public const float YearClosingImgHeight = 225f;
        public const float YearClosingImgXMultiplier = 1.5f;
        public const float YearClosingImgYMultiplier = 0.3f;
        public static ImageData YearClosingSlideImg(float slideWidth, float slideHeight)
        {
            ImageData yearImg = YearImg();
            yearImg.Width = YearClosingImgWidth;
            yearImg.Height = YearClosingImgHeight;
            yearImg.X = (float)(slideWidth - (yearImg.Width * YearClosingImgXMultiplier));
            yearImg.Y = (float)(yearImg.Height * YearClosingImgYMultiplier);
            return yearImg;
        }

        internal static BulletTextBoxData CreateBulletTextBox(string[] bulletPoints, int[] levels, float slideWidth, float slideHeight)
        {
            if (bulletPoints.Length != levels.Length)
            {
                throw new ArgumentException("Bullet points and levels must have the same length.");
            }
            return new BulletTextBoxData
            {
                Text = string.Join("\n", bulletPoints),
                BulletPoints = bulletPoints,
                Levels = levels,
                X = 0,
                Y = slideHeight / 4,
                Width = slideWidth,
                Height = slideHeight,
                FontName = PloniFont,
                FontSize = 18,
                Bold = MsoTriState.msoFalse,
                Italic = MsoTriState.msoFalse,
                Underline = MsoTriState.msoFalse,
                Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignRight,
                TextColor = ClalitDarkBlue
            };
        }
    }
}
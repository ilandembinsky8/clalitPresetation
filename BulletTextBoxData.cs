using Microsoft.Office.Core;

namespace PPTCreatorApp
{
    public class BulletTextBoxData : TextBoxData
    {
        public string[] BulletPoints { get; set; }
        public int[] Levels { get; set; }

        public BulletTextBoxData()
        { }

        public BulletTextBoxData(string[] bulletPoints, int[] levels, float x, float y, float width, float height, string fontName, float fontSize, MsoTriState italic, MsoTriState underline, MsoTriState bold, Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft, System.Drawing.Color textColor = default) :
        base(x, y, width, height, fontName, fontSize, italic, underline, bold, alignment, textColor)
        {
            BulletPoints = bulletPoints;
            Text = bulletPoints != null ? string.Join("\n", bulletPoints) : string.Empty;
            Levels = levels;
            FontName = fontName;
            FontSize = fontSize;
            Italic = italic;
            Underline = underline;
            Bold = bold;
            Alignment = alignment;
            TextColor = textColor;
        }

    }
}
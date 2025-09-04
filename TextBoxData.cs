using Microsoft.Office.Core;

namespace PPTCreatorApp
{
    public class TextBoxData : SlideContentData
    {
        public string Text { get; set; }
        public string FontName { get; set; } = "Calibri";
        public float FontSize { get; set; } = 18;
        public MsoTriState Italic { get; set; } = MsoTriState.msoFalse;
        public Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment Alignment { get; set; } = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
        public MsoTriState Underline { get; set; } = MsoTriState.msoFalse;
        public MsoTriState Bold { get; set; } = MsoTriState.msoFalse;
        public System.Drawing.Color TextColor { get; set; } = System.Drawing.Color.Black;
        public TextSegmentStyle[] SegmentStyles { get; set; } = Array.Empty<TextSegmentStyle>();

        public TextBoxData()
        { }

        public TextBoxData(float x, float y, float width, float height, string fontName, float fontSize, MsoTriState italic, MsoTriState underline, MsoTriState bold, Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft, System.Drawing.Color textColor = default, TextSegmentStyle[] segmentStyles = null) :
        base(x, y, width, height)
        {
            FontName = fontName;
            FontSize = fontSize;
            Italic = italic;
            Underline = underline;
            Bold = bold;
            Alignment = alignment;
            TextColor = textColor;
            SegmentStyles = segmentStyles ?? Array.Empty<TextSegmentStyle>();
        }

    }
}
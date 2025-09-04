public class TextSegmentStyle
{
    public int Start { get; set; } // 0-based
    public int Length { get; set; }
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
    public System.Drawing.Color Color { get; set; }

    public TextSegmentStyle()
    {
        Color = System.Drawing.Color.Empty; // Default color
    }
}
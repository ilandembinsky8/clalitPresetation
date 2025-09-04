using System;

namespace PPTCreatorApp
{
    public class ImageData : SlideContentData
    {
        public byte[] Data { get; set; }
        public string FileName { get; set; }
        public string ContentType { get; set; }

        public ImageData(float x, float y, float width, float height, byte[] data, string fileName, string contentType) :
        base(x, y, width, height)
        {
            Data = data ?? throw new ArgumentNullException(nameof(data));
            FileName = fileName ?? throw new ArgumentNullException(nameof(fileName));
            ContentType = contentType ?? throw new ArgumentNullException(nameof(contentType));
        }
    }
}
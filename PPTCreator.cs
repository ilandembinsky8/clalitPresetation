using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Web;
using System.Numerics;
using System.Reflection;
using System.ComponentModel;
using Microsoft.Office.Interop.PowerPoint;

namespace PPTCreatorApp
{
    internal class PPTCreator
    {
        private string _filePath;
        private string _fileName;

        private string _exeFolderPath 
        {
            get
            {  // Get the full path of the executing assembly (the .exe file)
                string exePath = Assembly.GetExecutingAssembly().Location;

                // Get the directory of the .exe file
                return Path.GetDirectoryName(exePath);
            }
        }
        private string _saveDirectory;
        private FileDataHandler _fileDataHandler;
        private PowerPointData _pptData;

        private PowerPoint.Application _pptApp;
        private PowerPoint.Presentation _presentation;
        private PowerPoint.CustomLayout _cLayout;
        private int _slideCount = 0;

        private const string POSTER_SUB_TYPE_IMAGE = "Image";
        private const string POSTER_SUB_TYPE_TEXT = "Text";
        private const string POSTER_SUB_TYPE_FREE_TEXT = "FreeText";

        public PPTCreator(FileDataHandler fileHandler, PowerPointData pptData)
        {
            _pptData = pptData;
            _fileDataHandler = fileHandler;
        }

        private void FormatText(PowerPoint.TextFrame text, string fontName, float size, MsoTriState isBold, MsoTriState isItalic, MsoTriState isUnderlined, System.Drawing.Color textColor, PowerPoint.PpParagraphAlignment alignment)
        {
            text.TextRange.ParagraphFormat.TextDirection = PowerPoint.PpDirection.ppDirectionRightToLeft;
            PowerPoint.Font titleFont = text.TextRange.Font;
            titleFont.Name = fontName;
            titleFont.Size = size;
            titleFont.Bold = isBold;
            titleFont.Italic = isItalic;
            titleFont.Underline = isUnderlined;
            text.TextRange.ParagraphFormat.Alignment = alignment;
            text.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
            titleFont.Color.RGB = System.Drawing.ColorTranslator.ToOle(textColor);
        }

        private PowerPoint.CustomLayout GetCustomLayout(string layoutName, string designName)
        {
            foreach (PowerPoint.Design design in _presentation.Designs)
            {
                if (design.Name.Equals(designName, StringComparison.OrdinalIgnoreCase))
                {
                    foreach (PowerPoint.CustomLayout layout in design.SlideMaster.CustomLayouts)
                    {
                        if (layout.Name.Equals(layoutName, StringComparison.OrdinalIgnoreCase))
                        {
                            return layout;
                        }
                    }
                }
            }
            throw new ArgumentException($"Custom layout '{layoutName}' not found in design '{designName}'.");
        }

        private void CreateNewPresentation()
        {
            // create instace of ppt
            _pptApp = new PowerPoint.Application();

            // create new presentation
            _presentation = _pptApp.Presentations.Add(MsoTriState.msoTrue);

            // apply theme to the presentation
            // needs to be rewriten to recive the layout file on start
            _presentation.ApplyTemplate(Path.Combine(AppContext.BaseDirectory, "Templates", "CaliltTheme.thmx"));

            // add custom layout
            _cLayout = GetCustomLayout("Blank", "4_Office Theme");
            if (_cLayout == null)
            {
                Console.WriteLine("Custom layout 'Blank' not found. Using default layout.");
                _cLayout = _presentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutText];
                Console.WriteLine($"Custom layout set: {_cLayout.Name}");
            }
        }

        /// <summary>
        /// Adds a new slide to the presentation.
        /// </summary>
        /// <param name="layout">The layout to use for the slide.</param>
        /// <returns>The newly created slide.</returns>
        private PowerPoint.Slide AddNewSlide(PowerPoint.CustomLayout layout = null)
        {
            _slideCount++;
            PowerPoint.Slides slides = _presentation.Slides;
            return _presentation.Slides.AddSlide(_slideCount, layout ?? _cLayout);
        }

        /// <summary>
        /// Adds a new content slide to the presentation.
        /// </summary>
        /// <param name="slideTitle">The title of the slide.</param>
        /// <param name="layout">The layout to use for the slide.</param>
        /// <returns>The newly created content slide.</returns>
        private PowerPoint.Slide AddNewContentSlide(string slideTitle, PowerPoint.CustomLayout layout = null)
        {
            PowerPoint.Slide slide = AddNewSlide(layout);
            AddTextBox(slide, SlideContentDataFactory.CreateContentSlideTitle(slideTitle, _presentation.PageSetup.SlideWidth, _presentation.PageSetup.SlideHeight));
            AddBulletTextBox(slide, SlideContentDataFactory.CreateBulletTextBox(new[] { "נקודה ראשונה", "נקודה ראשונה ראשונה", "נקודה ראשונה שנייה", "נקודה שנייה", "נקודה שלישית" }, new[] { 1, 2, 2, 1, 1 }, _presentation.PageSetup.SlideWidth, _presentation.PageSetup.SlideHeight));
            return slide;
        }

        private void AddBulletTextBox(Slide slide, BulletTextBoxData bulletTextBoxData)
        {
            PowerPoint.Shape textBox = AddTextBox(slide, bulletTextBoxData);
            PowerPoint.TextRange paragraphs = textBox.TextFrame.TextRange.Paragraphs();
            int index = 0;
            foreach (PowerPoint.TextRange paragraph in paragraphs)
            {
                paragraph.ParagraphFormat.Bullet.Visible = MsoTriState.msoTrue;
                paragraph.ParagraphFormat.Bullet.Character = 8226; // Bullet character
                paragraph.IndentLevel = bulletTextBoxData.Levels[index];
                index++;
            }
        }

        /// <summary>
        /// Adds a text box to the slide.
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="textBoxData"></param>
        private PowerPoint.Shape AddTextBox(PowerPoint.Slide slide, TextBoxData textBoxData)
        {
            PowerPoint.Shape textBox = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                textBoxData.X,
                textBoxData.Y,
                textBoxData.Width,
                textBoxData.Height
            );
            textBox.TextFrame.TextRange.Text = textBoxData.Text;
            FormatText(textBox.TextFrame, textBoxData.FontName, textBoxData.FontSize, textBoxData.Bold, textBoxData.Italic, textBoxData.Underline, textBoxData.TextColor, textBoxData.Alignment);
            return textBox;
        }

        /// <summary>
        /// Adds an image to the slide.
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="imageData"></param>
        private PowerPoint.Shape AddImage(PowerPoint.Slide slide, ImageData imageData)
        {
            PowerPoint.Shape imageShape = slide.Shapes.AddPicture(
                imageData.FileName,
                MsoTriState.msoFalse,
                MsoTriState.msoCTrue,
                imageData.X,
                imageData.Y,
                imageData.Width,
                imageData.Height
            );
            return imageShape;
        }

        /// <summary>
        /// Sets the background image for the slide.
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="imagePath"></param>
        private void SetBackground(PowerPoint.Slide slide, string imagePath)
        {
            slide.FollowMasterBackground = MsoTriState.msoFalse;
            slide.Background.Fill.UserPicture(imagePath);
        }

        private void CreateSlides()
        {
            List<PowerPoint.Slide> slides = new();

            CreateFirstTileSlide(slides);
            CreateContentSlide(slides, "רקע");
            CreateClosingSlide(slides);
        }

        private void CreateContentSlide(List<PowerPoint.Slide> slides, string title)
        {
            slides.Add(AddNewContentSlide(title));
        }

        private void CreateClosingSlide(List<PowerPoint.Slide> slides)
        {
            slides.Add(AddNewSlide(GetCustomLayout("2_Blank", "4_Office Theme")));
            SetBackground(slides[slides.Count - 1], (Path.Combine(AppContext.BaseDirectory, "Templates", "ClosingSlideBackground.png")));
            AddTextBox(slides[slides.Count - 1], SlideContentDataFactory.GoodluckTextBox(_presentation.PageSetup.SlideWidth, _presentation.PageSetup.SlideHeight));
            AddImage(slides[slides.Count - 1], SlideContentDataFactory.YearClosingSlideImg(_presentation.PageSetup.SlideWidth, _presentation.PageSetup.SlideHeight));
        }

        public void CreateFirstTileSlide(List<PowerPoint.Slide> slides)
        {
            if (slides.Count != 0)
            {
                throw new ArgumentException("title slide has to be the first slide.");
            }
            slides.Add(AddNewSlide(GetCustomLayout("2_Blank", "4_Office Theme")));
            SetBackground(slides[0], (Path.Combine(AppContext.BaseDirectory, "Templates", "TitleSlideBackground.png")));
            AddImage(slides[0], SlideContentDataFactory.CreateMainCameraImg(_presentation.PageSetup.SlideHeight));
            AddTextBox(slides[0], SlideContentDataFactory.CreateTitle("כרמיאל\nתמונת מצב בראי התחרות", _presentation.PageSetup.SlideWidth, _presentation.PageSetup.SlideHeight));
            AddImage(slides[0], SlideContentDataFactory.YearTitleSlideImg(_presentation.PageSetup.SlideWidth, _presentation.PageSetup.SlideHeight));
        }

        public void CreatePowerPointFile()
        {
            // where images should be saved
            _saveDirectory = _exeFolderPath + @"\PresentationImages";
            if (!Directory.Exists(_saveDirectory))
            {
                Directory.CreateDirectory(_saveDirectory);
                Console.WriteLine($"Created presentation images directory: {_saveDirectory}");
            }

            CreateNewPresentation();

            CreateSlides();

            string filePath = $@"{_exeFolderPath}\PowerPoint_Presentation.pptx";
            _presentation.SaveAs(filePath, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
        }

        public void Dispose()
        {
            // Release COM objects to prevent memory leaks
            if (_presentation != null)
            {
                _presentation.Close();
                Marshal.ReleaseComObject(_presentation);
            }

            if (_pptApp != null)
            {
                _pptApp.Quit();
                Marshal.ReleaseComObject(_pptApp);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
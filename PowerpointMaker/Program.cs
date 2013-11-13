using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerpointMaker
{
    /*
        Discoveries:
     * 
     * Presentations.Open Method
     *  - Opens the specified presentation. 
     * Presentations.Open2007 Method
     *  - Opens the specified presentation and provides the option to repair the presentation file. 
     * 
     * http://msdn.microsoft.com/en-us/library/microsoft.office.interop.powerpoint.presentations.open2007(v=office.14).aspx
     * http://msdn.microsoft.com/en-us/library/microsoft.office.interop.powerpoint.presentations.open(v=office.14).aspx
     */


    internal class Maker
    {
        public static Maker Presentation {
            get
            {
                return new Maker();
            }
        }

        private readonly Application _powerPoint;
        private readonly List<Presentation> _presentations = new List<Presentation>();
        
        public Maker()
        {
            _powerPoint = new Application();
        }

        public PresentationWrapper New()
        {
            const MsoTriState withWindow = MsoTriState.msoFalse;
            var newPresentation = _powerPoint.Presentations.Add(withWindow);
            return RememberToCloseLater(newPresentation);
        }

        public PresentationWrapper OpenFrom(string templateFilePath)
        {
            if (!File.Exists(templateFilePath))
            {
                throw new FileNotFoundException(templateFilePath);
            }
            var templatedPresentation = OpenTemplate(templateFilePath);
            return RememberToCloseLater(templatedPresentation);
        }

        private Presentation OpenTemplate(string filename)
        {
            // These WORK!
//            const MsoTriState openReadOnly = MsoTriState.msoFalse;
//            const MsoTriState openACopy = MsoTriState.msoTrue;
//            const MsoTriState displayAWindow = MsoTriState.msoTrue;
            const MsoTriState openWritable = MsoTriState.msoFalse;
            const MsoTriState openACopy = MsoTriState.msoTrue;
            const MsoTriState displayAWindow = MsoTriState.msoFalse;
            return _powerPoint.Presentations.Open(filename, openWritable, openACopy, displayAWindow);
        }

        private PresentationWrapper RememberToCloseLater(Presentation presentation)
        {
            _presentations.Add(presentation);
            return new PresentationWrapper(presentation, this);
        }

        public void Done()
        {
            foreach (var presentation in _presentations)
            {
                presentation.Close();
            }
            _powerPoint.Quit();
        }
    }

    internal class PresentationWrapper
    {
        private readonly Presentation _presentation;
        private readonly Maker _maker;

        public PresentationWrapper(Presentation presentation, Maker maker)
        {
            _presentation = presentation;
            _maker = maker;
        }

        public SlideWrapper AddSlide()
        {
            var slide = _presentation.Slides.Add(1, PpSlideLayout.ppLayoutText);
            return new SlideWrapper(slide, this);
        }

        public Maker Show()
        {
            _presentation.SlideShowSettings.Run();
            return _maker;
        }

        public Maker Save(string filename)
        {
            // First I thought - What would be best? Powerpoint, HTML, Open Document or PDF? But then I thought...  - JH, 13.11.2013
            http://www.youtube.com/watch?v=iqc7CEE0ekE

            SaveAsOpenDocumentPresentation(filename);
            SaveAsPdf(filename);
            SaveAsPresentation(filename);
            return _maker;
        }

        private void SaveAsPdf(string filename)
        {
            const MsoTriState andEmbedFonts = MsoTriState.msoCTrue;
            const PpSaveAsFileType asPdf = PpSaveAsFileType.ppSaveAsPDF;
            _presentation.SaveAs(filename, asPdf, andEmbedFonts);
        }

        private void SaveAsOpenDocumentPresentation(string filename)
        {
            const MsoTriState andEmbedFonts = MsoTriState.msoCTrue;
            const PpSaveAsFileType asOpenDocumentPresentation = PpSaveAsFileType.ppSaveAsOpenDocumentPresentation;
            _presentation.SaveAs(filename, asOpenDocumentPresentation, andEmbedFonts);
        }

        private void SaveAsPresentation(string filename)
        {
            // Under PowerPoint 2010 this is most likely a pptx file.
            const PpSaveAsFileType asPresentation = PpSaveAsFileType.ppSaveAsDefault;
            const MsoTriState andEmbedFonts = MsoTriState.msoCTrue;
            _presentation.SaveAs(filename, asPresentation, andEmbedFonts);
        }
    }

    class SlideWrapper
    {
        private readonly Slide _slide;
        private readonly PresentationWrapper _daddy;

        public SlideWrapper(Slide slide, PresentationWrapper daddy)
        {
            _slide = slide;
            _daddy = daddy;
        }

        public SlideWrapper Font(string name, double size)
        {
            //  range.Font.Name = "Arial";
            //  range.Font.Size = 48;
            return this;
        }

        public SlideWrapper Top(string text)
        {
            var range = _slide.Shapes[1].TextFrame.TextRange;
            range.Text = "Hello";
            return this;
        }

        public SlideWrapper Center(string text)
        {
            var range = _slide.Shapes[2].TextFrame.TextRange;
            range.Text = "Hello";
            return this;
        }

        public PresentationWrapper Bottom(string text)
        {
            var range = _slide.Shapes[2].TextFrame.TextRange;
            range.Text = "Hello";
            return _daddy;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            const string template = "Pink Template.potx";

            new Maker()
                .OpenFrom(template)
                .AddSlide()
                    .Top("Hello")
                    .Center("Good")
                    .Bottom("Bye")
                .AddSlide()
                    .Top("Hello")
                    .Center("Hello")
                    .Bottom("Hello")
                .Save("Presentation")
                .Done();
        }
    }
}

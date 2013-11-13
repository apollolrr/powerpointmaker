using System.Drawing;
using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerpointMaker
{
    public class Content : BaseSlide
    {
        private readonly Slide _slide;

        public Content(Slide slide)
        {
            _slide = slide;
        }

        public Content Top(string text)
        {
            var range = _slide.Shapes[2].TextFrame.TextRange;
            range.Text = text;
            return this;
        }

        public Content Center(string text)
        {
            var range = _slide.Shapes[1].TextFrame.TextRange;
            range.Text = text;
            return this;
        }

        public Content Bottom(string text)
        {
            var range = _slide.Shapes[3].TextFrame.TextRange;
            range.Text = text;
            return this;
        }
    }

    public class Sourcecode : BaseSlide
    {
        private readonly Slide _slide;

        public Sourcecode(Slide slide)
        {
            _slide = slide;
        }

        public Sourcecode HightlightLine(int number)
        {           
            var textFrame = _slide.Shapes[1].TextFrame;
            var textRange = textFrame.TextRange.Sentences(number, 1);
            var white = Color.White.ToArgb();
            textRange.Font.Color.RGB = white;
            return this;
        }

        public Sourcecode FontSize(int number)
        {
            var textFrame = _slide.Shapes[1].TextFrame;
            textFrame.TextRange.Font.Size = number;
            return this;
        }


        // 166,166,166 // White
        public Sourcecode Code(string filename)
        {
            var code = File.ReadAllText(filename);
            var textFrame = _slide.Shapes[1].TextFrame;
            var textRange = textFrame.TextRange;
            var grey = Color.FromArgb(255, 166, 166, 166).ToArgb();
            textRange.Font.Color.RGB = grey;
            textRange.Text = code;
            return this;
        }
    }

    public class TitleSlide : BaseSlide
    {
        private readonly Slide _slide;

        public TitleSlide(Slide slide)
        {
            _slide = slide;
        }

        public TitleSlide Title(string text)
        {
            var range = _slide.Shapes[1].TextFrame.TextRange.Text = text;
            return this;
        }
    }

    public class Image : BaseSlide
    {
        private readonly Slide _slide;

        public Image(Slide slide)
        {
            _slide = slide;
        }

        public Image Caption(string text)
        {
            _slide.Shapes[2].TextFrame.TextRange.Text = text;
            return this;
        }

        public Image Title(string text)
        {
            _slide.Shapes[3].TextFrame.TextRange.Text = text;
            return this;
        }

        public Image File(string filename)
        {
            filename = Path.Combine(Directory.GetCurrentDirectory(), filename);
            if (!System.IO.File.Exists(filename))
            {
                throw new FileNotFoundException(filename);
            }
            _slide.Shapes.AddPicture(filename, MsoTriState.msoFalse, MsoTriState.msoTrue, 0,0,-1,-1);
            return this;
        }
    }

    public class BaseSlide
    {
    }
}
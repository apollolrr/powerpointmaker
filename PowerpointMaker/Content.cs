using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

// Refactoring Point - JH, 13.11.2013
namespace PowerpointMaker
{
    public abstract class BaseSlide
    {
        protected readonly Slide Slide;
        private readonly PresentationWrapper _daddy;

        public BaseSlide(Slide slide, PresentationWrapper daddy)
        {
            Slide = slide;
            _daddy = daddy;
        }

        public BaseSlide AddSlide(string layoutName)
        {
            return _daddy.AddSlide(layoutName);
        }

        public abstract void Parse(string[] lines);

        public void Save(string filename)
        {
            _daddy.Save(filename);
        }
    }

    public class Content : BaseSlide
    {
        public Content(Slide slide, PresentationWrapper presentationWrapper) : base(slide, presentationWrapper)
        {
        }

        public Content Top(string text)
        {
            var range = Slide.Shapes[2].TextFrame.TextRange;
            range.Text = text;
            return this;
        }

        public Content Center(string text)
        {
            var range = Slide.Shapes[1].TextFrame.TextRange;
            range.Text = text;
            return this;
        }

        public Content Bottom(string text)
        {
            var range = Slide.Shapes[3].TextFrame.TextRange;
            range.Text = text;
            return this;
        }

        // Refactoring Point - JH, 13.11.2013
        public override void Parse(string[] lines)
        {
            if (lines.Length > 3)
            {
                throw new TooManyLinesComplaint();
            }
            Top(lines[0])
            .Center(lines[1])
            .Bottom(lines[2]);
        }
    }

    public class TooManyLinesComplaint : Exception
    {
    }

    public class Sourcecode : BaseSlide
    {

        public Sourcecode(Slide slide, PresentationWrapper PresentationWrapper)
            : base(slide, PresentationWrapper)
        {
        }

        public Sourcecode HightlightLine(int number)
        {           
            var textFrame = Slide.Shapes[1].TextFrame;
            var textRange = textFrame.TextRange.Sentences(number, 1);
            var white = Color.White.ToArgb();
            textRange.Font.Color.RGB = white;
            return this;
        }

        public Sourcecode FontSize(int number)
        {
            var textFrame = Slide.Shapes[1].TextFrame;
            textFrame.TextRange.Font.Size = number;
            return this;
        }


        // 166,166,166 // White
        public Sourcecode Code(string filename)
        {
            var code = File.ReadAllText(filename);
            var textFrame = Slide.Shapes[1].TextFrame;
            var textRange = textFrame.TextRange;
            var grey = Color.FromArgb(255, 166, 166, 166).ToArgb();
            textRange.Font.Color.RGB = grey;
            textRange.Text = code;
            return this;
        }

        public override void Parse(string[] lines)
        {
            Code(lines[0]);
            FontSize(ParseFontSize(lines));
            foreach (var linenumber in ParseHighlight(lines))
            {
                HightlightLine(linenumber);
            }
        }

        private int ParseFontSize(string[] lines)
        {
            var size = (lines.SingleOrDefault(line => line.Contains("Size")) ?? "").Trim().Split(' ').LastOrDefault() ?? "";
            return int.Parse(size);
        }

        // Refactoring Point - JH, 13.11.2013
        private IEnumerable<int> ParseHighlight(string[] lines)
        {
            var keyword = "Highlight";
            var trim = (lines.SingleOrDefault(line => line.Contains(keyword)) ?? "").Trim();
            if (string.IsNullOrWhiteSpace(trim))
            {
                return new int[]{};
            }
            return trim.Remove(0, keyword.Length).Split(',').Select(int.Parse);
        }
    }

    public class TitleSlide : BaseSlide
    {

        public TitleSlide(Slide slide, PresentationWrapper PresentationWrapper)
            : base(slide, PresentationWrapper)
        {
        }

        public TitleSlide Title(string text)
        {
            var range = Slide.Shapes[1].TextFrame.TextRange.Text = text;
            return this;
        }

        // Refactoring Point - JH, 13.11.2013
        public override void Parse(string[] lines)
        {
            Title(lines[0]);
        }
    }

    public class Image : BaseSlide
    {
        public Image(Slide slide, PresentationWrapper PresentationWrapper)
            : base(slide, PresentationWrapper)
        {
        }

        public Image Caption(string text)
        {
            Slide.Shapes[2].TextFrame.TextRange.Text = text;
            return this;
        }

        public Image Title(string text)
        {
            Slide.Shapes[3].TextFrame.TextRange.Text = text;
            return this;
        }

        public Image File(string filename)
        {
            filename = Path.Combine(Directory.GetCurrentDirectory(), filename);
            if (!System.IO.File.Exists(filename))
            {
                throw new FileNotFoundException(filename);
            }
            Slide.Shapes.AddPicture(filename, MsoTriState.msoFalse, MsoTriState.msoTrue, 0,0,-1,-1);
            return this;
        }

        // Refactoring Point - JH, 13.11.2013
        public override void Parse(string[] lines)
        {
            File(lines[0]);
            Caption(lines[1]);
            Title(lines[2]);
        }
    }
}
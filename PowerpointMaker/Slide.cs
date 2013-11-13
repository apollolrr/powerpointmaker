using Microsoft.Office.Interop.PowerPoint;

namespace PowerpointMaker
{
    public class Slide
    {
        private readonly Microsoft.Office.Interop.PowerPoint.Slide _slide;
        private readonly Presentation _daddy;

        public Slide(Microsoft.Office.Interop.PowerPoint.Slide slide, Presentation daddy)
        {
            _slide = slide;
            _daddy = daddy;
        }

        public Slide Top(string text)
        {
            var range = _slide.Shapes[1].TextFrame.TextRange;
            range.Text = text;
            return this;
        }

        public Slide Center(string text)
        {
            var range = _slide.Shapes[2].TextFrame.TextRange;
            range.Text = text;
            return this;
        }

        public Presentation Bottom(string text)
        {
            var range = _slide.Shapes[3].TextFrame.TextRange;
            range.Text = text;
            return _daddy;
        }

        public Presentation Ok()
        {
            return _daddy;
        }
    }
}
using Microsoft.Office.Interop.PowerPoint;

namespace PowerpointMaker
{
    public class SlideWrapper
    {
        private readonly Slide _slide;
        private readonly PresentationWrapper _daddy;

        public SlideWrapper(Slide slide, PresentationWrapper daddy)
        {
            _slide = slide;
            _daddy = daddy;
        }

        public SlideWrapper Top(string text)
        {
            var range = _slide.Shapes[1].TextFrame.TextRange;
            range.Text = text;
            return this;
        }

        public SlideWrapper Center(string text)
        {
            var range = _slide.Shapes[2].TextFrame.TextRange;
            range.Text = text;
            return this;
        }

        public PresentationWrapper Bottom(string text)
        {
            var range = _slide.Shapes[3].TextFrame.TextRange;
            range.Text = text;
            return _daddy;
        }

        public PresentationWrapper Ok()
        {
            return _daddy;
        }
    }
}
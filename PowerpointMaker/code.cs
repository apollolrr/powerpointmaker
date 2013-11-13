namespace PowerpointMaker
{
    class code
    {
        public dynamic AddSlide(string layoutName)
        {
            if (!_layouts.ContainsKey(layoutName))
            {
                throw new UnknownLayoutException(layoutName, _layouts);
            }

            var index = _presentation.Slides.Count;
            var slide = _presentation.Slides.AddSlide(index, _layouts[layoutName]);
            return new Content(slide);
        }
    }
}

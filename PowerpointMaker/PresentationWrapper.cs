﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerpointMaker
{
    public class PresentationWrapper : IDisposable
    {
        private readonly Microsoft.Office.Interop.PowerPoint.Presentation _presentation;
        private readonly Maker _maker;
        private readonly Dictionary<string, CustomLayout> _layouts = new Dictionary<string, CustomLayout>();       

        public PresentationWrapper(Microsoft.Office.Interop.PowerPoint.Presentation presentation, Maker maker)
        {
            _presentation = presentation;
            _maker = maker;
            IdentifyLayouts();
        }

        private void IdentifyLayouts()
        {
            foreach (CustomLayout customLayout in _presentation.SlideMaster.CustomLayouts)
            {
                _layouts.Add(customLayout.MatchingName, customLayout);
            }
        }

        public dynamic AddTitleSlide()
        {          
            var index = _presentation.Slides.Count;
            var slide = _presentation.Slides.Add(index, PpSlideLayout.ppLayoutTitleOnly);
            return new TitleSlide(slide, this);
        }

        public dynamic AddSlide(string layoutName)
        {

            if (!_layouts.ContainsKey(layoutName))
            {
                throw new UnknownLayoutException(layoutName, _layouts);
            }

            var index = _presentation.Slides.Count;
            var slide = _presentation.Slides.AddSlide(index, _layouts[layoutName]);

            // Refactoring Point - JH, 13.11.2013
            if(layoutName == "Content")
                return new Content(slide, this);

            if (layoutName == "Sourcecode")
                return new Sourcecode(slide, this);

            if (layoutName == "Image")
                return new Image(slide, this);

            return null;
        }

        public Maker Show()
        {
            _presentation.SlideShowSettings.Run();
            return _maker;
        }

        public Maker Save(string filename)
        {
            filename = Path.Combine(Directory.GetCurrentDirectory(), filename);
            
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
            const PpSaveAsFileType asPresentation = PpSaveAsFileType.ppSaveAsOpenXMLPresentation;
            const MsoTriState andEmbedFonts = MsoTriState.msoCTrue;
            _presentation.SaveAs(filename, asPresentation, andEmbedFonts);
        }

        public void Dispose()
        {
            _presentation.Close();
        }
    }
}
using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerpointMaker
{
    public class Maker : IDisposable
    {
        public static Maker Presentation {
            get
            {
                return new Maker();
            }
        }

        private readonly Application _powerPoint;
        
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

        public PresentationWrapper OpenFrom(string filename)
        {
            var templateFilePath = AbsolutePathFor(filename);
            var presentation = OpenTemplate(templateFilePath);
            return RememberToCloseLater(presentation);
        }

        private static string AbsolutePathFor(string filename)
        {
            if (!File.Exists(filename))
            {
                throw new FileNotFoundException(filename);
            }
            return new FileInfo(filename).FullName;
        }

        private Microsoft.Office.Interop.PowerPoint.Presentation OpenTemplate(string filename)
        {
            const MsoTriState openWritable = MsoTriState.msoFalse;
            const MsoTriState openACopy = MsoTriState.msoTrue;
            const MsoTriState displayAWindow = MsoTriState.msoFalse;
            return _powerPoint.Presentations.Open(filename, openWritable, openACopy, displayAWindow);
        }

        private PresentationWrapper RememberToCloseLater(Microsoft.Office.Interop.PowerPoint.Presentation presentation)
        {          
            return new PresentationWrapper(presentation, this);
        }

        public void Dispose()
        {
            _powerPoint.Quit();
        }
    }
}
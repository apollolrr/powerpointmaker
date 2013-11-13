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
        private readonly List<Presentation> _presentations = new List<Presentation>();
        
        public Maker()
        {
            _powerPoint = new Application();
        }

//        public PresentationWrapper New()
//        {
//            const MsoTriState withWindow = MsoTriState.msoFalse;
//            var newPresentation = _powerPoint.Presentations.Add(withWindow);
//            return RememberToCloseLater(newPresentation);
//        }

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

        private Presentation OpenTemplate(string filename)
        {
            const MsoTriState openWritable = MsoTriState.msoFalse;
            const MsoTriState openACopy = MsoTriState.msoTrue;
            const MsoTriState displayAWindow = MsoTriState.msoFalse;
            //MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue
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
            _presentations.Clear();
            _powerPoint.Quit();
        }

        public void Dispose()
        {
            Done();
        }
    }
}
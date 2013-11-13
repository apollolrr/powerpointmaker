using System.Linq;

namespace PowerpointMaker
{
    /*
     *  Knowledge taken from here: 
     *  http://support.microsoft.com/kb/303718
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

    public class Program
    {
        public static void Main(string[] args)
        {
            const string template = "Pink Template.potx";
            using (var maker = new Maker())
            using(var presentation = maker.OpenFrom(template))
            {

                presentation
                    .AddSlide()
                    .Title("Title");

                presentation.
                    AddSlide("Content")
                    .Top("Top")
                    .Center("Center")
                    .Bottom("Bottom");

                presentation.
                    AddSlide("Sourcecode")
                    .Code("code.cs")
                    .FontSize(18)
                    .HightlightLine(1)
                    .HightlightLine(9)
                    .HightlightLine(14);

                presentation
                    .AddSlide("Content")
                    .Top("One")
                    .Center("Two")
                    .Bottom("Three");

                presentation
                    .AddSlide("Image")
                    .Caption("Caption")
                    .Title("Title")
                    .File("Image.jpg");

                presentation
                    .Save("Presentation");
            }
        }
    }
}
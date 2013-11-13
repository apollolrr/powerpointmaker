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
            ParseMal();
            MachMal();
        }

        private static void ParseMal()
        {
            new Parser().Go("slides.dsl");
        }

        private static void MachMal()
        {
            const string template = "Pink Template.potx";
            using (var maker = new Maker())
            using (var presentation = maker.OpenFrom(template))
            {
                presentation
                    .AddTitleSlide()
                        .Title("I <3 NY")
                    .AddSlide("Image")
                        .File("image.jpg")
                        .Caption("Brooklyn Bridge")
                        .Title("A Day in Manhattan")
                    .AddSlide("Content")
                        .Top("I love...")
                        .Center("New York")
                        .Bottom("...this city")
                    .AddSlide("Content")
                        .Top("Stay away from")
                        .Center("CENTRAL PARK")
                        .Bottom("during nighttime...")
                    .AddSlide("Sourcecode")
                        .Code("code.cs")
                        .FontSize(18)
                        .HightlightLine(1)
                        .HightlightLine(9)
                        .HightlightLine(14)
                    .Save("I-love-NY");
            }
        }
    }
}
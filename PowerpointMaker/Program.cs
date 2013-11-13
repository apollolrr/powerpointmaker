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
            {
                maker
                    .OpenFrom(template)
                    .AddSlide()
                        .Top("Title")
                        .Ok()

                    .AddSlide("Content")
                        .Top("Hello")
                        .Center("Good")
                        .Bottom("Bye")

                    .AddSlide("Content")
                        .Top("Hello")
                        .Center("Hello")
                        .Bottom("Hello")

                    .Save("Presentation");
            }            
        }
    }
}

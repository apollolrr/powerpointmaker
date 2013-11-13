using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerpointMaker
{
    internal class UnknownLayoutException : Exception
    {
        public UnknownLayoutException(string name, Dictionary<string, CustomLayout> layouts) : base(BuildMessage(name, layouts))
        {
            
        }

        private static string BuildMessage(string name, Dictionary<string, CustomLayout> layouts)
        {
            var message = new StringBuilder();
            message.AppendFormat("I don't know this layout \"{0}\". All I got is\n", name);
            foreach (var layout in layouts)
            {
                message.AppendLine(layout.Key);
            }
            return message.ToString();
        }

    }
}
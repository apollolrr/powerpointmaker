using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace PowerpointMaker
{
    // Refactoring Point - JH, 13.11.2013
    public class Parser
    {
        private Dictionary<string, string> _sectionmap = new Dictionary<string, string>()
            {
                { "Title", "Title"},
                { "Content", "Content"},
                { "Image", "Image"},
                { "Code", "Sourcecode"},
            };
        

        public void Go(string filename)
        {
            var lines = File.ReadAllLines(filename, Encoding.UTF8);


            using (var maker = new Maker())
            using (var presentation = maker.OpenFrom("Pink Template.potx"))
            {
                var partitions = PartitionDemShit(lines);
                foreach (var partition in partitions)
                {
                    BaseSlide slide = null;
                    if (partition.Item1 == "Title")
                    {
                        slide = ((BaseSlide) presentation.AddTitleSlide());
                    }
                    else
                    {
                        slide = (BaseSlide)presentation.AddSlide(partition.Item1);
                    }

                    slide.Parse(partition.Item2.Select(s => s.Trim()).ToArray());
                }
                presentation.Save("Generated_bitches");
            }
        }

        private IEnumerable<Tuple<string, List<string>>> PartitionDemShit(string[] lines)
        {
            dynamic slide = null;
            var partitions = new List<Tuple<string, List<string>>>();

            List<string> partition = null;
            foreach (var line in lines)
            {
                if (line.StartsWith("["))
                {
                    partition = new List<string>();
                    partitions.Add(new Tuple<string, List<string>>(SectionName(line),partition));
                    continue;
                }
                if (!string.IsNullOrWhiteSpace(line))
                {
                    partition.Add(line);
                }
            }
            return partitions;
        }

        private string SectionName(string line)
        {
            var type = line.Replace('[', ' ').Replace(']', ' ').Trim();
            return _sectionmap[type];
        }
    }       
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SplitPDF
{
    class DSAProject
    {
        public string Product { get; set; }
        public string Indication { get; set; }
        public string Segment { get; set; }
        public string Country { get; set; }
        public string Language { get; set; }
        public string Source { get; set; }
        public string Campaign { get; set; }
        public string Season { get; set; }
        public string Audience { get; set; }
        public string Notes { get; set; }
        public Dictionary<string, Presentation> Presentations { get; set; }


        public DSAProject()
        {
            Presentations = new Dictionary<string, Presentation>();
        }

    }

    class Presentation
    {
        public string PresentationName { get; set; }
        public int PresentationIndex { get; set; }
        public string Hidden { get; set; }
        public Dictionary<string, Slide> Slides { get; set; }
        public DSAProject project;
        public string PresentationID()
        {
            return project.Indication + project.Product + project.Segment + project.Country + Hidden + project.Campaign + project.Season + project.Source + PresentationIndex;
        }

    }



}

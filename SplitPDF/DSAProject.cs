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
            return project.Indication + "_" + project.Product + "_" + project.Segment + "_" + project.Country + "_" + Hidden + "_" + project.Campaign + "_" + project.Season + "_" + project.Source + "_" + PresentationIndex;
        }

        public string ExternalPresentationName()
        {
            return project.Indication + " " +  project.Season;
        }
        //Used By Slides
        public string PresentationPrefix()
        {
            return project.Indication + "_" + project.Product + "_" + project.Segment + "_" + project.Country + "_" + project.Language;
        }
    }
}

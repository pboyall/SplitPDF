using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SplitPDF
{
    class Slide
    {
//Metadata
        public string PageReference { get; set; }
        public int PageOrder{ get; set; }
        public string Title{ get; set; }
        public string Text{ get; set; }
        public string PageType{ get; set; }
        public string[] ChapterDetails{ get; set; }
        public string[] Comments{ get; set; }
        public string[] Owner{ get; set; }
        public string[] CommentsDate{ get; set; }
        public string Thumbnail{ get; set; }
        public int[] pdfPages{ get; set; }
        public string description{ get; set; }
        public string Source{ get; set; }
        public string Target{ get; set; }
        public int NavWeight{ get; set; }
        public string NavigationType { get; set; }
        public string thumbname{ get; set; }
        public string NavDesc{ get; set; }
        public string URL{ get; set; }
        public int PDFPageNo{ get; set; }
    }
}

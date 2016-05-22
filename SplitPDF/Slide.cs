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
        internal int SlideRef { get; set; }             //Unique ID so we can track slides across PDF versions
        public string PageReference { get; set; }       //English Reference from Bookmark
        public int PageLevel { get; set; }              //Bookmark Level
        public int PageOrder{ get; set; }               //Sequence Number (initially we will just use PDF Page numbers)
        public int PageNumber { get; set; }               //Not really needed but saves querying the pdfPages dictionary
        public string Title{ get; set; }                //Title from text of PDF
        public string Text{ get; set; }                 //PDF Text
        public string PageType{ get; set; }             //Type of page (Main, reference, pop up, study)
        public string[] thisNav{ get; set; }            //Details of where in the Chapter/Section/SubSection hierarchy the page falls
        public Dictionary<int, Comment> Comments { get; set; }           //Annotation Comments       
        public string Thumbnail{ get; set; }            //Slide Thumbnail
        public SortedDictionary<int, string> pdfPages{ get; set; }        //The pages which make up a slide
        public string description{ get; set; }              //English Description of the slide
        public SlideNavigation navTable { get; set; }

        //Navigation Dictionary (aim is to replace Nav Table above with more fluid structure)
        Dictionary<string, SlideNavigation> NavLinks { get; set; }      //Other slides to which this slide links

        public Slide()
        {
            Comments = new Dictionary<int, Comment>();
            pdfPages = new SortedDictionary<int, string>();
        }

    }

    class Comment
    {
        public string Text { get; set; }
        public string Owner { get; set; }
        public DateTime CommentDate { get; set; }
    }

    //Slide may navigate to many other slides; but this serves double duty as the old NavTable structure hence loads of superfluous variables
    class SlideNavigation {

        public string Source { get; set; } //Optional, cos we know the source, it's this slide!
        public string Target { get; set; }
        public string NavigationType { get; set; }          //TODO change to enum.  PopUp, Reference, NewSlide ... 

        public int NavWeight { get; set; }      
        public string thumbname { get; set; }           //Optional, tis the same as the Thumbnail
        public string NavDesc { get; set; }             //
        public string URL { get; set; }             
        public string PDFPageNo { get; set; }              //Optional, same as the pagenumber above

    }


}

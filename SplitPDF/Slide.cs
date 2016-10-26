using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SplitPDF
{
    class Slide:ICloneable
    {


        //Metadata
        //Reference Image - Thumbnail
        //Slide Number - SlideNumber 
        //Display Order - PageOrder
        //Source - Source
        //Key Message Name - PageReference
        //Description - PageReference
        //Product Message Category
        //External ID
        //Additional Notes

        //Paul's list
        //Presentation Name - from parent presentation
        //Status
        //Position - PageOrder
        //Visible Name - Page Reference
        //Concatengy Key Message Name - ExternalID



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
        public SortedDictionary<float, string> pdfPages{ get; set; }        //The pages which make up a slide
        public string description{ get; set; }              //English Description of the slide
        public SlideNavigation navTable { get; set; }
        
        //*** Fields needed to tie up with Metadata
        public decimal SlideNumber {
            get { return PageOrder; }//Read only field - no point setting it at this, er, point
        }              
        public string Source { get; set; }              //To tie up with Metadata
        public string ProductMessageCategory { get; set; }              //To tie up with Metadata, should be an enum from somewhere
        public string ExternalID { get
            {
                return presentation.PresentationPrefix().ToString() +  "_" + PageReference + "_" + Source;
            }
        }

        //*** Fields needed to tie up with Paul's sheet
        public Presentation presentation { get; set; }
        public string status { get; set; }


        //Navigation Dictionary (aim is to replace Nav Table above with more fluid structure)
        public Dictionary<string, SlideNavigation> NavLinks { get; set; }      //Other slides to which this slide links
        //Maybe add a "Parent Slide" property [could be covered by navTable Source] or ChildSlides Dictionary?


        public Slide()
        {
            Comments = new Dictionary<int, Comment>();
            pdfPages = new SortedDictionary<float, string>();
            thisNav = new string[splitPDF.maxLevels];
            navTable = new SlideNavigation();
            Source = "LO"; //Always is for us
        }

       public object Clone()
        {
            return new Slide
            {
                SlideRef = this.SlideRef,
                PageReference = this.PageReference,
                PageLevel = this.PageLevel,
                PageOrder = this.PageOrder,
                PageNumber = this.PageNumber,
                Title = this.Title,
                Text = this.Text,
                PageType = this.PageType,
                thisNav = this.thisNav,
                Comments = new Dictionary<int, Comment>(Comments),
                Thumbnail = this.Thumbnail,
                pdfPages = new SortedDictionary<float, string>(this.pdfPages),
                description = this.description,
                navTable = (SlideNavigation)this.navTable.Clone(),
                Source = this.Source,
                ProductMessageCategory = this.ProductMessageCategory,
                presentation = this.presentation
            };
        }
    }

    class Comment : ICloneable
    {
        public string Text { get; set; }
        public string Owner { get; set; }
        public DateTime CommentDate { get; set; }
        public int pagenumber;

        public object Clone()
        {
            return new Comment
            {
                Text = this.Text,
                Owner = this.Owner,
                CommentDate = this.CommentDate,
                pagenumber = this.pagenumber
            };
        }
    }

    //Slide may navigate to many other slides; but this serves double duty as the old NavTable structure hence loads of superfluous variables
    class SlideNavigation : ICloneable
    {

        public string Source { get; set; } //Optional, cos we know the source, it's this slide!
        public string Target { get; set; }
        public string NavigationType { get; set; }          //TODO change to enum.  PopUp, Reference, NewSlide ... 

        public int NavWeight { get; set; }      
        public string thumbname { get; set; }           //Optional, tis the same as the Thumbnail
        public string NavDesc { get; set; }             //
        public string URL { get; set; }             
        public string PDFPageNo { get; set; }              //Optional, same as the pagenumber above

        public object Clone()
        {
            return new SlideNavigation
            {
                Source = this.Source,
                Target = this.Target,
                NavigationType = this.NavigationType,
                NavWeight = this.NavWeight,
                thumbname = this.thumbname,
                NavDesc = this.NavDesc,
                URL = this.URL,
                PDFPageNo = this.PDFPageNo
            };
        }

    }

    //Might need a generic "Slide Sub Child" class and then inherit.  Dunno.

    class SlideReference : ICloneable
    {
        internal int SlideRef { get; set; }
        public string Text { get; set; }

        public object Clone()
        {
            return new SlideReference
            {
                SlideRef = this.SlideRef,
                Text = this.Text
            };
        }

    }



}

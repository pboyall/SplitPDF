using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Drawing.Imaging;
using System.Data;
using iTextSharp.text.pdf.parser;
using System.Xml.Linq;
using System.Globalization;
using ImageMagick;
using System.Web;
using System.Xml;
using System.Reflection;

namespace SplitPDF
{

    public enum gitlabupload
    {
        New,
        Update,
        False
    };

    public class splitPDF
    {
        internal static int maxLevels = 6;



        //Auto property for parameters to control the creation process
        public string ExcelFile { get; set; }
        public string TabName { get; set; }
        public string outputfile { get; set; }
        public string inputfile { get; set; }
        public string comparisonfile { get; set; }

        public Boolean createPDFs { get; set; }         //set to true to create individual PDFs
        public Boolean createThumbs { get; set; }         //set to true to create thumbnails
        public Boolean consolidatePages { get; set; }         //set to true to consolidate "RunOn" PDFs
        public Boolean exportMeta { get; set; }         //set to true to dump metadata sheets out
        public Boolean extractText { get; set; }         //set to true to pull text from PDF
        public Boolean exportNav { get; set; }         //set to true to do nav
        public gitlabupload exportGit{ get; set; }         //set to New to upload new issues into gitlab, update to only update, no to not connect to gitlab


        public PDFRenderer renderer { get; set; }
        internal DSAProject thisproject { get; set; }
        internal Presentation thispresentation { get; set; }

        //Configuration values
        private string[] nvType = new string[3];             //TODO Magic numbers for now
        private string[] NavLevel = new string[maxLevels];

        //Manage the flow of data.  Remove the Datatables later.
        private Dictionary<string, Dictionary<string, object>> bookmarksDict;           //i.e. Chapters //Was <string, object> before having to cater to levels and abusing dictionary
        private string[] currentNav = new string[maxLevels];
        private DataTable metatable;            //28b metadata
        private  DataTable navtable;
        private DataTable metadatatable;        //AbbVie metadata
        //Dictionary<string, string> metamapping = new Dictionary<string, string>();

        public splitPDF()
        {
            //TODO Mapping from number to text - sort this to be from config some other time
            nvType[0] = "Primary";
            nvType[1] = "Ref";
            nvType[2] = "Popup";
            NavLevel[0] = "Main";
            NavLevel[1] = "Child";
            NavLevel[2] = "SubChild";
            NavLevel[3] = "SubSubChild";
            NavLevel[4] = "Sub3Child";
            NavLevel[5] = "Sub4Child";
            renderer = new PDFRenderer();
            exportGit = gitlabupload.False;
        }

        public void newProject()
        {
            thisproject = new DSAProject();
            //return thisproject;
        }

        public void newPresentation(int loopcounter, string directory)
        {
            thispresentation = new Presentation();
            thispresentation.Hidden = "N";          //Magic
            thispresentation.project = thisproject;
            thispresentation.PresentationIndex = loopcounter;
            thispresentation.PresentationName = directory;
            thisproject.Presentations.Add(thispresentation.PresentationIndex.ToString(), thispresentation);
            //return thispresentation;
        }

        public void setProjectProperty(string key, string value)
        {
            setInstanceProperty<string>(thisproject, key, value);
        }



        //Get bookmarks, return number of bookmarks
        public int BookMarkList(string filename)
        {
            int numberofbookmarks = 0;
            string bookmarklist = "";
            PdfReader pdfReader = new PdfReader(filename);
            IList<Dictionary<string, object>> bookmarks = SimpleBookmark.GetBookmark(pdfReader);
            bookmarksDict = new Dictionary<string, Dictionary<string, object>>();//Was <string, object> before having to cater to levels and abusing dictionary
            //bookmarks will be null if no bookmarks found
            if (bookmarksDict != null)
            {
                iterateBookmarks(ref bookmarklist, bookmarks, 1);
            }
            if (bookmarks != null) { 
            numberofbookmarks = bookmarks.Count;
            }else { numberofbookmarks = 0; }
            return numberofbookmarks;
        }
        //Recursive routine to iterate the bookmarks in the passed in bookmark dictionary (bookmarks) and add them to the *global* bookmark dictionary object (bookmarksDict)
        private void iterateBookmarks(ref string bookmarklist, IList<Dictionary<string, object>> bookmarks, int level)
        {
            Console.WriteLine("Iterating Level " + level);
            if (bookmarks != null)
            {
                foreach (var bd in bookmarks)
                {
                    string bookmarkname, bookmarkpage, bookmarkpagenumber;
                    bookmarkname = ""; bookmarkpage = ""; bookmarkpagenumber = "";
                    Dictionary<string, object> bookmarkdetails = new Dictionary<string, object>();
                    if (bd.ContainsKey("Kids"))
                    {
                        //Deal with children
                        IList<Dictionary<string, object>> bdkids = (IList<Dictionary<string, object>>)bd["Kids"];
                        iterateBookmarks(ref bookmarklist, bdkids, level + 1);
                    }
                    bookmarkname = bd.Values.ToArray().GetValue(0).ToString();
                    try
                    {
                        bookmarkpage = bd["Page"].ToString();
                        bookmarkpagenumber = bookmarkpage.Substring(0, bookmarkpage.IndexOf(" "));
                        bookmarkdetails.Add("Name", bookmarkname);
                        bookmarkdetails.Add("Level", level);
                        bookmarkdetails.Add("PDFPage", bookmarkpagenumber);
                        //Check if this page exists already and add a level hyphen to it
                        bookmarksDict.Add("Page" + bookmarkpagenumber + "-" + level, bookmarkdetails);
                        bookmarklist = bookmarklist + " - " + bookmarkname;
                    }
                    catch {
                        //Don't add bookmarks which lack a page number
                    }
                }
            }
        }

        //TODO sort out a hashmap for field names
        //Just hard coded column listing, not going as far as building a full field/object mapping type system!
        private DataTable createDataTable(string tablename, Dictionary<string, string> fields)
        {
            DataTable table = new DataTable(tablename);
            if (tablename == "DSA")
            {
                //Hard coded column list for now just to test
                table.Columns.Add("PageReference", typeof(string));
                table.Columns.Add("PageOrder", typeof(int));
                table.Columns.Add("PDFPages", typeof(string));
                table.Columns.Add("Title", typeof(string));
                table.Columns.Add("Text", typeof(string));
                table.Columns.Add("PageType", typeof(string));
                table.Columns.Add("Chapter", typeof(string));
                table.Columns.Add("Section", typeof(string));
                table.Columns.Add("Subsection", typeof(string));
                table.Columns.Add("Sub-Subsection", typeof(string));
                table.Columns.Add("Sub3section", typeof(string));
                table.Columns.Add("Sub4section", typeof(string));
                table.Columns.Add("Comments", typeof(string));
                table.Columns.Add("Owner", typeof(string));
                table.Columns.Add("Thumbnail", typeof(string));
                table.Columns.Add("ParentID", typeof(string));
                table.Columns.Add("ID", typeof(string));
                table.Columns.Add("Summary", typeof(string));

                //table.Columns.Add("Thumbnail", typeof(System.Drawing.Image));

            }
            //These ones for building sitemaps - adding to main Spreadsheet but might be better to have separately
            if (tablename == "DSANav")
            {
                table.Columns.Add("Source", typeof(string));
                table.Columns.Add("Target", typeof(string));
                table.Columns.Add("Weight", typeof(string));
                table.Columns.Add("NavType", typeof(string));
                table.Columns.Add("Thumbnail", typeof(string));
                table.Columns.Add("URL", typeof(string));
                table.Columns.Add("Description", typeof(string));
                table.Columns.Add("PDFPage", typeof(string));
            }
            //For AbbVie metadata
            if (tablename == "Metadata")
            {
                table.Columns.Add("ReferenceImage", typeof(string));
                table.Columns.Add("SlideNumber", typeof(string));
                table.Columns.Add("DisplayOrder", typeof(string));
                table.Columns.Add("Source", typeof(string));
                table.Columns.Add("KeyMessageName", typeof(string));
                table.Columns.Add("Description", typeof(string));
                table.Columns.Add("ProductMessageCategory", typeof(string));
                table.Columns.Add("ExternalID", typeof(string));
                table.Columns.Add("AdditionalNotes", typeof(string));
            }

            return table;



        }

        private void createSplitPDF(int pagenumber, PdfReader reader)
        {
            //Create Single PDF for this page only
            if (createPDFs)
            {
                string filename = pagenumber.ToString() + ".pdf";
                Document document = new Document();
                PdfCopy copy = new PdfCopy(document, new FileStream(outputfile + "\\" + filename, FileMode.Create));
                document.Open();
                copy.AddPage(copy.GetImportedPage(reader, pagenumber));
                document.Close();
            }
        }

        private void getAnnotations(PdfReader reader, int pagenumber, ref Slide thisSlide)
        {
            PdfDictionary page = reader.GetPageN(pagenumber);
            PdfArray annots = page.GetAsArray(PdfName.ANNOTS);
            int commentCounter = 1;
            if (annots != null)
            {
                foreach (PdfObject annot in annots.ArrayList)
                {
                    PdfDictionary annotation = (PdfDictionary)PdfReader.GetPdfObject(annot);
                    PdfName subType = (PdfName)annotation.Get(PdfName.SUBTYPE);
                    if (PdfName.TEXT.Equals(subType) || PdfName.HIGHLIGHT.Equals(subType) || PdfName.INK.Equals(subType) || PdfName.FREETEXT.Equals(subType))
                    {
                        PdfString title = annotation.GetAsString(PdfName.T);            //Seems to store author
                        //PdfString contents = annotation.GetAsString(PdfName.CONTENTS);   //Visible Text
                        //string newcontents = annotation.GetAsString(PdfName.CONTENTS).ToString();   //Visible Text
                        var comment = annotation.GetAsString(PdfName.CONTENTS) ?? new PdfString("") ;
                        string commentext = comment.ToUnicodeString();   //Visible Text
                        commentext = System.Security.SecurityElement.Escape(commentext);
                        //pageComments = pageComments + contents.ToString() + "\r\n";
                        //pgOwner = pgOwner + title.ToString() + "\r\n";
                        Comment thisComment = new Comment();
                        thisComment.CommentDate = DateTime.Now;         //Not technically correct, but will suffice
                        thisComment.pagenumber = pagenumber;
                        thisComment.Owner = title?.ToString() ?? "";
                        thisComment.Text = commentext;
                        thisSlide.Comments.Add(commentCounter, thisComment);
                        commentCounter++;

                        //Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes

                    }
                }
            }
        }

        private string ManageBookmarks(int pagenumber, ref Slide thisSlide, ref SortedDictionary<float, Slide> Slides, out int subpage)
        {
            int BookmarkLevel;
            string retval = "";
            Slide prevSlide = (Slide)thisSlide.Clone();            //Store Transiently to cater for the new bookmark loop where we reset at the end of each subbookmark that exists on the same page
            subpage = 0;
            if (String.IsNullOrEmpty(currentNav[0])) { currentNav[0] = "Entry"; }                        //Dummy Node to start things off
                                                                                                         //Check to see Bookmark for this page unless it's a run on page in which case we ignore it
            Dictionary<string, Dictionary<string, object>> bookmarkdic = new Dictionary<string, Dictionary<string, object>>();
            bookmarkdic = bookmarksDict.Where(d => d.Key.Contains("Page" + pagenumber.ToString() + "-")).ToDictionary(d => d.Key, d => d.Value);
            int numpages = bookmarkdic.Count;
            int loopcounter = 1;
            if (numpages > 0)
            //            if (bookmarksDict.ContainsKey(pagenumber.ToString()))
            {
                //object wibble;
                //Don't forget arrays start at 0, level count starts at 1
                //if (bookmarksDict.TryGetValue("Page" + pagenumber.ToString(), out wibble))
                foreach(var wibble in bookmarkdic)
                {
                    Dictionary<string, object> wibble1 = wibble.Value;
                    //Dictionary<string, object> wibble1 = (Dictionary<string, object>)wibble;
                    thisSlide.PageReference = wibble1["Name"].ToString();
                    BookmarkLevel = int.Parse(wibble1["Level"].ToString());
                    thisSlide.navTable.PDFPageNo = wibble1["PDFPage"].ToString();
                    int NavSub = (BookmarkLevel - 2) < 0 ? 0 : (BookmarkLevel - 2);              //Only go back 2 if more than 2 to start with (i.e. BookmarkLevel 1 can only go back to 0)
                    currentNav.CopyTo(thisSlide.thisNav, 0);                                    //Do not just assign
                    thisSlide.navTable = new SlideNavigation();
                    thisSlide.navTable.NavDesc = "Title:" + thisSlide.PageReference + ",Level:" + BookmarkLevel;
                    thisSlide.navTable.Source = currentNav[NavSub];           //Levels 1 and 2 go back to level 0, all others go back to previous level
                    thisSlide.thisNav[BookmarkLevel - 1] = thisSlide.PageReference;
                    for (int i = BookmarkLevel; i < maxLevels; i++)
                    {
                        thisSlide.thisNav[i] = "";//Clear out Children at each level below this one, so children retain the parent that would have been set earlier but parents get cleaned out
                    }
                    thisSlide.navTable.NavWeight= 100 / BookmarkLevel;                    //TODO sort magic number
                    thisSlide.navTable.Target= thisSlide.thisNav[BookmarkLevel - 1];              //Don't forget arrays start at 0
                    thisSlide.navTable.NavigationType = NavLevel[BookmarkLevel - 1];
//This will grow more complicated when we have thought more about it
                    if (thisSlide.PageType.Contains("Reference")) { thisSlide.navTable.NavigationType = nvType[1] + thisSlide.navTable.NavigationType; }
                    thisSlide.PageLevel = BookmarkLevel;

                    //Set the "olds" (redo as function)
                    for (int i = 0; i < maxLevels; i++)
                    {
                        currentNav[i] = thisSlide.thisNav[i];//Copy over for next loop
                    }
                    //Moved from outside loop to cater for multi-bookmark pages.  Only do this if there are more bookmarks than pages (i.e. 3 bookmarks on 1 page.  1 bookmark on this page doesn't trigger the if statement as 1 =1 so the Slides.Add is done outside.  Messy.
                    if (loopcounter < numpages)
                    { 
                        //if (loopcounter == 1){
                            thisSlide.PageOrder = pagenumber;                           //Just using pageNumber
                            thisSlide.pdfPages.Add(pagenumber, "Only Page");
                        //}
                        Slides.Add(pagenumber + ((float)loopcounter / 10), thisSlide);
                        //Reset
                        thisSlide = null;
                        thisSlide = (Slide)prevSlide.Clone();
                        loopcounter++;
                    }
                    subpage = loopcounter;//    Passing this out means that the final page gets numbered as the top page (e.g. 1.1. and 1.2 are done here and 1.3 is done outside, otherwise the outside page would have been numbered just "1")
                    retval = "New";
                }
            }
            else
            {
                //No Bookmark for this page so take it as part of the previous page
                retval = "RunOn";
            }
            return retval;
        }

        private void ExtractText(int pagenumber, PdfReader reader, ref Slide thisSlide)
        {
            var pdfText = "";
            if (extractText)
            {
                Rectangle mediabox = reader.GetPageSize(pagenumber);
                //Extract Text from the page.  Have to reinitialise Text Extraction Strategy each time as otherwise you end up with all the text from the PDf - weird
                PDFExtractor pdfExtract = new PDFExtractor();

                
                pdfText = PdfTextExtractor.GetTextFromPage(reader, pagenumber, pdfExtract.bodystrategy);
                thisSlide.Text = Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(pdfText)));
                //Extract title text from page (some duplicated code here, leave it for now)  Note that all these declarations need stuff only known inside the loop
                pdfText = PdfTextExtractor.GetTextFromPage(reader, pagenumber, pdfExtract.titlestrategy);
                thisSlide.Title = Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(pdfText)));
            }else
            {
                thisSlide.Title = "";
                thisSlide.Text = "";
            }
        }

        //Chop the input PDF into separate pages
        public int Split()
        {
            int pageCount;
            renderer.outputfile = outputfile;
            renderer.inputfile = inputfile;
            thispresentation.project = thisproject;
            #region SetupTable
            metatable = createDataTable("DSA", new Dictionary<string, string>());
            navtable = createDataTable("DSANav", new Dictionary<string, string>());
            metadatatable = createDataTable("Metadata", new Dictionary<string, string>());
            #endregion
            using (PdfReader reader = new PdfReader(inputfile))
            {
                pageCount = reader.NumberOfPages;
                try { TabName = reader.Info["Title"]; } catch (Exception e) { }//Just consume an error, seems hit and miss whether the PDF gives back a title or not
                                                                               //Iterate around the PDF, keep these outside loop so they propogate downwards
                Slide thisSlide;
                Slide oldSlide = new Slide();
                string id = "", parentid = "";
                SortedDictionary<float, Slide> Slides = new SortedDictionary<float, Slide>();   //Collection of all the Slides found in this presentation

                for (int pagenumber = 1; pagenumber <= reader.NumberOfPages; pagenumber++)
                {
                    string bookmarkreturn = "";
                    string thumbname = System.IO.Path.Combine(outputfile, "Thumb" + System.IO.Path.GetFileNameWithoutExtension(inputfile) + "-p" + pagenumber + ".png");
                    int subpage;            //Caters for pages with more than one bookmark
                    createSplitPDF(pagenumber, reader);
                    thisSlide = new Slide();
                    thisSlide.presentation = thispresentation;
                    thisSlide.PageNumber = pagenumber;
                    ExtractText(pagenumber, reader, ref thisSlide);
                    //Can do a comparison here but not really necessary as checking for bookmarks instead if (thisSlide.Text == oldSlide.Text)
                    thisSlide.PageType = getPageType(thisSlide);

                    getAnnotations(reader, pagenumber, ref thisSlide);          //Adds to thisSlide.Annotations.  Passing by ref - clunky but works for now.  
                    if (createThumbs) thisSlide.Thumbnail = renderer.GenerateThumbnail(pagenumber, thumbname); else  thisSlide.Thumbnail = thumbname;

                    #region CheckBookmarks
                    BookMarkList(inputfile);//Populates the dictionary, no return
                    bookmarkreturn = ManageBookmarks(pagenumber, ref thisSlide, ref Slides, out subpage);       //Adds all the Nav + PageReference + PageLevel
                    
                    if (bookmarkreturn != "New") {
                        copyOldValues(ref thisSlide, oldSlide);//Recover the old values
                        thisSlide.pdfPages.Add(pagenumber, "Non Bookmarked Page, presumed to be a run on");
                        thisSlide.PageType += "Multiple";
                        if (consolidatePages) Slides.Remove(pagenumber - 1); else thisSlide.navTable.NavWeight = 0;
                    }
                    else
                    {
                        thisSlide.PageOrder = pagenumber;                           //Just using pageNumber
                        thisSlide.pdfPages.Add(pagenumber , "Start Page");
                        //renderer.GenerateThumbnail(pagenumber, System.IO.Path.Combine(outputfile, thisSlide.PageReference + ".png"));
                    }
                    #endregion
                    //Add to Slides Collection
                    Slides.Add(pagenumber, thisSlide);

                    //Do the olds
                    oldSlide = thisSlide;
                    thisSlide = null;
                }
                thisSlide = null;
                bookmarksDict = null;
                int SlideCounter = 1;       //Required for JIRA Import to Generate ID columns
                //Add to datatables, don't add empty rows (where a slides runs across multiple PDF pages) to NavTable
                foreach(var SlideDict in Slides.OrderBy(ts=>ts.Key)) {
                    Slide theSlide = SlideDict.Value;
                    String pageComments = "", pageOwner = "";                    //Cater for multiple comments on a single page
                    int commentCounter = 1;
                    foreach (var CommentPair in theSlide.Comments)
                    {
                        Comment slideComment = CommentPair.Value;
                        pageComments = pageComments + " (" + commentCounter + "), page " +  slideComment.pagenumber + ":" + slideComment.Text;
                        pageOwner = pageOwner + " (" + commentCounter + "), " + slideComment.Owner;
                        commentCounter++;
                    }
                    string pdfPages = theSlide.pdfPages.Keys.Min().ToString() + "-" + theSlide.pdfPages.Keys.Max().ToString();
                    if (theSlide.navTable.NavWeight < 100)
                    {
                        parentid = id;
                        id = null;
                    }
                    else { }
                    id = SlideCounter.ToString();

                    //EncodeStrings(ref theSlide);
                    //pageComments = EncodeString(pageComments);

                    metatable.Rows.Add(theSlide.PageReference, theSlide.PageNumber, pdfPages, 
                        theSlide.Title,theSlide.Text, 
                        theSlide.PageType, theSlide.thisNav[0], theSlide.thisNav[1], theSlide.thisNav[2], theSlide.thisNav[3], theSlide.thisNav[4], 
                        theSlide.thisNav[5],
                        pageComments, pageOwner, theSlide.Thumbnail, id, parentid, theSlide.PageReference + pdfPages);

                    if (theSlide.navTable.NavWeight != 0) { navtable.Rows.Add(theSlide.navTable.Source, theSlide.navTable.Target, theSlide.navTable.NavWeight, theSlide.navTable.NavigationType, theSlide.Thumbnail, "", theSlide.navTable.NavDesc, theSlide.navTable.PDFPageNo); }
                    metadatatable.Rows.Add(theSlide.Thumbnail, theSlide.PageNumber, theSlide.PageOrder, theSlide.Source, theSlide.PageReference, theSlide.description, theSlide.ProductMessageCategory, theSlide.ExternalID, theSlide.Comments );



                    SlideCounter++;
                }
                return pageCount;
            }
        }


        private string EncodeString(string theString)
        {
            //return theString = new System.Xml.Linq.XText(theString).ToString();

            return XmlConvert.EncodeName(theString);
            // HttpUtility.HtmlEncode(xml);

        }

        private string EncodeStringXML(string theString)
        {
            XmlDocument doc = new XmlDocument();
            XmlNode node = doc.CreateElement("root");
            node.InnerText = theString;
            return node.InnerXml;
        }

        private void EncodeStrings(ref Slide theSlide)
        {
            theSlide.Title = Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(theSlide.Title)));
            theSlide.Text = Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(theSlide.Text)));
            theSlide.Title = EncodeString(theSlide.Title);
            theSlide.Text = EncodeString(theSlide.Text);
        }

        private static System.Drawing.Image GetImageFromFile(string fileName)
        {
            //check the existence of the file in disc
            if (File.Exists(fileName))
            {
                System.Drawing.Image image = System.Drawing.Image.FromFile(fileName);
                return image;
            }
            else
                return null;
        }

        private void copyOldValues(ref Slide thisSlide, Slide oldSlide)
        {
            //Duplicate relevant values from previous slide, I don't feel we should have rows for each page of the PDF, but doing that for now as Excel Base
            //Probably should write a "copyoldvalue" function instead of having each row separate like this - but that would involve reflecting the properties of the object, hassle and would not run under .net Core
            // Leaving Slide Ref out for now as not yet using it (everything still Excel Based)
            thisSlide.PageReference = oldSlide.PageReference;
            thisSlide.PageLevel = oldSlide.PageLevel;
            thisSlide.PageOrder = oldSlide.PageOrder;
            if (consolidatePages) {
                thisSlide.PageNumber = oldSlide.PageNumber;                    //Consolidate then it would be just the start number
                if ((thisSlide.Text != oldSlide.Text))//Title and Text are set by ExtractText and for Run-On pages should be the same, if they are different and consolidating need to merge them
                {
                    thisSlide.Text = oldSlide.Text + thisSlide.Text;                //Don't treat as run on?  Write out new row? Or Merge with existing? If Consolidating, merge them.  Otherwise leave separate
                }
                if ((thisSlide.Title != oldSlide.Title))//Title and Text are set by ExtractText and for Run-On pages should be the same, if they are different and consolidating need to merge them
                {
                    thisSlide.Title= oldSlide.Title + thisSlide.Title;                //Don't treat as run on?  Write out new row? Or Merge with existing? If Consolidating, merge them.  Otherwise leave separate
                }
                copyAnnotations(ref thisSlide, oldSlide);                       //Comments are done outside - they really are per page.  Consolidate using copyAnnotations function written
                thisSlide.Thumbnail = oldSlide.Thumbnail;
            }
            //PageType is set outside here
            //Set the "olds" (redo as function) - .net does arrays by reference so if you just do thisNav = oldNav then you only every have one nav array!
            oldSlide.thisNav.CopyTo(thisSlide.thisNav, 0);
            thisSlide.pdfPages = new SortedDictionary<float, string>(oldSlide.pdfPages);         
            //Description has no value set currently - was to get a column for humans to write to in the spreadsheet
            thisSlide.navTable = (SlideNavigation)oldSlide.navTable.Clone();         
            //Navlinks not used yet
            //thisSlide.NavLinks = new Dictionary<string, SlideNavigation>(oldSlide.NavLinks);         
        }

        private void copyAnnotations(ref Slide thisSlide,  Slide oldSlide)
        {
            int oldSlideComments = oldSlide.Comments.Count;
            //Insert oldSlide's values into dictionary then add current Slide's comments (so retain order for the keys)
            Dictionary<int, Comment> newComments = new Dictionary<int, Comment>();
            foreach (var note in oldSlide.Comments)
            {
                newComments.Add(note.Key, note.Value);
                //thisSlide.Comments.Add(note.Key + thisSlide.Comments.Count, note.Value);
            }
            foreach (var note in thisSlide.Comments)
            {
                newComments.Add(note.Key + oldSlideComments, note.Value);
            }
            thisSlide.Comments = new Dictionary<int, Comment>(newComments);
        }

        private string getPageType(Slide thisSlide)
        {
            //Have a stab at working out page type
            //Based on the contents of the page we can identify reference pop ups (they have the same text as their parent with the additional layer of reference X)
            //Regular Expressions give me a headache
            if (thisSlide.Text.Contains("Reference"))
            {
                return "Reference";            //Really need an enum for these, but until we know what we're doing, text will work
            }
            else
            {
                return "Main";
            }


        }

        //Do a differences comparison to see where things have changed
        public Boolean compareToPrevious(string previousfile, string outputfile)
        {
            return true;
        }
 
        #region ExtractBackgroundImages
        //This extracts all the JPGs from the file (i.e. it will grab the screens without comments)
        private int extractJPGs()
        {
            FileInfo file = new FileInfo(inputfile);
            string name = file.Name.Substring(0, file.Name.LastIndexOf("."));
            PdfSharp.Pdf.PdfDocument document = PdfSharp.Pdf.IO.PdfReader.Open(inputfile);
            int imageCount = 0;
            // Iterate pages
            foreach (PdfSharp.Pdf.PdfPage page in document.Pages)
            {
                // Get resources dictionary
                PdfSharp.Pdf.PdfDictionary resources = page.Elements.GetDictionary("/Resources");
                if (resources != null)
                {
                    // Get external objects dictionary
                    PdfSharp.Pdf.PdfDictionary xObjects = resources.Elements.GetDictionary("/XObject");
                    if (xObjects != null)
                    {
                        ICollection<PdfSharp.Pdf.PdfItem> items = xObjects.Elements.Values;
                        // Iterate references to external objects
                        foreach (PdfSharp.Pdf.PdfItem item in items)
                        {
                            PdfSharp.Pdf.Advanced.PdfReference reference = item as PdfSharp.Pdf.Advanced.PdfReference;
                            if (reference != null)
                            {
                                PdfSharp.Pdf.PdfDictionary xObject = reference.Value as PdfSharp.Pdf.PdfDictionary;
                                // Is external object an image?
                                if (xObject != null && xObject.Elements.GetString("/Subtype") == "/Image")
                                {
                                    ExportImage(xObject, ref imageCount);
                                }
                            }
                        }
                    }
                }
            }
            return imageCount;
        }

        static void ExportImage(PdfSharp.Pdf.PdfDictionary image, ref int count)
        {
            ExportJpegImage(image, ref count);
        }

        static void ExportJpegImage(PdfSharp.Pdf.PdfDictionary image, ref int count)
        {
            // Fortunately JPEG has native support in PDF and exporting an image is just writing the stream to a file.
            byte[] stream = image.Stream.Value;
            FileStream fs = new FileStream(String.Format("Image{0}.jpeg", count++), FileMode.Create, FileAccess.Write);
            BinaryWriter bw = new BinaryWriter(fs);
            bw.Write(stream);
            bw.Close();
        }

        #endregion

        
        public void ExportToExcel(string excelfile, string Tabname, string type)
        {
            DataTable table;
            //Name Tab by Date? By definition right now this will always be a new file
            switch (type.ToUpper()){
                case "META":  table = this.metatable;
                    break;
                case "NAV": table = this.navtable;
                    break;
                default: table = this.metatable;
                    break;
            }
            ExcelExport ee = new ExcelExport();
            //if (String.IsNullOrWhiteSpace(ExcelFile) == true) { ExcelFile = "./" + Guid.NewGuid().ToString() + ".xlsx"; }
            if (String.IsNullOrWhiteSpace(Tabname)) { Tabname = "Tab " + DateTime.Now.ToShortDateString().Replace('/', '-'); }
            if (type.ToUpper() == "NAV" && !exportNav ){
                    //Do nothing
                }
                else{
                try
                    {
                        ee.ExportToExcel(excelfile, Tabname, table);
                    }
                    catch (Exception e)
                    {
                        ///Issue with Spreadsheet Light and OPenxML
                    }

            }
        }

        public void ExportToGit(string projectid)
        {
            Gitlab gitter = new Gitlab();
            ExcelExport ee = new ExcelExport();
            gitter.project = projectid;
            //for each row read Thumbcol value and load data 
            int iEndRowIndex = ee.iStartRowIndex + this.metatable.Rows.Count + 1 - 1;

            for (int i = ee.iStartRowIndex; i < iEndRowIndex; i++)
            {
                string filepath;
                filepath = this.metatable.Rows[i - 1][ee.thumbCol - 1].ToString();
                string title = this.metatable.Rows[i - 1][0].ToString(); //4
                string description = "page " + this.metatable.Rows[i - 1][1].ToString() + " " + this.metatable.Rows[i - 1][0].ToString() + this.metatable.Rows[i - 1][4].ToString(); //5 and 14
<<<<<<< HEAD
                filepath = @"G:\PDFSplitting\output\arse.png";
                gitter.raiseIssue(filepath, title, description);
=======
                gitter.raiseIssue(filepath.Replace(".png", "-full.png"), title, description);
>>>>>>> c3b94f282d5697bdccb343b9fb71492791fa7c05
            }
        }

        public void ExportMetadata()
        {
            if (exportMeta)
            {
                try {
                ExcelExport ee = new ExcelExport();
                ee.ExportMetadata(thisproject, metatable);
                }catch(Exception e)
                {
                    ///Issue with Spreadsheet Light and OPenxML
                }
            }
        }

        public static void setInstanceProperty<PROPERTY_TYPE>(object instance, string propertyName, PROPERTY_TYPE value)
        {
            Type type = instance.GetType();
            PropertyInfo propertyInfo = type.GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public, null, typeof(PROPERTY_TYPE), new Type[0], null);
            propertyInfo.SetValue(instance, value, null);

            return;
        }

    }

    
}







/*if (PdfName.HIGHLIGHT.Equals(subType)) {
    //Have to highlight an area on the page to extract
    PdfArray coordinates = annotation.GetAsArray(PdfName.RECT);
    //Might have to use QuadPoint annotationDic.GetAsArray(PdfName.QUADPOINTS)
    Rectangle rect = new Rectangle(float.Parse(coordinates.ArrayList[0].ToString(), CultureInfo.InvariantCulture.NumberFormat), float.Parse(coordinates.ArrayList[1].ToString(), CultureInfo.InvariantCulture.NumberFormat),
    float.Parse(coordinates.ArrayList[2].ToString(), CultureInfo.InvariantCulture.NumberFormat), float.Parse(coordinates.ArrayList[3].ToString(), CultureInfo.InvariantCulture.NumberFormat));
    RenderFilter[] filter = { new RegionTextRenderFilter(rect) };
    strategy = new FilteredTextRenderListener(new LocationTextExtractionStrategy(), filter);
    String contents = PdfTextExtractor.GetTextFromPage(reader, pagenumber, strategy);
    pageComments = pageComments + counter + contents.ToString();
}*/


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

namespace SplitPDF
{

   public class splitPDF
    {
        //Auto property for parameters to control the creation process
        public string ExcelFile { get; set; }
        public string TabName { get; set; }
        public string outputfile { get; set; }
        public string inputfile { get; set; }
        internal PDFRenderer renderer { get; set; }
        public Boolean createPDFs { get; set; }         //set to true to create individual PDFs
        public static int maxLevels = 6;                    //TODO Magic numbers for now
        public string[] nvType = new string[3];             //TODO Magic numbers for now
        public string[] NavLevel = new string[maxLevels];

        private Dictionary<string, object> bookmarksDict;           //i.e. Chapters
        private DataTable metatable;
        private  DataTable navtable;
        
        private string[] currentNav = new string[maxLevels];

        //Per Row data
        public string[] thisNav = new string[maxLevels];
        public string NavigationType;
        public int NavWeight;
        public int BookmarkLevel;
        string Source;
        string PDFPage;
        string Target;
        string oldPageText;
        string NavDesc;
        string PDFPageNo;
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
        }

        //Get bookmarks, return number of bookmarks
        public int BookMarkList(string filename)
        {
            int numberofbookmarks = 0;
            string bookmarklist = "";
            PdfReader pdfReader = new PdfReader(filename);
            IList<Dictionary<string, object>> bookmarks = SimpleBookmark.GetBookmark(pdfReader);
            bookmarksDict = new Dictionary<string, object>();
            //bookmarks will be null if no bookmarks found
            if (bookmarksDict != null)
            {
                iterateBookmarks(ref bookmarklist, bookmarks, 1);
            }
            numberofbookmarks = bookmarks.Count;
            return numberofbookmarks;
        }
        //Recursive routine to iterate the bookmarks in the passed in bookmark dictionary (bookmarks) and add them to the *global* bookmark dictionary object (bookmarksDict)
        private void iterateBookmarks(ref string bookmarklist, IList<Dictionary<string, object>> bookmarks, int level)
        {
            Console.WriteLine("Iterating Level " + level);
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
                bookmarkpage = bd["Page"].ToString();
                bookmarkpagenumber = bookmarkpage.Substring(0, bookmarkpage.IndexOf(" "));
                bookmarkdetails.Add("Name", bookmarkname);
                bookmarkdetails.Add("Level", level);
                bookmarkdetails.Add("PDFPage", bookmarkpagenumber);
                bookmarksDict.Add(bookmarkpagenumber, bookmarkdetails);
                bookmarklist = bookmarklist + " - " + bookmarkname;
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
                table.Columns.Add("PageReference", typeof(int));
                table.Columns.Add("PageOrder", typeof(int));
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

        private string getAnnotations(PdfReader reader, int pagenumber, out string pageOwner)
        {
            PdfDictionary page = reader.GetPageN(pagenumber);
            PdfArray annots = page.GetAsArray(PdfName.ANNOTS);
            string pageComments = "", pgOwner = "";
            if (annots != null)
            {
                foreach (PdfObject annot in annots.ArrayList)
                {
                    PdfDictionary annotation = (PdfDictionary)PdfReader.GetPdfObject(annot);
                    PdfName subType = (PdfName)annotation.Get(PdfName.SUBTYPE);
                    if (PdfName.TEXT.Equals(subType) || PdfName.HIGHLIGHT.Equals(subType) || PdfName.INK.Equals(subType) || PdfName.FREETEXT.Equals(subType))
                    {
                        PdfString title = annotation.GetAsString(PdfName.T);            //Seems to store author
                        PdfString contents = annotation.GetAsString(PdfName.CONTENTS);  //Visible Text
                        pageComments = pageComments + contents.ToString() + "\r\n";
                        pgOwner = pgOwner + title.ToString() + "\r\n";
                    }
                }
            }
            pageOwner = pgOwner;
            return pageComments;
        }

        private string ManageBookmarks(int pagenumber, string pageType)
        {
            if (String.IsNullOrEmpty(currentNav[0])) { currentNav[0] = "Entry"; }                        //Dummy Node to start things off
                                                                                                         //Check to see Bookmark for this page unless it's a run on page in which case we ignore it
            if (bookmarksDict.ContainsKey(pagenumber.ToString()) && pageType != "RunOn")
            {
                object wibble;
                //Don't forget arrays start at 0, level count starts at 1
                if (bookmarksDict.TryGetValue(pagenumber.ToString(), out wibble))
                {
                    Dictionary<string, object> wibble1 = (Dictionary<string, object>)wibble;
                    string BookmarkTitle = wibble1["Name"].ToString();
                    BookmarkLevel = int.Parse(wibble1["Level"].ToString());
                    PDFPageNo = wibble1["PDFPage"].ToString();
                    NavDesc = "Title:" + BookmarkTitle + ",Level:" + BookmarkLevel;
                    int NavSub = (BookmarkLevel - 2) < 0 ? 0 : (BookmarkLevel - 2);              //Only go back 2 if more than 2 to start with (i.e. BookmarkLevel 1 can only go back to 0)
                    thisNav[BookmarkLevel - 1] = BookmarkTitle;
                    for (int i = BookmarkLevel; i < maxLevels; i++)
                    {
                        thisNav[i] = "";//Clear out Children at each level below this one, so children retain the parent that would have been set earlier but parents get cleaned out
                    }
                    NavWeight = 100 / BookmarkLevel;
                    Target = thisNav[BookmarkLevel - 1];              //Don't forget arrays start at 0
                    Source = currentNav[NavSub];           //Levels 1 and 2 go back to level 0, all others go back to previous level
                    NavigationType = NavLevel[BookmarkLevel - 1];
                    if (pageType == "Reference") { NavigationType = nvType[1]; }

                    //Set the "olds" (redo as function)
                    for (int i = 0; i < maxLevels; i++)
                    {
                        currentNav[i] = thisNav[i];//Copy over for next loop
                    }

                }
                else
                {
                    //No Bookmark for this page so just utterly ignore it
                    PDFPage = pagenumber.ToString();
                }

            }
            return "";
        }

        private string ExtractText(int pagenumber, PdfReader reader, out string TitleText)
        {
            Rectangle mediabox = reader.GetPageSize(pagenumber);
            //Extract Text from the page.  Have to reinitialise Text Extraction Strategy each time as otherwise you end up with all the text from the PDf - weird
            PDFExtractor pdfExtract = new PDFExtractor();

            var pdfText = "";
            pdfText = PdfTextExtractor.GetTextFromPage(reader, pagenumber, pdfExtract.bodystrategy);
            string pageText = Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(pdfText)));
            //Extract title text from page (some duplicated code here, leave it for now)  Note that all these declarations need stuff only known inside the loop
            pdfText = PdfTextExtractor.GetTextFromPage(reader, pagenumber, pdfExtract.titlestrategy);
            TitleText = Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(pdfText)));
            return pageText;
        }

        //Chop the input PDF into separate pages
        public int Split()
        {
            int pageCount;
            string thumbfile = System.IO.Path.GetFileNameWithoutExtension(inputfile);
            renderer.outputfile = outputfile;
            renderer.inputfile = inputfile;
            #region SetupTable
            metatable = createDataTable("DSA", new Dictionary<string, string>());
            navtable = createDataTable("DSANav", new Dictionary<string, string>());
            #endregion
            using (PdfReader reader = new PdfReader(inputfile))
            {
                pageCount = reader.NumberOfPages;
                try { TabName = reader.Info["Title"]; } catch (Exception e) { }//Just consume an error, seems hit and miss whether the PDF gives back a title or not
                                                                               //Iterate around the PDF, keep these outside loop so they propogate downwards
                NavigationType = ""; oldPageText = ""; Source = ""; Target = ""; NavDesc = ""; PDFPageNo = "";
                for (int pagenumber = 1; pagenumber <= reader.NumberOfPages; pagenumber++)
                {
                    NavWeight = 0;
                    string pageComments = "", pageOwner = "", pageText = "", pageType = "", titleText = "";
                    string thumbname = thumbfile + "-p" + pagenumber + ".jpg";

              
                    createSplitPDF(pagenumber, reader);
                    pageText = ExtractText(pagenumber, reader, out titleText);
                    if (pageText == oldPageText) { pageType = "RunOn"; }     //Page is one of those where a single slide has more content than will fit on one PDF Page.  Is there a better way to check this?

                    #region pageType
                    //Have a stab at working out page type
                    //Based on the contents of the page we can identify reference pop ups (they have the same text as their parent with the additional layer of reference X)
                    //Regular Expressions give me a headache
                    if (pageText.Contains("Reference"))
                    {
                        pageType = "Reference";             //Really need an enum for these, but until we know what we're doing, text will work
                    }
                    #endregion
                    pageComments = getAnnotations(reader, pagenumber, out pageOwner);
                    #region CheckBookmarks

                    BookMarkList(inputfile);//Populates the dictionary, no return
                    ManageBookmarks(pagenumber, pageType);

                    if (!bookmarksDict.ContainsKey(pagenumber.ToString()) || pageType == "RunOn")
                    {
                        NavWeight = 0; //Clear out targeting for sitemap as this page is not a separate page
                    }
                    #endregion
                    string imagefile = renderer.GenerateThumbnail(pagenumber);
                    //Add to datatable
                    
                    metatable.Rows.Add(pagenumber, pagenumber, titleText, pageText, pageType, thisNav[0], thisNav[1], thisNav[2], thisNav[3], thisNav[4], thisNav[5], pageComments, pageOwner, imagefile);
                    //Don't add empty rows (where a slides runs across multiple PDF pages)
                    if (NavWeight != 0) { navtable.Rows.Add(Source, Target, NavWeight, NavigationType, thumbname, "", NavDesc, PDFPageNo); }
                    oldPageText = pageText;
                }
                return pageCount;
            }

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

        internal void ExportToExcel(string excelfile, string Tabname, string type)
        {
            DataTable table;
            //Name Tab by Date? By definition right now this will always be a new file
            switch (type){
                case "meta":  table = this.metatable;
                    break;
                case "nav": table = this.navtable;
                    break;
                default: table = this.metatable;
                    break;
            }
            ExcelExport ee = new ExcelExport();
            //if (String.IsNullOrWhiteSpace(ExcelFile) == true) { ExcelFile = "./" + Guid.NewGuid().ToString() + ".xlsx"; }
            if (String.IsNullOrWhiteSpace(Tabname)) { Tabname = "Tab " + DateTime.Now.ToShortDateString().Replace('/', '-'); }
            ee.ExportToExcel(excelfile, Tabname, table);
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


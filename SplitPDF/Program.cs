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
    class Program
    {

         static void Main(string[] args)
        {
            splitPDF splitter = new splitPDF();
            //Path.GetTempPath()
            string mydirectory = System.IO.Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]).ToString();
            splitter.inputfile = mydirectory + "\\BookmarkTesting.pdf";
            splitter.outputfile = mydirectory + "\\Output";     //Just testing
            splitter.exportDPI = 300;
            splitter.thumbnailheight = 150;
            splitter.thumbnailwidth = 200;
            splitter.createPDFs = true;
            //splitter.BookMarkList(splitter.inputfile);
            //Save each PDF as a separate PDF page
            int returned = splitter.Split();
            //Create a row in a spreadsheet for each PDF
            string excelfile = splitter.outputfile + "\\" + Guid.NewGuid().ToString() + ".xlsx";
            splitter.ExportToExcel(excelfile, "Meta", splitter.metatable);     //No tabname for now - that would be if updating.  Later
            splitter.ExportToExcel(excelfile, "Nav", splitter.navtable);     //No tabname for now - that would be if updating.  Later

            /*
            splitter.readSiteMap();
            */
        }


        //Separate Class as Originally I hoped to only have to initialise all the expensive strategies and filters once - turns out you can't do that
        private class PDFExtractor{
            //Dimensions for the box on the page where the Title Text is stored  (change to struct later) Not sure how to work out what these dimensions should be, don't like the idea of trial and error!
            public float distanceInPixelsFromLeft = 174;
            public float distanceInPixelsFromBottom = 1950;
            public float width = 1000;
            public float height = 200 ;
            public ITextExtractionStrategy bodystrategy { get; set; }
            public ITextExtractionStrategy titlestrategy { get; set; }

            public PDFExtractor() {
                bodystrategy = new SimpleTextExtractionStrategy();
                var filters = new RenderFilter[1];
                var titlerect = new System.util.RectangleJ(distanceInPixelsFromLeft, distanceInPixelsFromBottom, width, height);
                filters[0] = new RegionTextRenderFilter(titlerect);
                titlestrategy = new FilteredTextRenderListener(new LocationTextExtractionStrategy(), filters);
            }
        }


        public class splitPDF
        {

            public string ExcelFile { get; set; }
            public string TabName { get; set; }
            public string outputfile { get; set; }
            public string inputfile { get; set; }
            public string xmlfile { get; set; }
            public int exportDPI { get; set; }
            public int thumbnailheight { get; set; }
            public int thumbnailwidth { get; set; }
            public Boolean createPDFs { get; set; }         //set to true to create individual PDFs

            //This perhaps isn't ideal, as in a perfect world we'd have a multi dimensional dictionary of dictionaries that could expand indefinitely.  
            //However, it's easier to understand this way
            public Dictionary<string, object> bookmarksDict;           //i.e. Chapters
            public DataTable metatable;
            public DataTable navtable;
            //Get bookmarks, return number of bookmarks
            public int BookMarkList(string filename)
            {
                int numberofbookmarks = 0;
                string bookmarklist = "";
                PdfReader pdfReader = new PdfReader(filename);
                IList<Dictionary<string, object>> bookmarks = SimpleBookmark.GetBookmark(pdfReader);
                bookmarksDict = new Dictionary<string, object>();
                //bookmarks will be null if no bookmarks found
                if (bookmarksDict != null) { 
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
                    bookmarkname = "";bookmarkpage = "";bookmarkpagenumber="";
                    Dictionary<string, object> bookmarkdetails = new Dictionary<string, object>();
                    if (bd.ContainsKey("Kids"))
                    {
                        //Deal with children
                        IList<Dictionary<string, object>> bdkids = (IList < Dictionary < string, object>>) bd["Kids"];
                        iterateBookmarks(ref bookmarklist, bdkids, level+1);
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
                if (tablename == "DSA") { 
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
                if (tablename == "DSANav") { 
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

            private string getAnnotations(PdfReader reader, int pagenumber, out string pageOwner ) {
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

            private string ManageBookmarks()
            {

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
                pdfText= PdfTextExtractor.GetTextFromPage(reader, pagenumber, pdfExtract.titlestrategy);
                TitleText = Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(pdfText)));
                return pageText;
            }

            //Chop the input PDF into separate pages
            public int Split()
            {
                int pageCount;
                
                string thumbfile = System.IO.Path.GetFileNameWithoutExtension(inputfile);
                FileInfo file = new FileInfo(inputfile);
                //Create the GhostScript stuff here as it's expensive.   Might be better to encapsulate in a separate class come to think of it
                string name = file.Name.Substring(0, file.Name.LastIndexOf("."));
                string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
                Ghostscript.NET.GhostscriptVersionInfo vesion;
                Ghostscript.NET.Rasterizer.GhostscriptRasterizer rasterizer = null;

                vesion = new Ghostscript.NET.GhostscriptVersionInfo(new System.Version(0, 0, 0), path + @"\gsdll64.dll", string.Empty, Ghostscript.NET.GhostscriptLicense.GPL);
                rasterizer = new Ghostscript.NET.Rasterizer.GhostscriptRasterizer();

                #region SetupTable
                metatable = createDataTable("DSA", new Dictionary<string, string>());
                navtable = createDataTable("DSANav", new Dictionary<string, string>());
                #endregion
                using (PdfReader reader = new PdfReader(inputfile))
                {
                    pageCount = reader.NumberOfPages;
                    try { TabName = reader.Info["Title"]; } catch (Exception e) { }//Just consume an error, seems hit and miss whether the PDF gives back a title or not
                    //Iterate around the PDF, keep these outside loop so they propogate downwards
                    string Chapter ="", Section = "", SubSection = "", SubSubSection = "", Sub3Section = "", Sub4Section = "", NavType = "", URL = "";
                    //Used to work out how to wire up navigation.  Pretty sure there might be a more elegant way to do this
                    string currentChapter = "", currentSection = "", currentSubSection = "", currentSubSubSection = "", currentSub3Section = "", oldPageType = "", oldPageText = "",Source = "",Target = "";
                    for (int pagenumber = 1; pagenumber <= reader.NumberOfPages; pagenumber++)
                    {
                        string pageComments = "", pageOwner = "", pageText = "", NavWeight = "", NavDesc = "",  pageType = "", titleText = ""; 
                        string thumbname = thumbfile + "-p" + pagenumber + ".jpg";
                        string BookmarkTitle = "", BookmarkLevel = "", PDFPage = "";

                        #region CreatePDF                        

                        #endregion CreatePDF
                        #region ExtractText    
                        pageText = ExtractText(pagenumber, reader, out titleText);
                        if (pageText == oldPageText) { pageType = "RunOn"; }     //Page is one of those where a single slide has more content than will fit on one PDF Page.  Is there a better way to check this?
                        #endregion ExtractText
                        #region pageType
                        //Have a stab at working out page type
                        //Based on the contents of the page we can identify reference pop ups (they have the same text as their parent with the additional layer of reference X)
                        //Regular Expressions give me a headache
                        if (pageText.Contains("Reference")){
                            pageType = "Reference";             //Really need an enum for these, but until we know what we're doing, text will work
                        }
                        #endregion 
                        #region GetAnnotation
                        //Get Annotations for this page
                        pageComments = getAnnotations(reader, pagenumber, out pageOwner);
                        #endregion
                        #region CheckBookmarks
                        int maxLevels = 6;
                        string[] currentNav = new string[maxLevels];
                        string[] thisNav = new string[maxLevels];
                        string[] NavLevel = new string[maxLevels];
                        string[] NvType = new string[3];
                        NvType[0] = "Primary";
                        NvType[1] = "Ref";
                        NvType[2] = "Popup";

                        NavLevel[0] = "Main";
                        NavLevel[1] = "Child";
                        NavLevel[2] = "SubChild";
                        NavLevel[3] = "SubSubChild";
                        NavLevel[4] = "Sub3Child";
                        NavLevel[5] = "Sub4Child";

                        if (currentNav[0] == "") { currentNav[0] = "Entry"; }



                        if (currentChapter == "") { currentChapter = "Entry"; } //Dummy Node to start things off
                        //Populates the dictionary, no return
                        BookMarkList(inputfile);
                        //Check to see Bookmark for this page unless it's a run on page in which case we ignore it
                        if (bookmarksDict.ContainsKey(pagenumber.ToString()) && pageType != "RunOn")
                        {
                            object wibble;
                            if (bookmarksDict.TryGetValue(pagenumber.ToString(), out wibble)) {
                                Dictionary<string, object> wibble1 = (Dictionary<string, object>)wibble;
                                BookmarkTitle = wibble1["Name"].ToString();
                                BookmarkLevel = wibble1["Level"].ToString();
                                PDFPage = wibble1["PDFPage"].ToString();
                                NavDesc = "Title" + BookmarkTitle + "Level " + BookmarkLevel;
                                Target = BookmarkTitle;
                                //Clear out Children at each level, so children retain the parent that would have been set earlier but parents get cleaned
                                //Tidy up the Nav Setting which is in every branch - doesn't really need to be.
                                switch (BookmarkLevel)
                                {
                                    case "1":
                                        Chapter = BookmarkTitle;Section = ""; SubSection = ""; SubSubSection = ""; Sub3Section = ""; Sub4Section = "";
                                        //Top Level Slide so Navigation will be back from the previous chapter?  At this point currentChapter will have the old Chapter in it (first time through the loop at level 1? Assumes PDF is in linear order)
                                        Source = currentChapter; NavType = "Main";NavWeight = "100";
                                        break;
                                    case "2":
                                        Section = BookmarkTitle; SubSection = ""; SubSubSection = ""; Sub3Section = ""; Sub4Section = "";
                                        //Level 2, so Navigation as a child of the current chapter
                                        Source = currentChapter; NavType = "Child"; NavWeight = "75";
                                        if (pageType=="Reference") { NavType = "Reference"; }
                                        break;
                                    case "3":
                                        SubSection = BookmarkTitle; SubSubSection = ""; Sub3Section = ""; Sub4Section = "";
                                        Source = currentSection; NavType = "SubChild"; NavWeight = "50";
                                        if (pageType == "Reference") { NavType = "SubReference";  }
                                        break;
                                    case "4":
                                        SubSubSection = BookmarkTitle; Sub3Section = ""; Sub4Section = "";
                                        Source = currentSubSection; NavType = "SubSubChild"; NavWeight = "30";
                                        if (pageType == "Reference") { NavType = "SubSubReference"; }
                                        break;
                                    case "5":
                                        Sub3Section = BookmarkTitle; Sub4Section = "";
                                        Source = currentSubSubSection; NavType = "Sub3Child"; NavWeight = "20";
                                        if (pageType == "Reference") { NavType = "Sub3Reference"; }
                                        break;
                                    case "6":
                                        Sub4Section = BookmarkTitle;
                                        Source = currentSub3Section; NavType = "Sub4Child"; NavWeight = "10";
                                        if (pageType == "Reference") { NavType = "Sub4Reference"; }
                                        break;
                                }
                            }
                        }
                        
                        if (!bookmarksDict.ContainsKey(pagenumber.ToString()) || pageType == "RunOn")
                        {
                            //Clear out targeting for sitemap as this page is not a separate page
                            Source = ""; Target = ""; NavWeight = ""; NavType = ""; thumbname = ""; BookmarkLevel = ""; PDFPage = ""; BookmarkTitle = "";NavDesc = "";
                        }

                        #endregion
                        string imagefile = GenerateThumbnail(rasterizer, pagenumber, vesion);
                        //Add to Excel Spreadsheet; This is probably where to also add the screenshot of the PDF to the Excel page, for now just linking to thumbname

                        metatable.Rows.Add(pagenumber, pagenumber, titleText, pageText, pageType, Chapter, Section, SubSection, SubSubSection, Sub3Section, Sub4Section, pageComments, pageOwner, imagefile);
                        //metatable.Rows.Add(pagenumber, pagenumber, titleText, pageText, pageType, Chapter, Section, SubSection, SubSubSection, Sub3Section, Sub4Section, pageComments, pageOwner, imagefile);
                        //Don't add empty rows (where a slides runs across multiple PDF pages)
                        if (NavWeight != "") { navtable.Rows.Add(Source, Target, NavWeight, NavType, thumbname, "", NavDesc, PDFPage); }
                        
                        #region Olds

                        //Set the "olds" (redo as function)
                        oldPageType = pageType;
                        oldPageText = pageText;
                        currentChapter = Chapter;
                        currentSection = Section;
                        currentSubSection = SubSection;
                        currentSubSubSection = SubSubSection;
                        currentSub3Section = Sub3Section;
                        #endregion


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

            public string GenerateThumbnail(Ghostscript.NET.Rasterizer.GhostscriptRasterizer rasterizer, int pagenumber, Ghostscript.NET.GhostscriptVersionInfo vesion)
            {
                string imagefile = System.IO.Path.Combine(outputfile, System.IO.Path.GetFileNameWithoutExtension(inputfile) + "-p" + pagenumber + ".jpg");
                string thumbFilePath = System.IO.Path.Combine(outputfile, "Thumb" + System.IO.Path.GetFileNameWithoutExtension(inputfile) + "-p" + pagenumber + ".png");
                rasterizer.Open(inputfile, vesion, false);
                    string pageFilePath = imagefile;
                    System.Drawing.Image img = rasterizer.GetPage(this.exportDPI, this.exportDPI, pagenumber);
                    img.Save(pageFilePath, ImageFormat.Jpeg);
                rasterizer.Close();
                img.Save(thumbFilePath, ImageFormat.Png);
                resizeImage(thumbFilePath);
                return thumbFilePath;
            }

            public void PDFToImage()
            {
                string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);

                Ghostscript.NET.Rasterizer.GhostscriptRasterizer rasterizer = null;
                Ghostscript.NET.GhostscriptVersionInfo vesion = new Ghostscript.NET.GhostscriptVersionInfo(new System.Version(0, 0, 0), path + @"\gsdll64.dll", string.Empty, Ghostscript.NET.GhostscriptLicense.GPL);

                using (rasterizer = new Ghostscript.NET.Rasterizer.GhostscriptRasterizer())
                {
                    rasterizer.Open(inputfile, vesion, false);

                    for (int i = 1; i <= rasterizer.PageCount; i++)
                    {
                        string pageFilePath = System.IO.Path.Combine(outputfile, System.IO.Path.GetFileNameWithoutExtension(inputfile) + "-p" + i.ToString() + ".jpg");
                        string thumbFilePath = System.IO.Path.Combine(outputfile, "Thumb" + System.IO.Path.GetFileNameWithoutExtension(inputfile) + "-p" + i.ToString() + ".png");
                        System.Drawing.Image img = rasterizer.GetPage(this.exportDPI, this.exportDPI, i);
                        img.Save(pageFilePath, ImageFormat.Jpeg);
                        img.Save(thumbFilePath, ImageFormat.Png);
                        resizeImage(thumbFilePath);
                    }

                    rasterizer.Close();
                }
            }

            private void resizeImage(string imagefilepath)
            {
                var image = new ImageMagick.MagickImage(imagefilepath);
                image.Resize(new ImageMagick.MagickGeometry(thumbnailwidth, thumbnailheight));
                image.Write(imagefilepath);
            }

            #region ExtractBackgroundImages
            //This extracts all the JPGs from the file (i.e. it will grab the screens without comments)
            private int extractJPGs() {
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

            internal void ExportToExcel(string excelfile, string Tabname, DataTable table)
            {
                //Name Tab by Date? By definition right now this will always be a new file
                
                ExcelExport ee = new ExcelExport();
                //if (String.IsNullOrWhiteSpace(ExcelFile) == true) { ExcelFile = "./" + Guid.NewGuid().ToString() + ".xlsx"; }
                if (String.IsNullOrWhiteSpace(Tabname)) { Tabname = "Tab " + DateTime.Now.ToShortDateString().Replace('/','-'); }

                ee.ExportToExcel(excelfile, Tabname, table);
                //ee.ExportToExcel(excelfile + "nav", "Nav", splitter.navtable);
/*                if (File.Exists(excelfile))
                {
                    excelfile = eh.AppendToExcel(table, Tabname, excelfile);
                }
                else
                {
                    excelfile = eh.ExportToExcel(table, Tabname, excelfile);
                }
                return excelfile;
*/
            }


            public string readSiteMap()
            {
                XDocument doc = XDocument.Load("TestSiteMap.graphml");
                ///graphml/graph/node/data/y:ShapeNode/y:NodeLabel`
                XDocument doc1 = XDocument.Parse(doc.ToString());
                string yednamespace = "{http://www.yworks.com/xml/graphml}";
                XNamespace ns = doc1.Root.Name.Namespace;
                //var Nodes = doc.Root.Elements().Select(x => x.Element("NodeLabel"));
                var Nodes  = doc.Descendants(yednamespace + "NodeLabel");
                //Crude - need to work out how to do the select properly above based on value
                foreach (var Node in Nodes)
                {
                    Console.WriteLine(Node.Value);
                    System.Diagnostics.Debug.WriteLine(Node.Value);
                    //Hard coded an update to test it out!
                    if (Node.Value == "Intro")
                    {
                        Node.Value = "Intro 1";
                    }
                }
                doc.Save("NewSiteMap.graphml");

                return "0";
            }

//For comparing previous set of images with new set, in order to identify changes
            public Boolean compareImages(string PathToImage1, string PathToImage2)
            {
                Boolean retval = false;




                return retval;
            }


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

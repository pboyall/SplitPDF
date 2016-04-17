using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using PdfSharp.Pdf;
using System.Drawing.Imaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data.SqlClient;
using System.Data;
using iTextSharp.text.pdf.parser;
using System.Xml.Linq;

namespace SplitPDF
{
    class Program
    {

        struct DataRow{
            string pageName;
            string pdfIndex;
            Image pdfImage;
        }

        static void Main(string[] args)
        {
            splitPDF splitter = new splitPDF();
            //Path.GetTempPath()
            string mydirectory = System.IO.Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]).ToString();
            splitter.inputfile = mydirectory + "\\Users.pdf";
            splitter.outputfile = mydirectory + "\\Output";     //Just testing
/*
            //Save each PDF as a separate PDF page
            int returned = splitter.Split();
            //Save each PDF page as a JPG
            splitter.PDFToImage(300);
            //Create a row in a spreadsheet for each PDF
            splitter.ExportToExcel("");     //No tabname for now - that would be if updating.  Later
*/
            splitter.readSiteMap();

        }



        public class splitPDF
        {

            public string ExcelFile { get; set; }
            public string TabName { get; set; }
            public string outputfile { get; set; }
            public string inputfile { get; set; }
            public string xmlfile { get; set; }
            DataTable table;
            //Dimensions for the box on the page where the Title Text is stored  (change to struct later) Not sure how to work out what these dimensions should be, don't like the idea of trial and error!
            float distanceInPixelsFromLeft = 174;
            float distanceInPixelsFromBottom = 1950;
            float width = 1000;
            float height = 200;

            public int Split()
            {
                //Create a disconnected Data Table
                table = new DataTable("DSA");
                //Hard coded column list for now just to test
                table.Columns.Add("PageReference", typeof(int));
                table.Columns.Add("PageOrder", typeof(int));
                table.Columns.Add("Description", typeof(string));
                table.Columns.Add("Label", typeof(string));
                //Not sure how to add a thumbnail column - another job for later

                FileInfo file = new FileInfo(inputfile);
                string name = file.Name.Substring(0, file.Name.LastIndexOf("."));
                int pageCount;

                using (PdfReader reader = new PdfReader(inputfile))
                {
                    pageCount = reader.NumberOfPages;
                    //Iterate around the PDF

                    for (int pagenumber = 1; pagenumber <= reader.NumberOfPages; pagenumber++)
                    {
                        string filename = pagenumber.ToString() + ".pdf";
                        
                        Document document = new Document();
                        PdfCopy copy = new PdfCopy(document, new FileStream(outputfile + "\\" + filename, FileMode.Create));

                        document.Open();

                        copy.AddPage(copy.GetImportedPage(reader, pagenumber));

                        document.Close();
//Extract Text from the page
                        ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                        var currentText = PdfTextExtractor.GetTextFromPage(reader, pagenumber, strategy);
                        string pageText =Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default,Encoding.UTF8,Encoding.Default.GetBytes(currentText)));
                        //Extract title text from page (some duplicated code here, leave it for now)
                        //Move outside loop when finished testing - handy for iterating sizes at the moment
                        Rectangle mediabox = reader.GetPageSize(pagenumber);

                        var rect = new System.util.RectangleJ(distanceInPixelsFromLeft,distanceInPixelsFromBottom,width,height);
                        var filters = new RenderFilter[1];
                        filters[0] = new RegionTextRenderFilter(rect);
                        strategy = new FilteredTextRenderListener(new LocationTextExtractionStrategy(),filters);
                        currentText = PdfTextExtractor.GetTextFromPage(reader, pagenumber, strategy);
                        string titleText = Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                        try { TabName = reader.Info["Title"]; } catch (Exception e) { }//Just consume an error
                        //Add to Excel Spreadsheet
                        //Hard coded test - next step to build the description using text from the page
                        table.Rows.Add(pagenumber, pagenumber, titleText, pageText);

                    }
                    return pageCount;
                }

            }

            public void PDFToImage(int dpi)
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

                         System.Drawing.Image img = rasterizer.GetPage(dpi, dpi, i);
                        img.Save(pageFilePath, ImageFormat.Jpeg);
                    }

                    rasterizer.Close();
                }
            }


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

            internal string ExportToExcel(string Tabname)
            {
                //Name Tab by Date? By definition right now this will always be a new file
                string excelfile = outputfile + "\\" + Guid.NewGuid().ToString() + ".xlsx";
                //if (String.IsNullOrWhiteSpace(ExcelFile) == true) { ExcelFile = "./" + Guid.NewGuid().ToString() + ".xlsx"; }
                if (String.IsNullOrWhiteSpace(Tabname)) { Tabname = "Tab " + DateTime.Now.ToShortDateString().Replace('/','-'); }
                if (File.Exists(excelfile))
                {
                    excelfile = new ExcelHelper().AppendToExcel(table, Tabname, excelfile);
                }
                else
                {
                    excelfile = new ExcelHelper().ExportToExcel(table, Tabname, excelfile);
                }
                return excelfile;
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
  
        }





    }
}

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
            //Presentation Builder
            DSAProject Proj = new DSAProject();
            Proj.Product = "HUMIRA";
            Proj.Indication = "Uveitis";
            Proj.Segment = "Uveitis";
            Proj.Country = "UK";
            Proj.Language = "EN";
            Proj.Source = "Uveitis";
            Proj.Campaign = "YOU";
            Proj.Season = "SUMMER 16";
            Proj.Audience = "REP";
            Proj.Notes = "Test Project";

            Presentation Home = new Presentation();
            Home.Hidden = "N";
            Home.PresentationIndex = 1;
            Home.PresentationName = "Home";
            Home.project = Proj;
            Proj.Presentations.Add(Home.PresentationIndex.ToString(), Home);

            //Test harness
            splitPDF splitter = new splitPDF();
            //Path.GetTempPath()
            //Set up properties
            string mydirectory = System.IO.Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]).ToString();
            //splitter.inputfile = mydirectory + "\\BookmarkTesting.pdf";
            //15863 Uveitis Detail Aid_Visual 1_02 BOOKMARKED
            //
            //15863 Uveitis Detail Aid_Impact and TNF alpha_03 BOOKMARKED
            string inputfile = "15863 Uveitis Detail Aid_Visual 2_02 BOOKMARKED";
            splitter.inputfile = mydirectory + "\\" + inputfile + ".pdf";
            splitter.outputfile = mydirectory + "\\Output";     //Just testing
            splitter.comparisonfile = mydirectory + "\\Output\\2244bb38-5e6b-450a-80dd-c490ec6344b0.xlsx";     //Just testing
            splitter.renderer.exportDPI = 150;
            splitter.renderer.thumbnailheight = 150;
            splitter.renderer.thumbnailwidth = 200;
            splitter.createPDFs = false;
            splitter.createThumbs = true;
            splitter.consolidatePages = true;
            //Execute code
            int returned = splitter.Split();
            string excelfile = splitter.outputfile + "\\" + inputfile + ".xlsx";
            splitter.ExportToExcel(excelfile, "Meta", "Meta");     //No tabname for now - that would be if updating.  Later Guid.NewGuid().ToString() 
            splitter.ExportToExcel(excelfile, "Nav", "Nav");     //No tabname for now - that would be if updating.  Later

            //Updates to Site Map can be called from here 
            /*
            splitter.readSiteMap();
            */
        }
    }
}

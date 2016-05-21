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
            //Test harness
            splitPDF splitter = new splitPDF();
            //Path.GetTempPath()
            //Set up properties
            string mydirectory = System.IO.Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]).ToString();
            splitter.inputfile = mydirectory + "\\BookmarkTesting.pdf";
            splitter.outputfile = mydirectory + "\\Output";     //Just testing
            splitter.renderer.exportDPI = 300;
            splitter.renderer.thumbnailheight = 150;
            splitter.renderer.thumbnailwidth = 200;
            splitter.createPDFs = true;
            //Execute code
            int returned = splitter.Split();
            string excelfile = splitter.outputfile + "\\" + Guid.NewGuid().ToString() + ".xlsx";
            splitter.ExportToExcel(excelfile, "Meta", "Meta");     //No tabname for now - that would be if updating.  Later
            splitter.ExportToExcel(excelfile, "Nav", "Nav");     //No tabname for now - that would be if updating.  Later

            //Updates to Site Map can be called from here 
            /*
            splitter.readSiteMap();
            */
        }
    }
}

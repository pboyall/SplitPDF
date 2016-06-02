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
using System.Reflection;

namespace SplitPDF
{
    class Program
    {

        static void Main(string[] args)
        {

            string sourcedirectory = "";
            string targetdirectory = "";
            int loopcounter = 1;

            //Test harness
            splitPDF splitter = new splitPDF();
            //Path.GetTempPath()
            //Set up properties
            string mydirectory = System.IO.Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]).ToString();

            if ((args.Length) > 0) { sourcedirectory = args[0]; }
            if ((args.Length) > 1) { targetdirectory = args[1]; }

            if (sourcedirectory == "") { sourcedirectory = mydirectory; }
            if (targetdirectory == "") { targetdirectory= mydirectory + "\\Output"; }
            splitter.comparisonfile = mydirectory + "\\Output\\2244bb38-5e6b-450a-80dd-c490ec6344b0.xlsx";     //Just testing

            //splitter.inputfile = mydirectory + "\\BookmarkTesting.pdf";
            //
            //15863 Uveitis Detail Aid_Impact and TNF alpha_03 BOOKMARKED
            //15863 Uveitis Detail Aid_Impact and TNF alpha_03 BOOKMARKED
            //"15863 Uveitis Detail Aid_Visual 1_02 BOOKMARKED";

            DSAProject Proj = new DSAProject();

            string[] systemvalues = System.IO.File.ReadAllLines(sourcedirectory + "\\Project.txt");
            foreach(string line in systemvalues)
            {
                int equalspos = line.IndexOf(":");
                if (equalspos > 0) { 
                    string value = line.Substring(equalspos + 1);
                    string key = line.Substring(0, equalspos);
                    setInstanceProperty<string>(Proj, key, value);
                }
            }

            //Presentation Builder
/*            
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
*/

            splitter.renderer.exportDPI = 150;
            splitter.renderer.thumbnailheight = 150;
            splitter.renderer.thumbnailwidth = 200;
            splitter.createPDFs = false;
            splitter.createThumbs = true;
            splitter.consolidatePages = true;
            splitter.outputfile = targetdirectory;
            //For each PDF in source directory, run the routine
            string[] dirs = Directory.GetFiles(sourcedirectory, "*.pdf");
            foreach (string dir in dirs)
            {
                splitter.inputfile = dir;
                //Create Presentation
                Presentation thisPres = new Presentation();
                thisPres.Hidden = "N";//TODO Magic
                thisPres.PresentationIndex = loopcounter;
                thisPres.PresentationName = System.IO.Path.GetFileNameWithoutExtension(dir);
                thisPres.project = Proj;
                Proj.Presentations.Add(thisPres.PresentationIndex.ToString(), thisPres);
                loopcounter++;
                //Execute code
                int returned = splitter.Split();
                string excelfile = splitter.outputfile + "\\" + thisPres.PresentationName + ".xlsx";
                splitter.ExportToExcel(excelfile, "Meta", "Meta");     //No tabname for now - that would be if updating.  Later Guid.NewGuid().ToString() 
                splitter.ExportToExcel(excelfile, "Nav", "Nav");     //No tabname for now - that would be if updating.  Later

                //Metadata Export
                splitter.ExportMetadata();
            }
            //Updates to Site Map can be called from here 
            /*
            splitter.readSiteMap();
            */
        }

        static void setInstanceProperty<PROPERTY_TYPE>(object instance, string propertyName, PROPERTY_TYPE value)
        {
            Type type = instance.GetType();
            PropertyInfo propertyInfo = type.GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public, null, typeof(PROPERTY_TYPE), new Type[0], null);

            propertyInfo.SetValue(instance, value, null);

            return;
        }

    }



}

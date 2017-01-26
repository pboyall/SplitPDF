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

        static int Main(string[] args)
        {

            string sourcedirectory = "";
            string targetdirectory = "";
            string comparisonfile = "";
            string[] systemvalues;
            int loopcounter = 1;

            //Test harness
            splitPDF splitter = new splitPDF();
            //Path.GetTempPath()
            //Set up properties
            string mydirectory = System.IO.Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]).ToString();

            if ((args.Length) > 0) { sourcedirectory = args[0]; }
            if ((args.Length) > 1) { targetdirectory = args[1]; }
            if ((args.Length) > 2) { comparisonfile = args[2]; }

            if (sourcedirectory == "") { sourcedirectory = mydirectory; }
            if (targetdirectory == "") { targetdirectory= mydirectory + "\\Output"; }
            if (comparisonfile == "") { comparisonfile = mydirectory + "\\2244bb38-5e6b-450a-80dd-c490ec6344b0.xlsx"; }

            splitter.comparisonfile = comparisonfile;     //Just testing

            DSAProject Proj = new DSAProject();

            try
            {
                systemvalues = System.IO.File.ReadAllLines(sourcedirectory + "\\Project.txt");
            }catch(Exception e)
            {
                //Console.log?
                return (1);
            }
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
            //TODO : Update to read from config
            splitter.renderer.exportDPI = 150;
            splitter.renderer.thumbnailheight = 150;
            splitter.renderer.thumbnailwidth = 200;
            splitter.createPDFs = false;
            splitter.createThumbs = true;
            splitter.consolidatePages = true;
            splitter.extractText = false;
            splitter.exportNav = false;
            splitter.outputfile = targetdirectory;
            splitter.thisproject = Proj;
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
                splitter.thispresentation = thisPres;
                
                int returned = splitter.Split();
                string excelfile = splitter.outputfile + "\\" + thisPres.PresentationName + ".xlsx";
                splitter.ExportToExcel(excelfile, "Meta", "Meta");     //No tabname for now - that would be if updating.  Later Guid.NewGuid().ToString() 
                //splitter.ExportToExcel(excelfile, "Nav", "Nav");     //No tabname for now - that would be if updating.  Later

                //Metadata Export
                splitter.ExportMetadata();

            }
            //Updates to Site Map can be called from here 
            /*
            splitter.readSiteMap();
            */
            return 0;
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

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


    public class PDFRenderer
    {

        internal Ghostscript.NET.GhostscriptVersionInfo vesion;
        internal Ghostscript.NET.Rasterizer.GhostscriptRasterizer rasterizer;
        public int exportDPI { get; set; }
        public int thumbnailheight { get; set; }
        public int thumbnailwidth { get; set; }
        public string outputfile { get; set; }
        public string inputfile { get; set; }

        public PDFRenderer()
        {

            string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            rasterizer = null;
            vesion = new Ghostscript.NET.GhostscriptVersionInfo(new System.Version(0, 0, 0), path + @"\gsdll64.dll", string.Empty, Ghostscript.NET.GhostscriptLicense.GPL);
            rasterizer = new Ghostscript.NET.Rasterizer.GhostscriptRasterizer();

        }

        public string GenerateThumbnail(int pagenumber, string thumbFilePath)
        {
            //TODO clean up hard coded name strings
            string imagefile  = System.IO.Path.Combine(outputfile, System.IO.Path.GetFileNameWithoutExtension(inputfile) + "-p" + pagenumber + ".png");
            rasterizer.Open(inputfile, vesion, false);
            string pageFilePath = imagefile;
            System.Drawing.Image img = rasterizer.GetPage(this.exportDPI, this.exportDPI, pagenumber);
            //img.Save(pageFilePath, ImageFormat.Png);
            //File.Delete(pageFilePath);
            rasterizer.Close();
            img.Save(thumbFilePath, ImageFormat.Png);
            resizeImage(thumbFilePath);
            return thumbFilePath;
        }


        private void resizeImage(string imagefilepath)
        {
            var image = new ImageMagick.MagickImage(imagefilepath);
            image.Resize(new ImageMagick.MagickGeometry(thumbnailwidth, thumbnailheight));
            image.Write(imagefilepath);
        }

        //For comparing previous set of images with new set, in order to identify changes
        public Boolean compareImages(string PathToImage1, string PathToImage2)
        {
            Boolean retval = false;




            return retval;
        }


    }
}



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
{ //Separate Class as Originally I hoped to only have to initialise all the expensive strategies and filters once - turns out you can't do that.  But still, encapsulates lots of the PDF mess.
    internal class PDFExtractor
    {
        //Dimensions for the box on the page where the Title Text is stored  (change to struct later) Not sure how to work out what these dimensions should be, don't like the idea of trial and error!
        public float distanceInPixelsFromLeft = 174;
        public float distanceInPixelsFromBottom = 1950;
        public float width = 1000;
        public float height = 200;
        public ITextExtractionStrategy bodystrategy { get; set; }
        public ITextExtractionStrategy titlestrategy { get; set; }

        public PDFExtractor()
        {
            bodystrategy = new SimpleTextExtractionStrategy();
            var filters = new RenderFilter[1];
            var titlerect = new System.util.RectangleJ(distanceInPixelsFromLeft, distanceInPixelsFromBottom, width, height);
            filters[0] = new RegionTextRenderFilter(titlerect);
            titlestrategy = new FilteredTextRenderListener(new LocationTextExtractionStrategy(), filters);
        }
    }
}

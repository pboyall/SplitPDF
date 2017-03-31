using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetLight;

namespace SplitPDF
{
    public class ReadExcel
    {
        public Dictionary<int, string> Read(string excelfile) {
            Dictionary<int, string> PDFPageToKeyMessage = new Dictionary<int, string>();
            using (SLDocument sl = new SLDocument(excelfile, "Sheet1"))
            {
                int PDFPageNumber = 0;
                string KeyMessage = "";
                SLWorksheetStatistics stats = sl.GetWorksheetStatistics();
                int iStartColumnIndex = stats.StartColumnIndex;
                for (int row = stats.StartRowIndex + 1; row <= stats.EndRowIndex; ++row)
                {
                    PDFPageNumber = sl.GetCellValueAsInt32(row, iStartColumnIndex);
                    KeyMessage = sl.GetCellValueAsString(row, iStartColumnIndex+1);
                    PDFPageToKeyMessage.Add(PDFPageNumber, KeyMessage);
                }
            }
            return PDFPageToKeyMessage;
        }

    }
}


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using SpreadsheetLight.Drawing;

namespace SplitPDF
{
    class ExcelExport
    {

        public int thumbCol = 14;
 
        public void ExportToExcel(string outputfile, string tabname, DataTable dt)
        {
            using (SLDocument sl = new SLDocument())
            {
                int iStartRowIndex = 1;
                int iStartColumnIndex = 2;

                sl.ImportDataTable(iStartRowIndex, iStartColumnIndex, dt, true);
                SLStyle style = sl.CreateStyle();
//                style.FormatCode = "yyyy/mm/dd hh:mm:ss";
//                sl.SetColumnStyle(4, style);
                // + 1 because the header row is included
                // - 1 because it's a counting thing, because the start row is counted.
                int iEndRowIndex = iStartRowIndex + dt.Rows.Count + 1 - 1;
                // - 1 because it's a counting thing, because the start column is counted.
                int iEndColumnIndex = iStartColumnIndex + dt.Columns.Count - 1;
                SLTable table = sl.CreateTable(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
                table.SetTableStyle(SLTableStyleTypeValues.Medium17);
                //table.HasTotalRow = true;
                //table.SetTotalRowFunction(5, SLTotalsRowFunctionValues.Sum);
                sl.InsertTable(table);
                //Rows 2 to end
                sl.SetRowHeight(2, dt.Rows.Count, 110);
                sl.SetColumnWidth(thumbCol, 30);
                //for each row read Thumbcol value and load data 
                for (int i = iStartRowIndex; i < dt.Rows.Count; i++)
                {
                    string filepath;
                    filepath = dt.Rows[i-1][thumbCol-1].ToString();
                    SLPicture pic = new SLPicture(filepath);
                    pic.SetPosition(i, thumbCol);
                    sl.InsertPicture(pic);
                }
                sl.SaveAs(outputfile);
                /*
                                SLPicture pic = new SLPicture("C:\\code\\SplitPDF\\SplitPDF\\bin\\Debug\\Output\\ThumbBookmarkTesting-p20.png");
                                pic.SetPosition(3, thumbCol);
                                sl.InsertPicture(pic);
                                sl.SaveAs(outputfile);
                */
                Console.WriteLine("End of program");


            }



        }

    }
}

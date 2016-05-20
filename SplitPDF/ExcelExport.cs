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
using System.IO;

namespace SplitPDF
{
    class ExcelExport
    {

        public int thumbCol = 14;
 
        public void ExportToExcel(string outputfile, string tabname, DataTable dt)
        {
            SLDocument sl;
            //Test if outputfile exists
            if (File.Exists(outputfile)) { sl = new SLDocument(outputfile); } else { sl = new SLDocument(); }
            string curSheet = sl.GetCurrentWorksheetName();
            if (curSheet.Equals(tabname)){
                    //Do Nothing
            } else { 
                List<string> sheets = sl.GetWorksheetNames();
                foreach (var sheet in sheets)
                {
                    if (sheet.Equals(tabname))
                    {
                        sl.SelectWorksheet(tabname);
                    }
                }
                curSheet = sl.GetCurrentWorksheetName();
                if (curSheet.Equals(tabname)) {
                    //Do nothing
                }else
                {
                    sl.AddWorksheet(tabname);
                }
            }
            sl.DeleteWorksheet(SLDocument.DefaultFirstSheetName);

            int iStartRowIndex = 1;
            int iStartColumnIndex = 2;

            sl.ImportDataTable(iStartRowIndex, iStartColumnIndex, dt, true);
            SLStyle style = sl.CreateStyle();
//                style.FormatCode = "yyyy/mm/dd hh:mm:ss";
//                sl.SetColumnStyle(4, style);
            int iEndRowIndex = iStartRowIndex + dt.Rows.Count + 1 - 1;
            // - 1 because it's a counting thing, because the start column is counted.
            int iEndColumnIndex = iStartColumnIndex + dt.Columns.Count - 1;
            SLTable table = sl.CreateTable(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
            table.SetTableStyle(SLTableStyleTypeValues.Medium17);
            sl.InsertTable(table);

            if (dt.TableName == "DSA")
            {
                //Rows 2 to end
                //Only have thumbnails on the metadata one
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
            }
            sl.SaveAs(outputfile);
            sl.Dispose();
        }

    }
}

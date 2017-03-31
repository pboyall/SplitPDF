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

        string metedatatabname = "Presentation-Slide metadata";

        public int thumbCol = 15;
        public int textCol = 6;
        SLDocument sl;
        public int iStartRowIndex = 1;
        public int iStartColumnIndex = 2;

        public void ExportToExcel(string outputfile, string tabname, DataTable dt)
        {
            
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

            sl.ImportDataTable(iStartRowIndex, iStartColumnIndex, dt, true);
            SLStyle style = sl.CreateStyle();
            style.SetWrapText(true);
            //                style.FormatCode = "yyyy/mm/dd hh:mm:ss";
            //                sl.SetColumnStyle(4, style);
            int iEndRowIndex = iStartRowIndex + dt.Rows.Count + 1 - 1;
            // - 1 because it's a counting thing, because the start column is counted.
            int iEndColumnIndex = iStartColumnIndex + dt.Columns.Count - 1;
            SLTable table = sl.CreateTable(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
            table.SetTableStyle(SLTableStyleTypeValues.Medium17);
            sl.InsertTable(table);
            try { 
            sl.AutoFitColumn(iStartColumnIndex, iEndColumnIndex);
            sl.AutoFitRow(iStartRowIndex + 1, iEndRowIndex);
            }
            catch (Exception e)
            {

            }

            if (dt.TableName == "DSA")
            {
                //Rows 2 to end
                //Only have thumbnails on the metadata one
                sl.SetRowHeight(iStartRowIndex+1, iEndRowIndex, 110);
                sl.SetColumnWidth(iStartColumnIndex, 20);
                sl.SetColumnWidth(textCol, 30);
                sl.SetColumnStyle(textCol, style);
                sl.SetColumnWidth(thumbCol, 30);

                //for each row read Thumbcol value and load data 
                for (int i = iStartRowIndex; i < iEndRowIndex; i++)
                {
                    string filepath;
                    filepath = dt.Rows[i-1][thumbCol-1].ToString();
                    try
                    {
                        SLPicture pic = new SLPicture(filepath);
                        pic.SetPosition(i, thumbCol);
                        sl.InsertPicture(pic);
                        pic = null;
                    }
                    catch (Exception e) { Console.Write("No Thumbnails"); }
                }
            }
            sl.SaveAs(outputfile.Replace(".pdf", ""));
            sl.Dispose();
        }


        public void ExportMetadata(DSAProject thisproject, DataTable dt)
        {
            iStartRowIndex = 35;
            iStartColumnIndex = 1;
            thumbCol = 1;

            //PSA_HUM_PSA_UK_EN_Destination You_SUMMER16_LO (Core slides)
            string outputfile = thisproject.Indication + "_" + thisproject.Product + "_" + thisproject.Segment + "_" + thisproject.Country + "_" + thisproject.Language + "_" + thisproject.Campaign + "_" + thisproject.Season + "_" + thisproject.Source + ".xlsx";

            //Test if outputfile exists
            if (File.Exists(outputfile)) { sl = new SLDocument(outputfile); } else { sl = new SLDocument(); }
            string curSheet = sl.GetCurrentWorksheetName();
            List<string> sheets = sl.GetWorksheetNames();
            foreach (var sheet in sheets)
            {
                if (sheet.Equals(metedatatabname))
                {
                    sl.SelectWorksheet(metedatatabname);
                }
            }
            curSheet = sl.GetCurrentWorksheetName();
            if (curSheet.Equals(metedatatabname))
            {
                //Do nothing
            }
            else
            {
                sl.AddWorksheet(metedatatabname);
            }
            sl.DeleteWorksheet(SLDocument.DefaultFirstSheetName);

            sl.ImportDataTable(iStartRowIndex, iStartColumnIndex, dt, true);
            SLStyle style = sl.CreateStyle();
            style.SetWrapText(true);
            int iEndRowIndex = iStartRowIndex + dt.Rows.Count + 1 - 1;
            // - 1 because it's a counting thing, because the start column is counted.
            int iEndColumnIndex = iStartColumnIndex + dt.Columns.Count - 1;
            SLTable table = sl.CreateTable(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
            table.SetTableStyle(SLTableStyleTypeValues.Medium17);
            sl.InsertTable(table);
            sl.SetRowHeight(iStartRowIndex + 1, iEndRowIndex, 110);
            sl.SetColumnWidth(iStartColumnIndex, 20);
            sl.SetColumnWidth(textCol, 30);
            sl.SetColumnStyle(textCol, style);
            sl.SetColumnWidth(thumbCol, 30);
            //for each row read Thumbcol value and load data 
            for (int i = iStartRowIndex; i < iEndRowIndex; i++)
            {
                string filepath;
                filepath = dt.Rows[i - 1][thumbCol - 1].ToString();
                try
                {
                    SLPicture pic = new SLPicture(filepath);
                    pic.SetPosition(i, thumbCol);
                    sl.InsertPicture(pic);
                    //
                }
                catch (Exception e) { Console.Write("No Thumbnails"); }
            }
            sl.SaveAs(outputfile);
            sl.Dispose();
        }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using System.Security.Principal;
using System.Data;
using System.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using System.Drawing.Imaging;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace SplitPDF
{
    internal class ExcelHelper
    {
        internal int thumbheight;
        internal int thumbwidth;
        private static int s_rId;
        private System.Drawing.Font font;
        private Graphics graphics;
        internal string addThumbNail(int rowNumber, int colNumber, string Tabname, string excelfile, string imagefile)
        {
            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(excelfile, isEditable: true))
            {
                var colOffset = 0;
                var rowOffset = 0;
                colNumber = 5;
                rowNumber = 10;
                //Iterate Workbook until we find the Sheet that they have named
                var workbook = excelDoc.WorkbookPart.Workbook;
                var sheets = workbook.Sheets.Cast<Sheet>().ToList();

                foreach (var worksheetPart in excelDoc.WorkbookPart.WorksheetParts)
                {
                    string partRelationshipId = excelDoc.WorkbookPart.GetIdOfPart(worksheetPart);
                    var correspondingPart = sheets.FirstOrDefault(
                        s => s.Id.HasValue && s.Id.Value == partRelationshipId);
                    string sheetName = correspondingPart.Name;
                    if (sheetName == Tabname)
                    {
                        var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                        if (!worksheetPart.Worksheet.ChildElements.OfType<Drawing>().Any())
                        {
                            worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                        }


                        if (drawingsPart.WorksheetDrawing == null)
                        {
                            drawingsPart.WorksheetDrawing = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
                        }

                        var worksheetDrawing = drawingsPart.WorksheetDrawing;

                        var imagePart = drawingsPart.AddImagePart(ImagePartType.Jpeg);

                        using (var stream = new FileStream(imagefile, FileMode.Open))
                        {
                            imagePart.FeedData(stream);
                        }
                        Bitmap bm = new Bitmap(imagefile);
                        DocumentFormat.OpenXml.Drawing.Extents extents = new DocumentFormat.OpenXml.Drawing.Extents();
                        var extentsCx = (long)bm.Width * (long)((float)914400 / bm.HorizontalResolution);
                        var extentsCy = (long)bm.Height * (long)((float)914400 / bm.VerticalResolution);
                        bm.Dispose();

                        var nvps = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
                        var nvpId = nvps.Count() > 0 ?
                            (UInt32Value)worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().Max(p => p.Id.Value) + 1 :
                            1U;
                        var oneCellAnchor = new Xdr.OneCellAnchor(
                             new Xdr.FromMarker
                             {
                                 ColumnId = new Xdr.ColumnId((colNumber - 1).ToString()),
                                 RowId = new Xdr.RowId((rowNumber - 1).ToString()),
                                 ColumnOffset = new Xdr.ColumnOffset(colOffset.ToString()),
                                 RowOffset = new Xdr.RowOffset(rowOffset.ToString())
                             },
                             new Xdr.Extent { Cx = extentsCx, Cy = extentsCy },
                             new Xdr.Picture(
                                 new Xdr.NonVisualPictureProperties(
                                     new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = "Picture " + nvpId, Description = imagefile },
                                     new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })
                                 ),
                                 new Xdr.BlipFill(
                                     new A.Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = A.BlipCompressionValues.Print },
                                     new A.Stretch(new A.FillRectangle())
                                 ),
                                 new Xdr.ShapeProperties(
                                     new A.Transform2D(
                                         new A.Offset { X = 0, Y = 0 },
                                         new A.Extents { Cx = extentsCx, Cy = extentsCy }
                                     ),
                                     new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }
                                 )
                             ),
                             new Xdr.ClientData()
                         );

                        int colindex = colNumber - 1;
                        //sheetData.Elements<Column>().Where(c => c. == colindex).FirstOrDefault();
                        SheetData sheetData = correspondingPart.GetFirstChild<SheetData>();
                        var theRow = sheetData.Elements<Row>().Where(c => c.RowIndex == rowNumber).FirstOrDefault();
                        theRow.Height = 122;
                        theRow.CustomHeight = true;


                        worksheetDrawing.Append(oneCellAnchor);
                        //Resize rows and columns to fit around image
                        //sheetData.Elements<Row>().


                        /*                        int iOffset = 10;

                                                var iImageHeight = bm.Height;
                                                var iImageWidth = bm.Width;

                                                int colStart = GetCellColIndex(oneCellAnchor[0]);
                                                int colEnd = GetCellColIndex(mergeCellCoordinates[0]) + 1;
                                                int rowStart = GetRowIndex(mergeCellCoordinates[1]);
                                                int rowEnd = GetRowIndex(mergeCellCoordinates[1]) + 1;
                                                DocumentFormat.OpenXml.Office2010.ExcelAc.List lstRows = worksheetPart.Worksheet.Descendants().Where(r => (r.RowIndex >= (uint)rowStart && r.RowIndex <= (uint)rowEnd)).ToList();
                                                double fCustomColWidth = (float)(((((iImageWidth + (2 * iOffset)) - 12) / 7) + 1) / (colEnd - colStart + 1));
                                                double fCustomRowHeight = ((((float)iImageHeight + (2 * (float)iOffset)) * 72) / (96 * lstRows.Count));

                                                column.Width = fCustomColWidth;
                                                column.CustomWidth = true;
                                                Row.Height = fCustomRowHeight
                                                Row.CustomHeight = true;
                        */




                    }
                    excelDoc.WorkbookPart.Workbook.Save();
                }
                excelDoc.Close();
            }
            // Close the document.
            return "";
        }

        private DoubleValue GetExcelCellWidth(double widthInPixel)
        {
            DoubleValue result = 0;
            if (widthInPixel > 12)
            {
                result = 1;
                result += (widthInPixel - 12) / 7;
            }
            else
                result = 1;

            return result;
        }

        internal class ColumnCaption
        {
            private static string[] Alphabets = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            private static ColumnCaption instance = null;
            private List<string> cellHeaders = null;
            public static ColumnCaption Instance
            {
                get
                {
                    if (instance == null)
                        return new ColumnCaption();
                    else return ColumnCaption.Instance;
                }
            }

            public ColumnCaption()
            {
                this.InitCollection();
            }

            private void InitCollection()
            {
                cellHeaders = new List<string>();

                foreach (string sItem in Alphabets)
                    cellHeaders.Add(sItem);

                foreach (string item in Alphabets)
                    foreach (string sItem in Alphabets)
                        cellHeaders.Add(item + sItem);
            }

            /// <summary>
            /// Returns the column caption for the given row & column index.
            /// </summary>
            /// <param name="rowIndex">Index of the row.</param>
            /// <param name="columnIndex">Index of the column.</param>
            /// <returns></returns>
            internal string Get(int rowIndex, int columnIndex)
            {
                return this.cellHeaders.ElementAt(columnIndex) + (rowIndex + 1).ToString();
            }
        }

        internal string ExportToExcel(DataTable table, string Tabname, string excelfile)
        {
            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Create(excelfile, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                CreateExcelParts(excelDoc, table, Tabname);
            }
            return excelfile;
        }

        //Read existing Excel

        internal string AppendToExcel(DataTable table, string Tabname, string excelfile)
        {
            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(excelfile, isEditable: true))
            {
                //Iterate Workbook until we find the Sheet that they have named
                var workbook = excelDoc.WorkbookPart.Workbook;
                var sheets = workbook.Sheets.Cast<Sheet>().ToList();
                bool added = false;

                foreach (var w in excelDoc.WorkbookPart.WorksheetParts)
                {
                    string partRelationshipId = excelDoc.WorkbookPart.GetIdOfPart(w);
                    var correspondingSheet = sheets.FirstOrDefault(
                        s => s.Id.HasValue && s.Id.Value == partRelationshipId);
                    string sheetName = correspondingSheet.Name;
                    if (sheetName == Tabname)
                    {
                        //Add rows to existing tab
                        SheetData sheetData = correspondingSheet.GetFirstChild<SheetData>();
                        AddRows(table, sheetData);
                        added = true;
                    }
                }
                //If not found, then new tab
                if (!added) { AddExcelSheet(excelDoc, table, Tabname); }

                excelDoc.WorkbookPart.Workbook.Save();
                // Close the document.
                excelDoc.Close();
            }


            return excelfile;
        }

        //Add Rows to existing sheet
        private void AddRows(DataTable table, SheetData sheetData)
        {
            System.Drawing.Font font = new System.Drawing.Font("Calibri", 11F);

            UInt32Value rowIndex = 1U;
            Row lastRow;
            try
            {
                lastRow = sheetData.Elements<Row>().LastOrDefault();
                rowIndex = lastRow.RowIndex;
            }
            catch { lastRow = null; }


            //Headings

            Row row1 = new Row();
            Columns columns1 = new Columns();
            for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
            {
                Cell cell = new Cell() { CellReference = ExcelHelper.ColumnCaption.Instance.Get((Convert.ToInt32((UInt32)rowIndex) - 2), columnIndex), DataType = CellValues.String };
                CellValue cellValue = new CellValue();
                columns1.Append(new Column() { Max = (UInt32Value)10U, BestFit = true, CustomWidth = true });//Addition
                cellValue.Text = table.Columns[columnIndex].ColumnName.ToString().FormatCode();
                cell.Append(cellValue);
                row1.Append(cell);
            }
            sheetData.Append(row1);

            //Data

            for (int rIndex = 0; rIndex < table.Rows.Count; rIndex++)
            {
                Row row = new Row()
                {
                    RowIndex = rowIndex++,
                    Spans = new ListValue<StringValue>() { InnerText = "1:3" },
                    DyDescent = 0.25D
                };

                for (int cIndex = 0; cIndex < table.Columns.Count; cIndex++)
                {
                    if (cIndex == 0)
                    {
                        Cell cell = new Cell() { CellReference = ExcelHelper.ColumnCaption.Instance.Get((Convert.ToInt32((UInt32)rowIndex) - 2), cIndex), DataType = CellValues.String };
                        CellValue cellValue = new CellValue();
                        string text = table.Rows[rIndex][cIndex].ToString();
                        //Addition
                        cellValue.Text = text;

                        double width = graphics.MeasureString(text, font).Width;
                        Column column = (columns1.ChildElements[cIndex] as Column);
                        DoubleValue currentWidth = GetExcelCellWidth(width + 5); //5 constant represents the padding
                        //sets the column width if the current cell width is maximum
                        if (column.Width != null)
                            column.Width = column.Width > currentWidth ? column.Width : currentWidth;
                        else
                            column.Width = currentWidth;
                        column.Min = UInt32Value.FromUInt32((uint)cIndex + 1);
                        column.Max = UInt32Value.FromUInt32((uint)cIndex + 2);
                        //End Addition
                        cell.Append(cellValue);
                        row.Append(cell);
                    }
                    else
                    {
                        Cell cell = new Cell() { CellReference = ExcelHelper.ColumnCaption.Instance.Get((Convert.ToInt32((UInt32)rowIndex) - 2), cIndex), DataType = CellValues.String };
                        CellValue cellValue = new CellValue();
                        cellValue.Text = table.Rows[rIndex][cIndex].ToString();
                        //Addition
                        string text = table.Rows[rIndex][cIndex].ToString();
                        cellValue.Text = text;

                        double width = graphics.MeasureString(text, font).Width;
                        Column column = (columns1.ChildElements[cIndex] as Column);
                        DoubleValue currentWidth = GetExcelCellWidth(width + 5); //5 constant represents the padding
                        //sets the column width if the current cell width is maximum
                        if (column.Width != null)
                            column.Width = column.Width > currentWidth ? column.Width : currentWidth;
                        else
                            column.Width = currentWidth;
                        column.Min = UInt32Value.FromUInt32((uint)cIndex + 1);
                        column.Max = UInt32Value.FromUInt32((uint)cIndex + 2);
                        //End Addition

                        cell.Append(cellValue);

                        row.Append(cell);
                    }
                }
                //Not sheetData.Append
                if (lastRow != null)
                {
                    sheetData.InsertAfter(row, lastRow);
                }
                else
                {
                    sheetData.Append(row);
                }
                lastRow = sheetData.Elements<Row>().LastOrDefault();
            }

        }


        //Add new tab to existing spreadsheet
        private void AddExcelSheet(SpreadsheetDocument spreadsheetDoc, DataTable data, string Tabname)
        {

            //Workbookpart
            WorkbookPart workbookpart = spreadsheetDoc.WorkbookPart;

            //workbookpart.Workbook = new Workbook();
            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            //WorksheetPart worksheetPart = workbookpart.WorksheetParts.LastOrDefault();
            SheetData sheetData = new SheetData();
            //Add Data
            AddRows(data, sheetData);
            worksheetPart.Worksheet = new Worksheet(sheetData);
            string relationshipId = spreadsheetDoc.WorkbookPart.GetIdOfPart(worksheetPart);

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDoc.WorkbookPart.Workbook.GetFirstChild<Sheets>();

            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = relationshipId,
                SheetId = sheetId,
                Name = Tabname
            };
            sheets.Append(sheet);

        }

        private void CreateExcelParts(SpreadsheetDocument spreadsheetDoc, DataTable data, string Tabname)
        {
            WorkbookPart workbookPart = spreadsheetDoc.AddWorkbookPart();
            CreateWorkbookPart(workbookPart, Tabname);

            int workBookPartCount = 1;

            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("rId" + (workBookPartCount++).ToString());
            CreateWorkbookStylesPart(workbookStylesPart);

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId" + (101).ToString());
            CreateWorksheetPart(workbookPart.WorksheetParts.ElementAt(0), data);

            SharedStringTablePart sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>("rId" + (workBookPartCount++).ToString());
            CreateSharedStringTablePart(sharedStringTablePart, data);


            workbookPart.Workbook.Save();
        }

        /// <summary>
        /// Creates the shared string table part.  This is used for strings that appear more than once (column headings, as far as I can see)
        /// </summary>
        /// <param name="sharedStringTablePart">The shared string table part.</param>
        /// <param name="sheetData">The sheet data.</param>
        private void CreateSharedStringTablePart(SharedStringTablePart sharedStringTablePart, DataTable sheetData)
        {
            UInt32Value stringCount = Convert.ToUInt32(sheetData.Rows.Count) + Convert.ToUInt32(sheetData.Columns.Count);

            SharedStringTable sharedStringTable = new SharedStringTable()
            {
                Count = stringCount,
                UniqueCount = stringCount
            };

            for (int columnIndex = 0; columnIndex < sheetData.Columns.Count; columnIndex++)
            {
                SharedStringItem sharedStringItem = new SharedStringItem();
                Text text = new Text();
                text.Text = sheetData.Columns[columnIndex].ColumnName;
                sharedStringItem.Append(text);
                sharedStringTable.Append(sharedStringItem);
            }
            //This adds all the data in one go - not entirely sure how it works 
            for (int rowIndex = 0; rowIndex < sheetData.Rows.Count; rowIndex++)
            {
                SharedStringItem sharedStringItem = new SharedStringItem();
                Text text = new Text();
                var theType = sheetData.Rows[rowIndex][0].GetType();
                Console.WriteLine(theType); //Always an integer
                text.Text = sheetData.Rows[rowIndex][0].ToString();
                sharedStringItem.Append(text);
                sharedStringTable.Append(sharedStringItem);
            }

            sharedStringTablePart.SharedStringTable = sharedStringTable;
        }

        /// <summary>
        /// Creates the worksheet part.
        /// </summary>
        /// <param name="worksheetPart">The worksheet part.</param>
        /// <param name="data">The data.</param>
        private void CreateWorksheetPart(WorksheetPart worksheetPart, DataTable data)
        {
            font = new System.Drawing.Font("Calibri", 11F);
            Bitmap image = new Bitmap(250, 150);
            graphics = Graphics.FromImage(image);
            Worksheet worksheet = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            SheetViews sheetViews = new SheetViews();
            SheetView sheetView = new SheetView() { WorkbookViewId = (UInt32Value)0U };
            Selection selection = new Selection() { ActiveCell = "A1" };
            sheetView.Append(selection);
            sheetViews.Append(sheetView);

            PageMargins pageMargins = new PageMargins()
            {
                Left = 0.7D,
                Right = 0.7D,
                Top = 0.75D,
                Bottom = 0.75D,
                Header = 0.3D,
                Footer = 0.3D
            };

            SheetFormatProperties sheetFormatPr = new SheetFormatProperties()
            {
                CustomHeight = true,
                DefaultRowHeight = 15D,
                DyDescent = 0.25D
            };

            SheetData sheetData = new SheetData();

            UInt32Value rowIndex = 1U;

            Row row1 = new Row()
            {
                CustomHeight=true,
                RowIndex = rowIndex++,
                Spans = new ListValue<StringValue>() { InnerText = "1:3" },
                DyDescent = 0.25D
            };
            row1.Height = thumbheight;
            row1.CustomHeight = true;

            //Headings
            //PB Try to generate columns here
            Columns columns1 = new Columns();
            for (int columnIndex = 0; columnIndex < data.Columns.Count; columnIndex++) {
                //Columns columns1 = worksheet.GetFirstChild<Columns>();

                Column column1;
                var theType = data.Columns[columnIndex].DataType;
                //var theType = data.Rows[rIndex][cIndex].GetType();
                Console.WriteLine("Type is" + theType);//Not sure why but can't get it to match properly, messy string conversion
                if (theType.ToString() == "System.Drawing.Image") {
                    DoubleValue currentImageWidth = GetExcelCellWidth(thumbwidth);
                    UInt32Value colmin = UInt32Value.FromUInt32((uint)columnIndex + 1);
                    UInt32Value colmax = UInt32Value.FromUInt32((uint)columnIndex + 2);
                    column1 = new Column() { Min = colmin, Max = colmax, Width = 16D, CustomWidth = true };
                }
                else
                {
                    column1 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 8D, CustomWidth = true };
                }
                columns1.Append(column1);
            }
            sheetData.Append(columns1);

//headings
            for (int columnIndex = 0; columnIndex < data.Columns.Count; columnIndex++)
            {
                Cell cell = new Cell() { CellReference = ExcelHelper.ColumnCaption.Instance.Get((Convert.ToInt32((UInt32)rowIndex) - 2), columnIndex), DataType = CellValues.String };
                CellValue cellValue = new CellValue();
                cellValue.Text = data.Columns[columnIndex].ColumnName.ToString().FormatCode();
                cell.Append(cellValue);
                row1.Append(cell);
            }
            sheetData.Append(row1);

            string drawingrID = GetNextRelationShipId();


            for (int rIndex = 0; rIndex < data.Rows.Count; rIndex++)
            {
                Row row = new Row()
                {
                    CustomHeight = true,
                    RowIndex = rowIndex++,
                    Spans = new ListValue<StringValue>() { InnerText = "1:3" },
                    DyDescent = 0.25D
                };
                row.Height = thumbheight;
                row.CustomHeight = true;
                for (int cIndex = 0; cIndex < data.Columns.Count; cIndex++)
                {
                    if (cIndex == 0)
                    {
                        Cell cell = new Cell() { CellReference = ExcelHelper.ColumnCaption.Instance.Get((Convert.ToInt32((UInt32)rowIndex) - 2), cIndex), DataType = CellValues.String };
                        CellValue cellValue = new CellValue();
                        cellValue.Text = data.Rows[rIndex][cIndex].ToString();
                        cell.Append(cellValue);
                        row.Append(cell);
                    }
                    else
                    {
                        var theType = data.Columns[cIndex].DataType;
                        Cell cell;
                        //var theType = data.Rows[rIndex][cIndex].GetType();
                        Console.WriteLine("Type is" + theType);//Not sure why but can't get it to match properly, messy string conversion
                        if (theType.ToString() == "System.Drawing.Image")
                        {
                            Console.WriteLine("This one is an image");
                            Column column = (columns1.ChildElements[cIndex] as Column);
                            DoubleValue currentImageWidth = GetExcelCellWidth(thumbwidth);
                            if (column.Width != null)
                                column.Width = column.Width > currentImageWidth ? column.Width : currentImageWidth;
                            else
                                column.Width = currentImageWidth;
                            column.Min = UInt32Value.FromUInt32((uint)cIndex + 1);
                            column.Max = UInt32Value.FromUInt32((uint)cIndex + 2);
                            cell = new Cell();
                            CellValue cellValue = new CellValue();
                            //Copied from ExcelBuilder.  If this doesn't work, AddThumbnail does provided we have a cell into which to insert it.
                            DrawingsPart drawingsPart = null;
                            Xdr.WorksheetDrawing worksheetDrawing = null;

                            if (worksheetPart.DrawingsPart == null)
                            {
                                drawingsPart = worksheetPart.AddNewPart<DrawingsPart>(drawingrID);
                                worksheetDrawing = new Xdr.WorksheetDrawing();
                                worksheetDrawing.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
                                worksheetDrawing.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                                drawingsPart.WorksheetDrawing = worksheetDrawing;
                            }
                            else if (worksheetPart.DrawingsPart != null && worksheetPart.DrawingsPart.WorksheetDrawing != null)
                            {
                                drawingsPart = worksheetPart.DrawingsPart;
                                worksheetDrawing = worksheetPart.DrawingsPart.WorksheetDrawing;
                            }

                            string imagerId = GetNextRelationShipId();
                            ImagePart imagePart = drawingsPart.AddNewPart<ImagePart>("image/png", imagerId);
                            GenerateImagePartContent(imagePart, data.Rows[rIndex][cIndex] as Image);
                            Xdr.TwoCellAnchor cellAnchor = AddTwoCellAnchor(rIndex, cIndex, rIndex + 1, cIndex + 1, imagerId);
                            worksheetDrawing.Append(cellAnchor);
                            worksheetDrawing.Save(drawingsPart);
                        }
                        else { 
                            cell = new Cell() {  CellReference = ExcelHelper.ColumnCaption.Instance.Get((Convert.ToInt32((UInt32)rowIndex) - 2), cIndex), DataType = CellValues.String };
                            CellValue cellValue = new CellValue();
                            cellValue.Text = data.Rows[rIndex][cIndex].ToString();
                            double width = graphics.MeasureString(cellValue.Text, font).Width;
                            Column column = (columns1.ChildElements[cIndex] as Column);
                            DoubleValue currentWidth = GetExcelCellWidth(width + 5); //5 constant represents the padding; sets the column width if the current cell width is maximum
                            if (column.Width != null)
                                column.Width = column.Width > currentWidth ? column.Width : currentWidth;
                            else
                                column.Width = currentWidth;
                            column.Min = UInt32Value.FromUInt32((uint)cIndex + 1);
                            column.Max = UInt32Value.FromUInt32((uint)cIndex + 2);
                            cell.Append(cellValue);
                        }
                        row.Append(cell);
                    }
                }
                sheetData.Append(row);
            }

            worksheet.Append(sheetViews);
            worksheet.Append(sheetFormatPr);
            worksheet.Append(sheetData);
            worksheet.Append(pageMargins);
            worksheetPart.Worksheet = worksheet;
        }

        private static System.Drawing.Image GetImageFromFile(string fileName)
        {
            //check the existence of the file in disc
            if (File.Exists(fileName))
            {
                System.Drawing.Image image = System.Drawing.Image.FromFile(fileName);
                return image;
            }
            else
                return null;
        }


        private string GetNextRelationShipId()
        {
            s_rId++;
            return "rId" + s_rId.ToString();
        }
        private void GenerateImagePartContent(ImagePart imagePart, Image image)
        {
            MemoryStream memStream = new MemoryStream();
            image.Save(memStream, ImageFormat.Png);
            memStream.Position = 0;
            imagePart.FeedData(memStream);
            memStream.Close();
        }


        private Xdr.TwoCellAnchor AddTwoCellAnchor(int startRow, int startColumn, int endRow, int endColumn, string imagerId)
        {
            Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor() { EditAs = Xdr.EditAsValues.OneCell };

            Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
            Xdr.ColumnId columnId1 = new Xdr.ColumnId();
            columnId1.Text = startColumn.ToString();
            Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
            columnOffset1.Text = "0";
            Xdr.RowId rowId1 = new Xdr.RowId();
            rowId1.Text = startRow.ToString();
            Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
            rowOffset1.Text = "0";

            fromMarker1.Append(columnId1);
            fromMarker1.Append(columnOffset1);
            fromMarker1.Append(rowId1);
            fromMarker1.Append(rowOffset1);

            Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
            Xdr.ColumnId columnId2 = new Xdr.ColumnId();
            columnId2.Text = endColumn.ToString();
            Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
            columnOffset2.Text = "0";// "152381";
            Xdr.RowId rowId2 = new Xdr.RowId();
            rowId2.Text = endRow.ToString();
            Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
            rowOffset2.Text = "0";//"152381";

            toMarker1.Append(columnId2);
            toMarker1.Append(columnOffset2);
            toMarker1.Append(rowId2);
            toMarker1.Append(rowOffset2);

            Xdr.Picture picture1 = new Xdr.Picture();

            Xdr.NonVisualPictureProperties nonVisualPictureProperties1 = new Xdr.NonVisualPictureProperties();
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Picture 1" };

            Xdr.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Xdr.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Xdr.BlipFill blipFill1 = new Xdr.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = imagerId };
            blip1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            Xdr.ShapeProperties shapeProperties1 = new Xdr.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 152381L, Cy = 152381L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);
            Xdr.ClientData clientData1 = new Xdr.ClientData();

            twoCellAnchor1.Append(fromMarker1);
            twoCellAnchor1.Append(toMarker1);
            twoCellAnchor1.Append(picture1);
            twoCellAnchor1.Append(clientData1);

            return twoCellAnchor1;
        }


        /// <summary>
        /// Creates the workbook styles part.
        /// </summary>
        /// <param name="workbookStylesPart">The workbook styles part.</param>
        private void CreateWorkbookStylesPart(WorkbookStylesPart workbookStylesPart)
        {
            Stylesheet stylesheet = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            StylesheetExtensionList stylesheetExtensionList = new StylesheetExtensionList();
            StylesheetExtension stylesheetExtension = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            DocumentFormat.OpenXml.Office2010.Excel.SlicerStyles slicerStyles = new DocumentFormat.OpenXml.Office2010.Excel.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };
            stylesheetExtension.Append(slicerStyles);
            stylesheetExtensionList.Append(stylesheetExtension);

            stylesheet.Append(stylesheetExtensionList);

            workbookStylesPart.Stylesheet = stylesheet;
        }

        /// <summary>
        /// Creates the workbook part.
        /// </summary>
        /// <param name="workbookPart">The workbook part.</param>
        private void CreateWorkbookPart(WorkbookPart workbookPart, string worksheetname)
        {
            Workbook workbook = new Workbook();
            Sheets sheets = new Sheets();

            Sheet sheet = new Sheet()
            {
                Name = worksheetname + 1,
                SheetId = Convert.ToUInt32(101),
                Id = "rId" + (101).ToString()
            };
            sheets.Append(sheet);

            CalculationProperties calculationProperties = new CalculationProperties()
            {
                CalculationId = (UInt32Value)123456U  // some default Int32Value
            };

            workbook.Append(sheets);
            workbook.Append(calculationProperties);

            workbookPart.Workbook = workbook;
        }


        public static Row InsertRow(uint rowIndex, WorksheetPart worksheetPart, Row insertRow, bool isNewLastRow = false)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            Row retRow = !isNewLastRow ? sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex) : null;

            // If the worksheet does not contain a row with the specified row index, insert one.
            if (retRow != null)
            {
                // if retRow is not null and we are inserting a new row, then move all existing rows down.
                if (insertRow != null)
                {
                    UpdateRowIndexes(worksheetPart, rowIndex, false);
                    UpdateMergedCellReferences(worksheetPart, rowIndex, false);
                    UpdateHyperlinkReferences(worksheetPart, rowIndex, false);

                    // actually insert the new row into the sheet
                    retRow = sheetData.InsertBefore(insertRow, retRow);  // at this point, retRow still points to the row that had the insert rowIndex

                    string curIndex = retRow.RowIndex.ToString();
                    string newIndex = rowIndex.ToString();

                    foreach (Cell cell in retRow.Elements<Cell>())
                    {
                        // Update the references for the rows cells.
                        cell.CellReference = new StringValue(cell.CellReference.Value.Replace(curIndex, newIndex));
                    }

                    // Update the row index.
                    retRow.RowIndex = rowIndex;
                }
            }
            else
            {
                // Row doesn't exist yet, shifting not needed.
                // Rows must be in sequential order according to RowIndex. Determine where to insert the new row.
                Row refRow = !isNewLastRow ? sheetData.Elements<Row>().FirstOrDefault(row => row.RowIndex > rowIndex) : null;

                // use the insert row if it exists
                retRow = insertRow ?? new Row() { RowIndex = rowIndex };

                IEnumerable<Cell> cellsInRow = retRow.Elements<Cell>();

                if (cellsInRow.Any())
                {
                    string curIndex = retRow.RowIndex.ToString();
                    string newIndex = rowIndex.ToString();

                    foreach (Cell cell in cellsInRow)
                    {
                        // Update the references for the rows cells.
                        cell.CellReference = new StringValue(cell.CellReference.Value.Replace(curIndex, newIndex));
                    }

                    // Update the row index.
                    retRow.RowIndex = rowIndex;
                }

                sheetData.InsertBefore(retRow, refRow);
            }

            return retRow;
        }


        /// <summary>
        /// Updates all of the Row indexes and the child Cells' CellReferences whenever
        /// a row is inserted or deleted.
        /// </summary>
        /// <param name="worksheetPart">Worksheet Part</param>
        /// <param name="rowIndex">Row Index being inserted or deleted</param>
        /// <param name="isDeletedRow">True if row was deleted, otherwise false</param>
        private static void UpdateRowIndexes(WorksheetPart worksheetPart, uint rowIndex, bool isDeletedRow)
        {
            // Get all the rows in the worksheet with equal or higher row index values than the one being inserted/deleted for reindexing.
            IEnumerable<Row> rows = worksheetPart.Worksheet.Descendants<Row>().Where(r => r.RowIndex.Value >= rowIndex);

            foreach (Row row in rows)
            {
                uint newIndex = (isDeletedRow ? row.RowIndex - 1 : row.RowIndex + 1);
                string curRowIndex = row.RowIndex.ToString();
                string newRowIndex = newIndex.ToString();

                foreach (Cell cell in row.Elements<Cell>())
                {
                    // Update the references for the rows cells.
                    cell.CellReference = new StringValue(cell.CellReference.Value.Replace(curRowIndex, newRowIndex));
                }

                // Update the row index.
                row.RowIndex = newIndex;
            }
        }

        /// <summary>
        /// Updates the MergedCelss reference whenever a new row is inserted or deleted. It will simply take the
        /// row index and either increment or decrement the cell row index in the merged cell reference based on
        /// if the row was inserted or deleted.
        /// </summary>
        /// <param name="worksheetPart">Worksheet Part</param>
        /// <param name="rowIndex">Row Index being inserted or deleted</param>
        /// <param name="isDeletedRow">True if row was deleted, otherwise false</param>
        private static void UpdateMergedCellReferences(WorksheetPart worksheetPart, uint rowIndex, bool isDeletedRow)
        {
            if (worksheetPart.Worksheet.Elements<MergeCells>().Count() > 0)
            {
                MergeCells mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();

                if (mergeCells != null)
                {
                    // Grab all the merged cells that have a merge cell row index reference equal to or greater than the row index passed in
                    List<MergeCell> mergeCellsList = mergeCells.Elements<MergeCell>().Where(r => r.Reference.HasValue)
                                                                                     .Where(r => GetRowIndex(r.Reference.Value.Split(':').ElementAt(0)) >= rowIndex ||
                                                                                                 GetRowIndex(r.Reference.Value.Split(':').ElementAt(1)) >= rowIndex).ToList();

                    // Need to remove all merged cells that have a matching rowIndex when the row is deleted
                    if (isDeletedRow)
                    {
                        List<MergeCell> mergeCellsToDelete = mergeCellsList.Where(r => GetRowIndex(r.Reference.Value.Split(':').ElementAt(0)) == rowIndex ||
                                                                                       GetRowIndex(r.Reference.Value.Split(':').ElementAt(1)) == rowIndex).ToList();

                        // Delete all the matching merged cells
                        foreach (MergeCell cellToDelete in mergeCellsToDelete)
                        {
                            cellToDelete.Remove();
                        }

                        // Update the list to contain all merged cells greater than the deleted row index
                        mergeCellsList = mergeCells.Elements<MergeCell>().Where(r => r.Reference.HasValue)
                                                                         .Where(r => GetRowIndex(r.Reference.Value.Split(':').ElementAt(0)) > rowIndex ||
                                                                                     GetRowIndex(r.Reference.Value.Split(':').ElementAt(1)) > rowIndex).ToList();
                    }

                    // Either increment or decrement the row index on the merged cell reference
                    foreach (MergeCell mergeCell in mergeCellsList)
                    {
                        string[] cellReference = mergeCell.Reference.Value.Split(':');

                        if (GetRowIndex(cellReference.ElementAt(0)) >= rowIndex)
                        {
                            string columnName = GetColumnName(cellReference.ElementAt(0));
                            cellReference[0] = isDeletedRow ? columnName + (GetRowIndex(cellReference.ElementAt(0)) - 1).ToString() : IncrementCellReference(cellReference.ElementAt(0), CellReferencePartEnum.Row);
                        }

                        if (GetRowIndex(cellReference.ElementAt(1)) >= rowIndex)
                        {
                            string columnName = GetColumnName(cellReference.ElementAt(1));
                            cellReference[1] = isDeletedRow ? columnName + (GetRowIndex(cellReference.ElementAt(1)) - 1).ToString() : IncrementCellReference(cellReference.ElementAt(1), CellReferencePartEnum.Row);
                        }

                        mergeCell.Reference = new StringValue(cellReference[0] + ":" + cellReference[1]);
                    }
                }
            }
        }

        /// <summary>
        /// Updates all hyperlinks in the worksheet when a row is inserted or deleted.
        /// </summary>
        /// <param name="worksheetPart">Worksheet Part</param>
        /// <param name="rowIndex">Row Index being inserted or deleted</param>
        /// <param name="isDeletedRow">True if row was deleted, otherwise false</param>
        private static void UpdateHyperlinkReferences(WorksheetPart worksheetPart, uint rowIndex, bool isDeletedRow)
        {
            Hyperlinks hyperlinks = worksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();

            if (hyperlinks != null)
            {
                Match hyperlinkRowIndexMatch;
                uint hyperlinkRowIndex;

                foreach (Hyperlink hyperlink in hyperlinks.Elements<Hyperlink>())
                {
                    hyperlinkRowIndexMatch = Regex.Match(hyperlink.Reference.Value, "[0-9]+");
                    if (hyperlinkRowIndexMatch.Success && uint.TryParse(hyperlinkRowIndexMatch.Value, out hyperlinkRowIndex) && hyperlinkRowIndex >= rowIndex)
                    {
                        // if being deleted, hyperlink needs to be removed or moved up
                        if (isDeletedRow)
                        {
                            // if hyperlink is on the row being removed, remove it
                            if (hyperlinkRowIndex == rowIndex)
                            {
                                hyperlink.Remove();
                            }
                            // else hyperlink needs to be moved up a row
                            else
                            {
                                hyperlink.Reference.Value = hyperlink.Reference.Value.Replace(hyperlinkRowIndexMatch.Value, (hyperlinkRowIndex - 1).ToString());

                            }
                        }
                        // else row is being inserted, move hyperlink down
                        else
                        {
                            hyperlink.Reference.Value = hyperlink.Reference.Value.Replace(hyperlinkRowIndexMatch.Value, (hyperlinkRowIndex + 1).ToString());
                        }
                    }
                }

                // Remove the hyperlinks collection if none remain
                if (hyperlinks.Elements<Hyperlink>().Count() == 0)
                {
                    hyperlinks.Remove();
                }
            }
        }

        /// <summary>
        /// Given a cell name, parses the specified cell to get the row index.
        /// </summary>
        /// <param name="cellReference">Address of the cell (ie. B2)</param>
        /// <returns>Row Index (ie. 2)</returns>
        public static uint GetRowIndex(string cellReference)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellReference);

            return uint.Parse(match.Value);
        }

        /// <summary>
        /// Increments the reference of a given cell.  This reference comes from the CellReference property
        /// on a Cell.
        /// </summary>
        /// <param name="reference">reference string</param>
        /// <param name="cellRefPart">indicates what is to be incremented</param>
        /// <returns></returns>
        public static string IncrementCellReference(string reference, CellReferencePartEnum cellRefPart)
        {
            string newReference = reference;

            if (cellRefPart != CellReferencePartEnum.None && !String.IsNullOrEmpty(reference))
            {
                string[] parts = Regex.Split(reference, "([A-Z]+)");

                if (cellRefPart == CellReferencePartEnum.Column || cellRefPart == CellReferencePartEnum.Both)
                {
                    List<char> col = parts[1].ToCharArray().ToList();
                    bool needsIncrement = true;
                    int index = col.Count - 1;

                    do
                    {
                        // increment the last letter
                        col[index] = Letters[Letters.IndexOf(col[index]) + 1];

                        // if it is the last letter, then we need to roll it over to 'A'
                        if (col[index] == Letters[Letters.Count - 1])
                        {
                            col[index] = Letters[0];
                        }
                        else
                        {
                            needsIncrement = false;
                        }

                    } while (needsIncrement && --index >= 0);

                    // If true, then we need to add another letter to the mix. Initial value was something like "ZZ"
                    if (needsIncrement)
                    {
                        col.Add(Letters[0]);
                    }

                    parts[1] = new String(col.ToArray());
                }

                if (cellRefPart == CellReferencePartEnum.Row || cellRefPart == CellReferencePartEnum.Both)
                {
                    // Increment the row number. A reference is invalid without this componenet, so we assume it will always be present.
                    parts[2] = (int.Parse(parts[2]) + 1).ToString();
                }

                newReference = parts[1] + parts[2];
            }

            return newReference;
        }


        /// <summary>
        /// Given a cell name, parses the specified cell to get the column name.
        /// </summary>
        /// <param name="cellReference">Address of the cell (ie. B2)</param>
        /// <returns>Column name (ie. A2)</returns>
        private static string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }
        public enum CellReferencePartEnum
        {
            None,
            Column,
            Row,
            Both
        }
        private static List<char> Letters = new List<char>() { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', ' ' };

        protected static void InsertImage(Worksheet ws, long x, long y, long? width, long? height, string sImagePath)
        {
            try
            {
                WorksheetPart wsp = ws.WorksheetPart;
                DrawingsPart dp;
                ImagePart imgp;
                Xdr.WorksheetDrawing wsd;

                ImagePartType ipt;
                switch (sImagePath.Substring(sImagePath.LastIndexOf('.') + 1).ToLower())
                {
                    case "png":
                        ipt = ImagePartType.Png;
                        break;
                    case "jpg":
                    case "jpeg":
                        ipt = ImagePartType.Jpeg;
                        break;
                    case "gif":
                        ipt = ImagePartType.Gif;
                        break;
                    default:
                        return;
                }

                if (wsp.DrawingsPart == null)
                {
                    //----- no drawing part exists, add a new one

                    dp = wsp.AddNewPart<DrawingsPart>();
                    imgp = dp.AddImagePart(ipt, wsp.GetIdOfPart(dp));
                    wsd = new Xdr.WorksheetDrawing();
                }
                else
                {
                    //----- use existing drawing part

                    dp = wsp.DrawingsPart;
                    imgp = dp.AddImagePart(ipt);
                    dp.CreateRelationshipToPart(imgp);
                    wsd = dp.WorksheetDrawing;
                }

                using (FileStream fs = new FileStream(sImagePath, FileMode.Open))
                {
                    imgp.FeedData(fs);
                }

                int imageNumber = dp.ImageParts.Count<ImagePart>();
                if (imageNumber == 1)
                {
                    Drawing drawing = new Drawing();
                    drawing.Id = dp.GetIdOfPart(imgp);
                    ws.Append(drawing);
                }

                NonVisualDrawingProperties nvdp = new NonVisualDrawingProperties();
                nvdp.Id = new UInt32Value((uint)(1024 + imageNumber));
                nvdp.Name = "Picture " + imageNumber.ToString();
                nvdp.Description = "";
                DocumentFormat.OpenXml.Drawing.PictureLocks picLocks = new DocumentFormat.OpenXml.Drawing.PictureLocks();
                picLocks.NoChangeAspect = true;
                picLocks.NoChangeArrowheads = true;
                NonVisualPictureDrawingProperties nvpdp = new NonVisualPictureDrawingProperties();
                nvpdp.PictureLocks = picLocks;
                NonVisualPictureProperties nvpp = new NonVisualPictureProperties();
                nvpp.NonVisualDrawingProperties = nvdp;
                nvpp.NonVisualPictureDrawingProperties = nvpdp;

                DocumentFormat.OpenXml.Drawing.Stretch stretch = new DocumentFormat.OpenXml.Drawing.Stretch();
                stretch.FillRectangle = new DocumentFormat.OpenXml.Drawing.FillRectangle();

                BlipFill blipFill = new BlipFill();
                DocumentFormat.OpenXml.Drawing.Blip blip = new DocumentFormat.OpenXml.Drawing.Blip();
                blip.Embed = dp.GetIdOfPart(imgp);
                blip.CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print;
                blipFill.Blip = blip;
                blipFill.SourceRectangle = new DocumentFormat.OpenXml.Drawing.SourceRectangle();
                blipFill.Append(stretch);

                DocumentFormat.OpenXml.Drawing.Transform2D t2d = new DocumentFormat.OpenXml.Drawing.Transform2D();
                DocumentFormat.OpenXml.Drawing.Offset offset = new DocumentFormat.OpenXml.Drawing.Offset();
                offset.X = 0;
                offset.Y = 0;
                t2d.Offset = offset;
                Bitmap bm = new Bitmap(sImagePath);

                DocumentFormat.OpenXml.Drawing.Extents extents = new DocumentFormat.OpenXml.Drawing.Extents();

                if (width == null)
                    extents.Cx = (long)bm.Width * (long)((float)914400 / bm.HorizontalResolution);
                else
                    extents.Cx = width;

                if (height == null)
                    extents.Cy = (long)bm.Height * (long)((float)914400 / bm.VerticalResolution);
                else
                    extents.Cy = height;

                bm.Dispose();
                t2d.Extents = extents;
                ShapeProperties sp = new ShapeProperties();
                sp.BlackWhiteMode = DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.Auto;
                sp.Transform2D = t2d;
                DocumentFormat.OpenXml.Drawing.PresetGeometry prstGeom = new DocumentFormat.OpenXml.Drawing.PresetGeometry();
                prstGeom.Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle;
                prstGeom.AdjustValueList = new DocumentFormat.OpenXml.Drawing.AdjustValueList();
                sp.Append(prstGeom);
                sp.Append(new DocumentFormat.OpenXml.Drawing.NoFill());

                DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture picture = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture();
                picture.NonVisualPictureProperties = nvpp;
                picture.BlipFill = blipFill;
                picture.ShapeProperties = sp;

                Position pos = new Position();
                pos.X = x;
                pos.Y = y;
                Extent ext = new Extent();
                ext.Cx = extents.Cx;
                ext.Cy = extents.Cy;
                AbsoluteAnchor anchor = new AbsoluteAnchor();
                anchor.Position = pos;
                anchor.Extent = ext;
                anchor.Append(picture);
                anchor.Append(new ClientData());
                wsd.Append(anchor);
                wsd.Save(dp);
            }
            catch (Exception ex)
            {
                throw ex; // or do something more interesting if you want
            }
        }

        protected static void InsertImage(Worksheet ws, long x, long y, string sImagePath)
        {
            InsertImage(ws, x, y, null, null, sImagePath);
        }

    }

    public static class Extensions
    {
        public static string FormatCode(this string sourceString)
        {
            if (sourceString.Contains("<"))
                sourceString = sourceString.Replace("<", "&lt;");

            if (sourceString.Contains(">"))
                sourceString = sourceString.Replace(">", "&gt;");

            return sourceString;
        }

    }



}



/*
Easier Code
public static void CreateSpreadsheetWorkbook(string filepath)
    {
        // Create a spreadsheet document by supplying the filepath.
        // By default, AutoSave = true, Editable = true, and Type = xlsx.
        SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
            Create(filepath, SpreadsheetDocumentType.Workbook);

        // Add a WorkbookPart to the document.
        WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
        workbookpart.Workbook = new Workbook();

        // Add a WorksheetPart to the WorkbookPart.
        WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        // Add Sheets to the Workbook.
        Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
            AppendChild<Sheets>(new Sheets());

        // Append a new worksheet and associate it with the workbook.
        Sheet sheet = new Sheet()
        {
            Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "mySheetNameISHere!"
        };
        sheets.Append(sheet);

        workbookpart.Workbook.Save();

        // Close the document.
        spreadsheetDocument.Close();
    }            



*/


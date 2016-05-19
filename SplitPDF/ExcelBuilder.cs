using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;

namespace SplitPDF
{
    public class ExcelBuilder
    {
        /// <summary>
        /// static field to maintain the relationship id
        /// </summary>
        private static int s_rId;
        /// <summary>
        /// Field for Data source
        /// </summary>
        private DataTable m_table;
        /// <summary>
        /// Collection to maintain string value of cell and to serialize the content in SharedString xml part
        /// </summary>
        private List<string> sharedStrings;
        /// <summary>
        /// Collection to maintain the image collection added into the excel workbook
        /// </summary>
        private Dictionary<string, Image> ImageCollection;
        /// <summary>
        /// Field to represent the default font - used to measure text for autofit or adjusting the cell width
        /// </summary>
        private System.Drawing.Font font;
        /// <summary>
        /// Graphics object to measure the cell value for autofit or adjusting the cell width
        /// </summary>
        private Graphics graphics;


        /// <summary>
        /// Create an instance of ExcelBuilder
        /// </summary>
        public ExcelBuilder()
        {
            Bitmap image = new Bitmap(100, 100);
            graphics = Graphics.FromImage(image);
            ImageCollection = new Dictionary<string, Image>();
            font = new System.Drawing.Font("Calibri", 11F);
            sharedStrings = new List<string>();
        }

        /// <summary>
        /// Set the DataSource of the Excel builder
        /// </summary>
        /// <param name="table"></param>
        public void SetDataSource(DataTable table)
        {
            m_table = table;
        }
        /// <summary>
        /// Create a new Excel file
        /// </summary>
        /// <param name="filePath">Path of the output file</param>
        public void CreatePackage(string filePath)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                CreateParts(package);
            }
        }

        /// <summary>
        /// Adds child parts and generates content of the specified part.
        /// </summary>
        /// <param name="workbook">Iinstance of SpreadsheetDocument</param>
        private void CreateParts(SpreadsheetDocument workbook)
        {
            WorkbookPart workbookPart = workbook.AddWorkbookPart();
            GenerateWorkbookPart1Content(workbookPart);

            WorksheetPart worksheetPart1 = workbookPart.AddNewPart<WorksheetPart>(GetNextRelationShipId());
            GenerateWorksheetPart1Content(worksheetPart1);

            SharedStringTablePart sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>(GetNextRelationShipId());
            GenerateSharedStringTablePart1Content(sharedStringPart);

        }

        /// <summary>
        /// Creates a worksheet reference and add it into sheetcollection of workbook part
        /// </summary>
        /// <param name="workbookPart1"></param>
        private void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x15" } };
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            workbook1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            workbook1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");

            Sheets sheetCollection = new Sheets();
            Sheet sheet1 = new Sheet() { Name = "Sheet1", SheetId = (UInt32Value)1U, Id = "rId1" };
            sheetCollection.Append(sheet1);

            workbook1.Append(sheetCollection);
            workbookPart1.Workbook = workbook1;
        }

        // Generates content of worksheetPart1.
        private void GenerateWorksheetPart1Content(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            // Creates an Columns instance and adds its children.
            Columns columns1 = new Columns();
            for (int i = 0; i < m_table.Columns.Count; i++)
            {
                columns1.Append(new Column() { Max = (UInt32Value)10U, BestFit = true, CustomWidth = true });
            }
            worksheet1.Append(columns1);

            SheetData sheetData1 = new SheetData();
            string drawingrID = GetNextRelationShipId();
            AppendSheetData(sheetData1, worksheetPart1, drawingrID, columns1);
            worksheet1.Append(sheetData1);

            if (worksheetPart1.DrawingsPart != null && worksheetPart1.DrawingsPart.WorksheetDrawing != null)
            {
                Drawing drawing1 = new Drawing() { Id = drawingrID };
                worksheet1.Append(drawing1);
            }
            worksheetPart1.Worksheet = worksheet1;
        }
        /// <summary>
        /// Fills the contents from DataTable to SheetData instance of worksheet part
        /// </summary>
        /// <param name="sheetData1">Instance of SheetData</param>
        /// <param name="worksheetPart">Instance of WorksheetPart</param>
        /// <param name="drawingrID">relationship id of drawing part</param>
        /// <param name="columns">Instance of Column collection</param>
        private void AppendSheetData(SheetData sheetData1, WorksheetPart worksheetPart, string drawingrID, Columns columns)
        {
            for (int rowIndex = 0; rowIndex < m_table.Rows.Count; rowIndex++)
            {
                Row row = new Row() { RowIndex = (UInt32Value)(rowIndex + 1U) };
                DataRow tableRow = m_table.Rows[rowIndex];
                for (int colIndex = 0; colIndex < tableRow.ItemArray.Length; colIndex++)
                {
                    Cell cell = new Cell();
                    CellValue cellValue = new CellValue();
                    object data = tableRow.ItemArray[colIndex];

                    if (data is int || data is float || data is double)
                    {
                        //if the data is int or float or double, then the data can be serialized along within the cell itself
                        cellValue.Text = data.ToString();
                        cell.Append(cellValue);
                    }
                    else if (data is string)
                    {
                        //if the data is string, then it should be written into Sharedstring part
                        //so using varaible named "sharedString" of type List<string> to hold the 
                        //string value while filling the contents
                        cell.DataType = CellValues.SharedString;
                        string text = data.ToString();
                        if (!sharedStrings.Contains(text))
                            sharedStrings.Add(text);
                        cellValue.Text = sharedStrings.IndexOf(text).ToString();

                        //Measure the text with default font and calculate the current cell width
                        double width = graphics.MeasureString(text, font).Width;
                        Column column = (columns.ChildElements[colIndex] as Column);
                        DoubleValue currentWidth = GetExcelCellWidth(width + 5); //5 constant represents the padding
                        //sets the column width if the current cell width is maximum
                        if (column.Width != null)
                            column.Width = column.Width > currentWidth ? column.Width : currentWidth;
                        else
                            column.Width = currentWidth;
                        column.Min = UInt32Value.FromUInt32((uint)colIndex + 1);
                        column.Max = UInt32Value.FromUInt32((uint)colIndex + 2);

                        cell.Append(cellValue);
                    }
                    else if (data is Image)
                    {
                        //Calculate & sets the column width & Row height based on the image size
                        Size imageSize = (data as Image).Size;
                        row.Height = imageSize.Height;
                        row.CustomHeight = true;
                        Column column = (columns.ChildElements[colIndex] as Column);
                        DoubleValue currentImageWidth = GetExcelCellWidth(imageSize.Width);
                        if (column.Width != null)
                            column.Width = column.Width > currentImageWidth ? column.Width : currentImageWidth;
                        else
                            column.Width = currentImageWidth;
                        column.Min = UInt32Value.FromUInt32((uint)colIndex + 1);
                        column.Max = UInt32Value.FromUInt32((uint)colIndex + 2);

                        //if the data is Image, we need to serailize its characteristics information in the drawing part
                        //and then raw image need to be added as Image part within file or package
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
                        Xdr.TwoCellAnchor cellAnchor = AddTwoCellAnchor(rowIndex, colIndex, rowIndex + 1, colIndex + 1, imagerId);
                        worksheetDrawing.Append(cellAnchor);
                        ImagePart imagePart = drawingsPart.AddNewPart<ImagePart>("image/png", imagerId);
                        GenerateImagePartContent(imagePart, data as Image);
                    }
                    row.Append(cell);
                }
                sheetData1.Append(row);
            }
        }
        /// <summary>
        /// Calculate the cell width in excel units by taking the actual width in pixel 
        /// </summary>
        /// <param name="widthInPixel">Actual GDI based width in pixel</param>
        /// <returns></returns>
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
        /// <summary>
        /// Generates the image part
        /// </summary>
        /// <param name="imagePart">Instance of the image part</param>
        /// <param name="image">Instance of the image which need to be added into the package</param>
        private void GenerateImagePartContent(ImagePart imagePart, Image image)
        {
            MemoryStream memStream = new MemoryStream();
            image.Save(memStream, ImageFormat.Png);
            memStream.Position = 0;
            imagePart.FeedData(memStream);
            memStream.Close();
        }
        /// <summary>
        /// Represents the bounds of the image, reference to image part and other characteristics using TwoCellAnchor class
        /// </summary>
        /// <param name="startRow">Starting row of the image</param>
        /// <param name="startColumn">starting column of the image</param>
        /// <param name="endRow">Ending row of the image</param>
        /// <param name="endColumn">ending column of the image</param>
        /// <param name="imagerId">Image's relationship id</param>
        /// <returns></returns>
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
        /// Generates the SharedString xml part using the string collection in SharedStrings (List<string>)
        /// </summary>
        /// <param name="part"></param>
        private void GenerateSharedStringTablePart1Content(SharedStringTablePart part)
        {
            SharedStringTable sharedStringTable1 = new SharedStringTable();
            sharedStringTable1.Count = new UInt32Value((uint)sharedStrings.Count);
            sharedStringTable1.UniqueCount = new UInt32Value((uint)sharedStrings.Count);

            foreach (string item in sharedStrings)
            {
                SharedStringItem sharedStringItem = new SharedStringItem();
                Text text = new Text();
                text.Text = item;

                sharedStringItem.Append(text);
                sharedStringTable1.Append(sharedStringItem);
            }
            part.SharedStringTable = sharedStringTable1;
        }
        /// <summary>
        /// Gets the next relationship id
        /// </summary>
        /// <returns></returns>
        private string GetNextRelationShipId()
        {
            s_rId++;
            return "rId" + s_rId.ToString();
        }
    }
}

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

namespace SplitPDF
{
    internal class ExcelHelper
    {
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

            for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
            {
                Cell cell = new Cell() { CellReference = ExcelHelper.ColumnCaption.Instance.Get((Convert.ToInt32((UInt32)rowIndex) - 2), columnIndex), DataType = CellValues.String };
                CellValue cellValue = new CellValue();
                cellValue.Text = table.Columns[columnIndex].ColumnName.ToString().FormatCode();
                cell.Append(cellValue);

                row1.Append(cell);
            }
            sheetData.Append(row1);


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
                        cellValue.Text = table.Rows[rIndex][cIndex].ToString();
                        cell.Append(cellValue);

                        row.Append(cell);
                    }
                    else
                    {
                        Cell cell = new Cell() { CellReference = ExcelHelper.ColumnCaption.Instance.Get((Convert.ToInt32((UInt32)rowIndex) - 2), cIndex), DataType = CellValues.String };
                        CellValue cellValue = new CellValue();
                        cellValue.Text = table.Rows[rIndex][cIndex].ToString();
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

            for (int rowIndex = 0; rowIndex < sheetData.Rows.Count; rowIndex++)
            {
                SharedStringItem sharedStringItem = new SharedStringItem();
                Text text = new Text();
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
                DefaultRowHeight = 15D,
                DyDescent = 0.25D
            };

            SheetData sheetData = new SheetData();

            UInt32Value rowIndex = 1U;

            Row row1 = new Row()
            {
                RowIndex = rowIndex++,
                Spans = new ListValue<StringValue>() { InnerText = "1:3" },
                DyDescent = 0.25D
            };

            //Headings

            for (int columnIndex = 0; columnIndex < data.Columns.Count; columnIndex++)
            {
                Cell cell = new Cell() { CellReference = ExcelHelper.ColumnCaption.Instance.Get((Convert.ToInt32((UInt32)rowIndex) - 2), columnIndex), DataType = CellValues.String };
                CellValue cellValue = new CellValue();
                cellValue.Text = data.Columns[columnIndex].ColumnName.ToString().FormatCode();
                cell.Append(cellValue);

                row1.Append(cell);
            }
            sheetData.Append(row1);

            for (int rIndex = 0; rIndex < data.Rows.Count; rIndex++)
            {
                Row row = new Row()
                {
                    RowIndex = rowIndex++,
                    Spans = new ListValue<StringValue>() { InnerText = "1:3" },
                    DyDescent = 0.25D
                };

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
                        Cell cell = new Cell() { CellReference = ExcelHelper.ColumnCaption.Instance.Get((Convert.ToInt32((UInt32)rowIndex) - 2), cIndex), DataType = CellValues.String };
                        CellValue cellValue = new CellValue();
                        cellValue.Text = data.Rows[rIndex][cIndex].ToString();
                        cell.Append(cellValue);

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


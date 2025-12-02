extern alias Interop2013;
using ClosedXML.Excel;
using MathNet.Numerics;
//using NPOI.OpenXmlFormats.Dml.
using NPOI.HSSF.UserModel;
using NPOI.OOXML.XSSF.UserModel;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXml4Net.Util;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using SixLabors.ImageSharp.PixelFormats;
using DataTable = System.Data.DataTable;
using IndexedColors = NPOI.SS.UserModel.IndexedColors;

namespace ExcelExport
{
    class Program
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);


        static void Main(string[] args)
        {
            Directory.SetCurrentDirectory(AppDomain.CurrentDomain.BaseDirectory); //change directory to exe file location (for task-scheduler)

            QueryManager obj = new QueryManager();//create query manager instance
            obj.connection = obj.ConnectDB();// connect to database

            var watch = System.Diagnostics.Stopwatch.StartNew(); //start timer
            string currentDateTime = DateTime.Now.ToString("MM/dd/yyyy HH:mm");

            //get queries and execute them                
            string[] sqlFiles = obj.GetQueryFiles();

            foreach (string sqlFile in sqlFiles)
            {
                try
                {
                    string query = File.ReadAllText(sqlFile);
                    string fileName = System.IO.Path.GetFileNameWithoutExtension(sqlFile);
                    //Console.WriteLine(query);
                    //obj.ExecuteQuery(query);
                    ExportToExcelWithoutPivotTable(obj.ExecuteQuery(query), fileName);
                    //AddSlicer(fileName);
                    //SendEmail(fileName);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error reading file '{sqlFile}': {ex.Message}");

                    using (StreamWriter errorWriter = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + @"..\..\..\..\logs\errorLog.log", true))
                    {
                        errorWriter.WriteLine($"Current Time: {currentDateTime} - Error: {ex.ToString()}");
                        errorWriter.WriteLine();
                    }
                    //logger.Error(string.Format(ex.Message));
                }
            }

            obj.connection.Dispose();
            

            //stop timer
            watch.Stop();
            Console.WriteLine($"Execution Time: {watch.Elapsed} ms");

            using (StreamWriter timeWriter = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + @"..\..\..\..\logs\Timer.log", true))
            {
                timeWriter.WriteLine($"Current Time: {currentDateTime} - Execution Time: {watch.Elapsed} ms");
            }

            //Console.ReadLine();
        }

        private static void ExportToExcelWithoutPivotTable(DataTable dataTable, string fileName)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");


            //Create Header Stylers
            XSSFCellStyle headerStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            XSSFDataFormat dataFormat = (XSSFDataFormat)workbook.CreateDataFormat();

            //Header Font Styling
            XSSFFont headerFont = (XSSFFont)workbook.CreateFont();
            headerFont.FontHeightInPoints = (short)11;
            headerFont.FontName = "Calibri";
            headerFont.Color = IndexedColors.White.Index;
            headerFont.IsBold = false;
            headerFont.IsItalic = false;

            headerStyle.SetFont(headerFont);

            //Header Background Color Styling
            //byte[] headerColor = new byte[] { 68, 114, 196 };
            Rgb24 headerColor = new Rgb24(68, 114, 196);
            headerStyle.SetFillForegroundColor(new XSSFColor(headerColor));
            headerStyle.FillPattern = FillPattern.SolidForeground;

            //Header Border Styling
            headerStyle.BorderBottom = BorderStyle.Thin;
            headerStyle.BorderTop = BorderStyle.Thin;
            headerStyle.BorderLeft = BorderStyle.Thin;
            headerStyle.BorderRight = BorderStyle.Thin;

            // Create header row
            IRow headerRow = sheet1.CreateRow(0);
            ICell headerTempCell;

            //Set Column Header Size
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                headerRow.CreateCell(i).SetCellValue(dataTable.Columns[i].ColumnName);
                headerTempCell = headerRow.GetCell(i);
                headerTempCell.CellStyle = headerStyle;
                sheet1.AutoSizeColumn(headerTempCell.ColumnIndex);
            }

            //Create Data Stylers
            XSSFCellStyle dataStyle1 = (XSSFCellStyle)workbook.CreateCellStyle();
            XSSFCellStyle dataStyle2 = (XSSFCellStyle)workbook.CreateCellStyle();
            XSSFCellStyle numericStyle = (XSSFCellStyle)workbook.CreateCellStyle();

            //Data Font Styling
            XSSFFont dataFont = (XSSFFont)workbook.CreateFont();
            dataFont.FontHeightInPoints = (short)11;
            dataFont.FontName = "Calibri";
            dataFont.Color = IndexedColors.Black.Index;
            dataFont.IsBold = false;
            dataFont.IsItalic = false;

            dataStyle1.SetFont(dataFont);
            dataStyle2.SetFont(dataFont);

            numericStyle.SetDataFormat(dataFormat.GetFormat("0"));

            //Data Background Color Styling
            //byte[] accent1 = new byte[3] { 221, 235, 247 };
            Rgb24 accent1 = new Rgb24(221, 235, 247);
            dataStyle1.SetFillForegroundColor(new XSSFColor(accent1));
            dataStyle1.FillPattern = FillPattern.SolidForeground;

            //byte[] accent2 = new byte[3] { 255, 255, 255 };
            Rgb24 accent2 = new Rgb24(255, 255, 255);
            dataStyle2.SetFillForegroundColor(new XSSFColor(accent2));
            dataStyle2.FillPattern = FillPattern.SolidForeground;

            ICell dataTempCell;

            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                IRow dataRow = sheet1.CreateRow(i + 1);


                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    dataRow.CreateCell(j).SetCellValue(dataTable.Rows[i][j].ToString());
                    dataTempCell = dataRow.GetCell(j);

                    if (dataTempCell.RowIndex.IsEven())
                        dataTempCell.CellStyle = dataStyle1;
                    else
                        dataTempCell.CellStyle = dataStyle2;

                    //Change cell value to double for numbers
                    if (double.TryParse(dataTempCell.StringCellValue, out _) && dataTempCell.StringCellValue != "")
                    {
                        dataTempCell.SetCellValue(double.Parse(dataTempCell.StringCellValue));
                        //dataTempCell.CellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0");
                    }

                    //if(DateUtil.IsCellDateFormatted(dataTempCell))
                    //{
                    //    XSSFCellStyle dateStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                    //    HSSFDataFormat df = (HSSFDataFormat)workbook.CreateDataFormat();
                    //    //dataTempCell.CellStyle = styles["cell"];
                    //    dateStyle.DataFormat = df.GetFormat("@");
                    //    dataTempCell.SetCellValue(DateTime.Now);

                    //    //dataTempCell.SetCellValue(dataTempCell.DateCellValue);
                        
                    //}
                }
            }

            //Resize Columns again after adding data
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                sheet1.AutoSizeColumn(i);
            }

            //Create Pivot Table and add to workbook
            AddPivotTable(workbook, sheet1, fileName);

            //AddExistingSheetToNewFile(workbook);
            



            // Save the workbook to a file
            using (FileStream fs = new FileStream(AppDomain.CurrentDomain.BaseDirectory + @"..\..\..\..\Reports\" + DateTime.Now.ToString("dd-MMM") + " - " + fileName + ".xlsx", FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }

            Console.WriteLine("Excel file for " + fileName + " generated successfully.");
            workbook.Close();
        }
        private static void AddPivotTable(IWorkbook workbook, ISheet dataSheet, string fileName)
        {

            //Create Pivot Table
            int numberOfSheets = workbook.NumberOfSheets;
            XSSFSheet sheet2 = (XSSFSheet)workbook.CreateSheet($"Pivot Table {numberOfSheets}");

            int firstRow = dataSheet.FirstRowNum;
            int lastRow = dataSheet.LastRowNum;
            int firstCol = dataSheet.GetRow(0).FirstCellNum;
            int lastCol = dataSheet.GetRow(0).LastCellNum;

            CellReference topLeft = new CellReference(firstRow, firstCol);
            CellReference botRight = new CellReference(lastRow, lastCol - 1);
            CellReference location = new CellReference("A5");
            AreaReference areaReference = new AreaReference(topLeft, botRight);

            XSSFPivotTable pivotTable1 = sheet2.CreatePivotTable(areaReference, location, dataSheet);

            Emailer emailer = new Emailer();
            var emailData = emailer.GetRecipients(); //Read from json file containing all info regarding the email

            foreach (KeyValuePair<string, EmailValue> kvp in emailData.emailValues)
            {
                if (kvp.Key == fileName) // If name of query from json file and name of query from the queries folder match, then create email with attachment and send
                {
                    if (kvp.Value.pivotTable)
                    {
                        //Add values to pivot table
                        for (int i = 0; i < kvp.Value.valueLabels.Length; i++)
                        {

                            //Get Column header for Column the pivot table function will be used on
                            int colIndex = kvp.Value.valueLabels[i];
                            CellReference pivotHeaderCR = new CellReference(dataSheet.FirstRowNum, kvp.Value.valueLabels[i]);
                            IRow pivotHeaderRow = dataSheet.GetRow(pivotHeaderCR.Row);
                            ICell pivotHeaderCell = pivotHeaderRow.GetCell(colIndex);


                            if (kvp.Value.valueFunctions[i] == DataConsolidateFunction.COUNT.Name)
                            {
                                pivotTable1.AddColumnLabel(DataConsolidateFunction.COUNT, kvp.Value.valueLabels[i], "Count of " + pivotHeaderCell.StringCellValue);

                                var dataFields = pivotTable1.GetCTPivotTableDefinition().dataFields.dataField;
                                if (dataFields.Count > 0)
                                {
                                    // Only set numFmtId for the last added field
                                    dataFields[dataFields.Count - 1].numFmtId = 3;
                                }
                            }

                            else if (kvp.Value.valueFunctions[i] == DataConsolidateFunction.SUM.Name)
                            {
                                pivotTable1.AddColumnLabel(DataConsolidateFunction.SUM, kvp.Value.valueLabels[i], "Sum of " + pivotHeaderCell.StringCellValue);

                                var dataFields = pivotTable1.GetCTPivotTableDefinition().dataFields.dataField;
                                if (dataFields.Count > 0)
                                {
                                    // Only set numFmtId for the last added field
                                    dataFields[dataFields.Count - 1].numFmtId = 3;
                                }

                            }

                            else if (kvp.Value.valueFunctions[i] == DataConsolidateFunction.AVERAGE.Name)
                            {
                                pivotTable1.AddColumnLabel(DataConsolidateFunction.AVERAGE, kvp.Value.valueLabels[i], "Average of " + pivotHeaderCell.StringCellValue);

                                var dataFields = pivotTable1.GetCTPivotTableDefinition().dataFields.dataField;
                                if (dataFields.Count > 0)
                                {
                                    // Only set numFmtId for the last added field
                                    dataFields[dataFields.Count - 1].numFmtId = 2;
                                }
                            }
                        }

                        
                        //Add rows to pivot table
                        for (int i = 0; i < kvp.Value.rowLabels.Length; i++)
                        {
                            if (kvp.Value.rowLabels.Length > 0)
                            {
                                pivotTable1.AddRowLabel(kvp.Value.rowLabels[i]);
                               
                                pivotTable1.GetCTPivotTableDefinition().pivotFields.GetPivotFieldArray(kvp.Value.rowLabels[i]).sortType = ST_FieldSortType.ascending;

                                //CollapseFields(pivotTable1, dataSheet, fileName);

                            }
                        }

                        if (kvp.Value.moveΣValue)
                        {
                            //Set default for Σ values to row label
                            pivotTable1.GetCTPivotTableDefinition().dataOnRows = true;

                            //Add new row field for data fields for Σ values
                            pivotTable1.GetCTPivotTableDefinition().rowFields.AddNewField().x = -2;
                            pivotTable1.GetCTPivotTableDefinition().rowFields.count = (uint)(kvp.Value.rowLabels.Length + 1);

                            //Remove data field from column
                            pivotTable1.GetCTPivotTableDefinition().colFields.field.Clear();
                        }


                        //Add columns to pivot table
                        if (kvp.Value.columnLabels.Length > 0)
                        {
                            for (int i = 0; i < kvp.Value.columnLabels.Length; i++)
                            {
                                AddColLabel(pivotTable1, kvp.Value.columnLabels[i], areaReference, lastCol, lastRow);
                                //pivotTable1.GetCTPivotTableDefinition().pivotFields.GetPivotFieldArray(kvp.Value.columnLabels[i]).dataField = true;
                            }
                        }

                        //Add Report Filter to Pivot Table
                        if (kvp.Value.filter.Length > 0)
                        {
                            for (int i = 0; i < kvp.Value.filter.Length; i++)
                            {
                                pivotTable1.AddReportFilter(kvp.Value.filter[i]);
                            }
                        }

                        //pivotTable1.AddReportFilter(0);
                        //pivotTable1.AddReportFilter(22);

                       // CollapseFields(pivotTable1, dataSheet, fileName);                

                    }

                    //Pivot Table 2
                    if (kvp.Value.pivotTable2)
                    {
                        numberOfSheets++;
                        XSSFSheet sheet3 = (XSSFSheet)workbook.CreateSheet($"Pivot Table {numberOfSheets}");
                        XSSFPivotTable pivotTable2 = sheet3.CreatePivotTable(areaReference, location, dataSheet);

                        //Add values to 2nd pivot table
                        for (int i = 0; i < kvp.Value.valueLabels2.Length; i++)
                        {
                            //Get Column header for Column the pivot table function will be used on
                            int colIndex = kvp.Value.valueLabels2[i];
                            CellReference pivotHeaderCR = new CellReference(dataSheet.FirstRowNum, kvp.Value.valueLabels2[i]);
                            IRow pivotHeaderRow = dataSheet.GetRow(pivotHeaderCR.Row);
                            ICell pivotHeaderCell = pivotHeaderRow.GetCell(colIndex);


                            if (kvp.Value.valueFunctions2[i] == DataConsolidateFunction.COUNT.Name)
                            {
                                pivotTable2.AddColumnLabel(DataConsolidateFunction.COUNT, kvp.Value.valueLabels2[i], "Count of " + pivotHeaderCell.StringCellValue);

                                var dataFields = pivotTable2.GetCTPivotTableDefinition().dataFields.dataField;
                                if (dataFields.Count > 0)
                                {
                                    // Only set numFmtId for the last added field
                                    dataFields[dataFields.Count - 1].numFmtId = 3;
                                }
                            }

                            else if (kvp.Value.valueFunctions2[i] == DataConsolidateFunction.SUM.Name)
                            {
                                pivotTable2.AddColumnLabel(DataConsolidateFunction.SUM, kvp.Value.valueLabels2[i], "Sum of " + pivotHeaderCell.StringCellValue);

                                var dataFields = pivotTable2.GetCTPivotTableDefinition().dataFields.dataField;
                                if (dataFields.Count > 0)
                                {
                                    // Only set numFmtId for the last added field
                                    dataFields[dataFields.Count - 1].numFmtId = 3;
                                }
                            }

                            else if (kvp.Value.valueFunctions2[i] == DataConsolidateFunction.AVERAGE.Name)
                            {
                                pivotTable2.AddColumnLabel(DataConsolidateFunction.AVERAGE, kvp.Value.valueLabels2[i], "Average of " + pivotHeaderCell.StringCellValue);

                                var dataFields = pivotTable2.GetCTPivotTableDefinition().dataFields.dataField;
                                if (dataFields.Count > 0)
                                {
                                    // Only set numFmtId for the last added field
                                    dataFields[dataFields.Count - 1].numFmtId = 2;
                                }
                            }
                        }

                        //Add rows to pivot table
                        for (int i = 0; i < kvp.Value.rowLabels2.Length; i++)
                        {
                            if (kvp.Value.rowLabels2.Length > 0)
                            {
                                pivotTable2.AddRowLabel(kvp.Value.rowLabels2[i]);
                                //pivotTable2.GetCTPivotTableDefinition().pivotFields.GetPivotFieldArray(kvp.Value.rowLabels2[i]).dataField = true;
                                pivotTable2.GetCTPivotTableDefinition().pivotFields.GetPivotFieldArray(kvp.Value.rowLabels2[i]).sortType = ST_FieldSortType.ascending;
                            }
                        }

                        if (kvp.Value.moveΣValue2)
                        {
                            //Set default for Σ values to row label
                            pivotTable2.GetCTPivotTableDefinition().dataOnRows = true;

                            var ctPivot = pivotTable2.GetCTPivotTableDefinition();
                            if (ctPivot.rowFields == null)
                            {
                                ctPivot.rowFields = new CT_RowFields();
                            }

                            //Add new row field for data fields for Σ values
                            pivotTable2.GetCTPivotTableDefinition().rowFields.AddNewField().x = -2;
                            pivotTable2.GetCTPivotTableDefinition().rowFields.count = (uint)(kvp.Value.rowLabels2.Length + 1);

                            //Remove data field from column
                            pivotTable2.GetCTPivotTableDefinition().colFields.field.Clear();
                        }

                        //Add columns to 2nd pivot table
                        if (kvp.Value.columnLabels2.Length > 0)
                        {
                            for (int i = 0; i < kvp.Value.columnLabels2.Length; i++)
                            {
                                AddColLabel(pivotTable2, kvp.Value.columnLabels2[i], areaReference, lastCol, lastRow);
                                //pivotTable2.GetCTPivotTableDefinition().pivotFields.GetPivotFieldArray(kvp.Value.columnLabels2[i]).dataField = true;

                            }
                        }             
                    }

                    //Pivot Table 3
                    if (kvp.Value.pivotTable3)
                    {
                        numberOfSheets++;
                        XSSFSheet sheet4 = (XSSFSheet)workbook.CreateSheet($"Pivot Table {numberOfSheets}");
                        XSSFPivotTable pivotTable3 = sheet4.CreatePivotTable(areaReference, location, dataSheet);

                        //Add values to 3nd pivot table
                        for (int i = 0; i < kvp.Value.valueLabels3.Length; i++)
                        {
                            //Get Column header for Column the pivot table function will be used on
                            int colIndex = kvp.Value.valueLabels3[i];
                            CellReference pivotHeaderCR = new CellReference(dataSheet.FirstRowNum, kvp.Value.valueLabels3[i]);
                            IRow pivotHeaderRow = dataSheet.GetRow(pivotHeaderCR.Row);
                            ICell pivotHeaderCell = pivotHeaderRow.GetCell(colIndex);


                            if (kvp.Value.valueFunctions3[i] == DataConsolidateFunction.COUNT.Name)
                            {
                                pivotTable3.AddColumnLabel(DataConsolidateFunction.COUNT, kvp.Value.valueLabels3[i], "Count of " + pivotHeaderCell.StringCellValue);
                                //pivotTable1.GetCTPivotTableDefinition().pivotFields.GetPivotFieldArray(kvp.Value.valueLabels[i]).dataField = true;

                                var dataFields = pivotTable3.GetCTPivotTableDefinition().dataFields.dataField;
                                if (dataFields.Count > 0)
                                {
                                    // Only set numFmtId for the last added field
                                    dataFields[dataFields.Count - 1].numFmtId = 3;
                                }
                            }

                            else if (kvp.Value.valueFunctions3[i] == DataConsolidateFunction.SUM.Name)
                            {
                                pivotTable3.AddColumnLabel(DataConsolidateFunction.SUM, kvp.Value.valueLabels3[i], "Sum of " + pivotHeaderCell.StringCellValue);
                                //pivotTable1.GetCTPivotTableDefinition().pivotFields.GetPivotFieldArray(kvp.Value.valueLabels[i]).dataField = true;

                                var dataFields = pivotTable3.GetCTPivotTableDefinition().dataFields.dataField;
                                if (dataFields.Count > 0)
                                {
                                    // Only set numFmtId for the last added field
                                    dataFields[dataFields.Count - 1].numFmtId = 3;
                                }
                            }

                            else if (kvp.Value.valueFunctions3[i] == DataConsolidateFunction.AVERAGE.Name)
                            {
                                pivotTable3.AddColumnLabel(DataConsolidateFunction.AVERAGE, kvp.Value.valueLabels3[i], "Average of " + pivotHeaderCell.StringCellValue);
                                //pivotTable1.GetCTPivotTableDefinition().pivotFields.GetPivotFieldArray(kvp.Value.valueLabels[i]).dataField = true;

                                var dataFields = pivotTable3.GetCTPivotTableDefinition().dataFields.dataField;
                                if (dataFields.Count > 0)
                                {
                                    // Only set numFmtId for the last added field
                                    dataFields[dataFields.Count - 1].numFmtId = 2;
                                }
                            }
                        }

                        //Add rows to pivot table
                        for (int i = 0; i < kvp.Value.rowLabels3.Length; i++)
                        {
                            if (kvp.Value.rowLabels3.Length > 0)
                            {
                                pivotTable3.AddRowLabel(kvp.Value.rowLabels3[i]);
                                //pivotTable3.GetCTPivotTableDefinition().pivotFields.GetPivotFieldArray(kvp.Value.rowLabels3[i]).dataField = true;
                                pivotTable3.GetCTPivotTableDefinition().pivotFields.GetPivotFieldArray(kvp.Value.rowLabels3[i]).sortType = ST_FieldSortType.ascending;
                            }
                        }

                        //Add columns to 3nd pivot table
                        if (kvp.Value.columnLabels3.Length > 0)
                        {
                            for (int i = 0; i < kvp.Value.columnLabels3.Length; i++)
                            {
                                AddColLabel(pivotTable3, kvp.Value.columnLabels3[i], areaReference, lastCol, lastRow);
                                //pivotTable3.GetCTPivotTableDefinition().pivotFields.GetPivotFieldArray(kvp.Value.columnLabels3[i]).dataField = true;
                            }
                        }

                        
                    }

                    

                    else if(!kvp.Value.pivotTable)
                    {
                        workbook.RemoveSheetAt(1);
                    }

                    //Hide Duplicate Columns
                    if (kvp.Value.duplicateColumns.Length > 0)
                    {
                        for (int i = 0; i < kvp.Value.duplicateColumns.Length; i++)
                        {
                            dataSheet.SetColumnHidden(kvp.Value.duplicateColumns[i], true);
                        }
                    }

                    
                }
            }
        }

        public static void AddSlicer(IWorkbook workbook)
        {
            XLWorkbook wb = (XLWorkbook)workbook;
            var ws = wb.Worksheets.Add("Test");

        }

        public static void AddColLabel(XSSFPivotTable pivotTable, int columnIndex, AreaReference areaReference, int lastColumn, int lastRow)
        {
            AreaReference pivotArea = areaReference;
            int lastRowIndex = lastRow;
            int lastColIndex = lastColumn;

            if (columnIndex > lastColIndex)
            {
                //throw new IndexOutOfBoundsException();
                throw new IndexOutOfRangeException();
            }
            CT_PivotFields pivotFields = pivotTable.GetCTPivotTableDefinition().pivotFields;

            CT_PivotField pivotField = new CT_PivotField();
                //CT_PivotField.Factory.newInstance();
            CT_Items items = pivotField.AddNewItems();

            pivotField.axis = ST_Axis.axisCol; //setAxis(STAxis.AXIS_COL);
            pivotField.showAll = false;// setShowAll(false);

            for (int i = 0; i <= lastRowIndex; i++)
            {
                items.AddNewItem().t = ST_ItemType.@default;
            }
            items.count = items.SizeOfItemArray();// setCount(items.sizeOfItemArray());
            pivotFields.SetPivotFieldArray(columnIndex, pivotField);

            CT_ColFields colFields;
            if (pivotTable.GetCTPivotTableDefinition().colFields != null)
            {
                colFields = pivotTable.GetCTPivotTableDefinition().colFields;
            }
            else
            {
                colFields = pivotTable.GetCTPivotTableDefinition().AddNewColFields();
            }

            colFields.AddNewField().x = columnIndex;// setX(columnIndex);
            colFields.count = colFields.SizeOfFieldArray();// setCount(rowFields.sizeOfFieldArray());

            



            //pivotTable.GetCTPivotTableDefinition().pivotFields.GetPivotFieldArray(columnIndex).dataField = true;
            //pivotField.dataField = true;
        }

        public static void CollapseFields(XSSFPivotTable pivotTable, ISheet dataSheet, string fileName)
        {

            Emailer emailer = new Emailer();
            var emailData = emailer.GetRecipients(); //Read from json file containing all info regarding the email

            foreach (KeyValuePair<string, EmailValue> kvp in emailData.emailValues)
            {
                if (kvp.Key == fileName)
                {
                    if (kvp.Value.collapseField)
                    {
                        //we need unique contents from 2nd row label for creating the pivot cache

                        //List<string> collapseRowValues = new List<string>();

                        //for (int r = 1; r < dataSheet.LastRowNum + 1; r++)
                        //{
                        //    IRow row = dataSheet.GetRow(r);
                        //    if (row != null)
                        //    {
                        //        ICell cell = row.GetCell(kvp.Value.rowLabels[1]);
                        //        if (cell != null)
                        //        {
                        //            collapseRowValues.Add(cell.StringCellValue);
                        //        }
                        //    }
                        //}

                        ////now go through all pivot items of first pivot field 
                        //List<CT_Item> itemList = pivotTable.GetCTPivotTableDefinition().pivotFields.GetPivotFieldArray(kvp.Value.rowLabels[1]).items.item;
                        //int i = 0;
                        //CT_Item? item = null;

                        //foreach (string value in collapseRowValues)
                        //{
                        //    item = itemList[i];
                        //    item.t = ST_ItemType.blank;
                        //    item.x = (uint)i++;
                        //    //item.s = value;
                        //    pivotTable.GetPivotCacheDefinition().GetCTPivotCacheDefinition().cacheFields.cacheField[kvp.Value.rowLabels[0]].sharedItems.Items.Add(item.n);
                        //    //CT_Item item2 = (CT_Item)pivotTable.GetPivotCacheDefinition().GetCTPivotCacheDefinition().cacheFields.cacheField[kvp.Value.rowLabels[1]].sharedItems.Items.Last();
                        //    //pivotTable.GetPivotCacheDefinition().GetCTPivotCacheDefinition().cacheFields.cacheField[kvp.Value.rowLabels[1]].sharedItems.Items.Last() = item2 ;
                        //    //item2.
                        //    item.sd = false;

                        //    var a = pivotTable.GetCTPivotTableDefinition();
                        //    var b = pivotTable.GetPivotCacheDefinition().GetCTPivotCacheDefinition().cacheFields.cacheField.;
                        //    pivotTable.GetCTPivotTableDefinition().pivotFields.pivotField[0].items.item[0].



                        //}

                        //while (i < itemList.Count) 
                        //{
                        //    item = itemList[i++];
                        //    item.sd = false;
                        //}



                        //CT_PivotField pivotField = pivotTable.GetCTPivotTableDefinition().pivotFields.pivotField[kvp.Value.rowLabels[1]];
                        //int j = 0;

                        //foreach (string value in collapseRowValues) 
                        //{
                        //    pivotField.items.item[j].t = ST_ItemType.blank;
                        //}


                        CT_PivotTableDefinition a = pivotTable.GetCTPivotTableDefinition();
                        CT_PivotFields b = a.pivotFields;
                        CT_RowFields c = a.rowFields;

                        foreach (var rf in c.field)
                        {
                            int index = (int)rf.x;

                            CT_PivotCacheDefinition d = pivotTable.GetPivotCacheDefinition().GetCTPivotCacheDefinition();

                            b.pivotField[index].showAll = false;
                            b.pivotField[index].outline = true;
                            b.pivotField[index].items.item[1].h = true;

                        }
                    }
                }
            }         
        }
        private static void SendEmail(string fileName)
        {
            Emailer emailer = new Emailer();
            var emailData = emailer.GetRecipients(); //Read from json file containing all info regarding the email
            //emailData.emailValues.;

            foreach (KeyValuePair<string, EmailValue> kvp in emailData.emailValues)
            {
                if (kvp.Key == fileName) // If name of query from json file and name of query from the queries folder match, then create email with attachment and send
                {
                    string reportPath = AppDomain.CurrentDomain.BaseDirectory + @"..\..\..\..\Reports\" + DateTime.Now.ToString("dd-MMM") + " - " + fileName + ".xlsx";
                    
                    for (int i = 0; i < kvp.Value.address.Length; i++)
                    {
                        emailer.SendEmail(kvp.Value.address[i],kvp.Value.cc[i], kvp.Value.subject, kvp.Value.body, reportPath);
                    }
                }
            }
        }

        private static void AddExistingSheetToNewFile(IWorkbook workbook)
        {
            IWorkbook workbook2 = new XSSFWorkbook();

            using (FileStream fs = new FileStream(AppDomain.CurrentDomain.BaseDirectory + @"..\..\..\..\Reports\" + DateTime.Now.ToString("dd-MMM") + " - " + "GL Code.xlsx", FileMode.Open, FileAccess.Read))
            {
                workbook2 = WorkbookFactory.Create(fs);
            }

            ISheet copySheet = workbook2.GetSheetAt(0);

            ISheet cSheet = workbook.CloneSheet(workbook2.GetSheetIndex(copySheet));
            //workbook.CreateSheet("GL Code");
            
        }
    }
}



/* Multiple Attachemnts: create all reports and store them in a new directory. Once all reports generated, then send email attaching all reports in that directory
 * 
 */




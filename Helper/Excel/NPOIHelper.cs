// ***********************************************************************
// Assembly         : NPOIHelper
// Author           : wangzhiming
// Created          : 03-14-2017
//
// Last Modified By : wangzhiming
// Last Modified On : 03-14-2017
// ***********************************************************************
// <copyright file="NPOIHelper.cs" company="">
//     Copyright ©  2017
// </copyright>
// <summary></summary>
// ***********************************************************************
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Reflection;

namespace NPOIHelper
{
    public struct CellLocation
    {
        public int SheetNumber;
        public int RowNumber;
        public int ColumnNumber;
    }
    /// <summary>
    /// Class NPOIHelper.
    /// </summary>
    public class NPOIHelper
    {
        /// <summary>
        /// 操作的Excel文件名（含路径）
        /// </summary>
        public string FileName { get; set; }
        public NPOI.HSSF.UserModel.HSSFWorkbook book2003 { get; set; }
        public NPOI.XSSF.UserModel.XSSFWorkbook book2007 { get; set; }
        
        /// <summary>
        /// 图片信息集合
        /// </summary>
        public List<PicturesInfo> PicList;

        /// <summary>
        /// 工作表
        /// </summary>
        public ISheet Sheet { get; set; }
        /// <summary>
        /// 工作表列表
        /// </summary>
        public List<ISheet> Sheets { get; set; }
        /// <summary>
        /// 获取或设置列头所在行号
        /// </summary>
        public int HeaderRow { get; set; } = -1;
        /// <summary>
        /// 数据首行
        /// </summary>
        public int FirstDataRow { get; set; } = 0;
        /// <summary>
        /// 数据行数
        /// </summary>
        public int DataRowCount
        {
            get;
            private set;
        }
        ICellStyle cellStyle1, cellStyle2;
        
        /// <summary>
        /// 带文件实例化
        /// </summary>
        /// <param name="_FileName"></param>
        public NPOIHelper(string _FileName)
        {
            Sheets = new List<ISheet>();
            FileName = _FileName;
            OpenExcelFile(FileName);
        }
        /// <summary>
        /// 
        /// </summary>
        public NPOIHelper()
        {
            Sheets = new List<ISheet>();
        }

        /// <summary>
        /// 打开一个Excel文件
        /// </summary>
        /// <param name="_FileName"></param>
        /// <returns></returns>
        public bool OpenExcelFile(string _FileName)
        {
            FileName = _FileName;
            return OpenExcelFile();
        }
        /// <summary>
        /// 打开构造函数中的Excel文件
        /// </summary>
        /// <returns></returns>
        public bool OpenExcelFile()
        {
            try
            {
                FileStream fs = new FileStream(FileName, FileMode.Open, FileAccess.ReadWrite);
                Sheets.Clear();
                string[] array = FileName.Split('.');
                if (array[array.Length - 1] == "xls")
                {
                    book2003 = new NPOI.HSSF.UserModel.HSSFWorkbook(fs);
                    for (int i = 0; i < book2003.NumberOfSheets; i++)
                    {
                        Sheets.Add(book2003.GetSheetAt(i));
                    }

                    cellStyle1 = book2003.CreateCellStyle();
                    cellStyle1.DataFormat = 0;
                    cellStyle2 = book2003.CreateCellStyle();
                    cellStyle2.DataFormat = 14;
                }
                else
                {
                    book2007 = new NPOI.XSSF.UserModel.XSSFWorkbook(fs);
                    for (int i = 0; i < book2007.NumberOfSheets; i++)
                    {
                        Sheets.Add(book2007.GetSheetAt(i));
                    }
                    cellStyle1 = book2007.CreateCellStyle();
                    cellStyle1.DataFormat = 0;
                    cellStyle2 = book2007.CreateCellStyle();
                    cellStyle2.DataFormat = 14;

                }
                if (Sheets.Count > 0)
                {
                    SetActiveSheet(0);
                    PicList = NPOIExtent.GetAllPictureInfos(Sheets[0]);
                }
            }
            catch
            {
                throw;
            }
            if (book2003 == null && book2007 == null) throw new Exception("book is null!");

            return true;
        }
        /// <summary>
        /// 设置当前活动工作表
        /// </summary>
        /// <param name="SheetName">工作表名称</param>
        /// <returns>true-匹配并设置，false-没有匹配的名称</returns>
        public bool SetActiveSheet(string SheetName)
        {
            for(int i = 0;i<Sheets.Count;i++)
            {
                if (Sheets[i].SheetName == SheetName)
                {
                    SetActiveSheet(i);
                    return true;
                }
            }
            return false;
        }
        /// <summary>
        /// 设置当前活动工作表
        /// </summary>
        /// <param name="SheetIndex">从0开始的工作表索引号</param>
        /// <returns></returns>
        public bool SetActiveSheet(int SheetIndex)
        {
            Sheet = Sheets[SheetIndex];
            DataRowCount = Sheet.PhysicalNumberOfRows - FirstDataRow;
            return true;
        }
        /// <summary>
        /// 在所有表中查找包含内容的单元格
        /// </summary>
        /// <param name="Content">查找要包含的内容</param>
        /// <returns>符合的集合</returns>
        public List<ICell> GetCellByContentInAllSheet(string Content)
        {
            List<ICell> cells = new List<ICell>();
            ISheet temp_sheet = Sheet;
            foreach (ISheet _sheet in Sheets)
            {
                Sheet = _sheet;
                cells.AddRange(GetCellByContent(Content));
            }
            Sheet = temp_sheet;
            if (book2003 == null && book2007 == null) throw new Exception("book is null!");

            return cells;
        }
        public List<CellLocation> GetCellLocationByContent(string Content)
        {
            List<CellLocation> clList = new List<CellLocation>();
            List<ICell> cells = GetCellByContent(Content);
            foreach (ICell  cell in cells)
            {
                CellLocation cl = new CellLocation();
                cl.ColumnNumber = cell.ColumnIndex;
                cl.RowNumber = cell.RowIndex;
                clList.Add(cl);
            }
            return clList;
        }
        /// <summary>
        /// 在当前活动工作表中查找包含内容的单元格
        /// </summary>
        /// <param name="Content"></param>
        /// <returns></returns>
        public List<ICell> GetCellByContent(string Content)
        {
            List<ICell> cells = new List<ICell>();
            for (int i = 0; i < Sheet.PhysicalNumberOfRows; i++)
            {
                ICell cell = GetCellByContent(i, Content);
                if(cell != null)
                    cells.Add(cell);
                
            }
            if (book2003 == null && book2007 == null) throw new Exception("book is null!");

            return cells;

        }
        public ICell GetCellByContent(int RowNum,string Content)
        {
            IRow row = Sheet.GetRow(RowNum);
            for (int j = 0; j <= row.PhysicalNumberOfCells; j++)
            {
                ICell cell = row.GetCell(j);
                if (cell != null)
                {
                    if (cell.ToString() == (Content))
                    {
                        return cell;
                    }
                }
            }
            return null;
        }
        public bool Save()
        {
            if (book2003 == null && book2007 == null) throw new Exception("book is null!");

            string[] array = FileName.Split('.');
            if (array[array.Length - 1] == "xls")
                Save(NPOIExtent.ExcelFormat.Excel2003);
            else
                Save(NPOIExtent.ExcelFormat.Excel2007);
            return true;
        }
        /// <summary>
        /// 保存到文件
        /// 
        /// 为防止保存出错，覆盖保存时会重命名原文件为.bak
        /// </summary>
        /// <returns>true - 成功</returns>
        public bool Save(NPOIExtent.ExcelFormat format)
        {
            if (File.Exists(FileName))
            {
                File.Copy(FileName, FileName + ".bak", true);
            }
            foreach(ISheet sheet in Sheets)
            {
                sheet.ForceFormulaRecalculation = true;
            }
            using (FileStream file = new FileStream(FileName, FileMode.Create, FileAccess.ReadWrite))
            {
                if (format == NPOIExtent.ExcelFormat.Excel2003)
                {
                    book2003.Write(file);
                }
                else
                {
                    book2007.Write(file);
                }
                file.Close();
            }
            return true;
        }
        /// <summary>
        /// 获取合并单元格的范围
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns>CellRangeAddress</returns>
        public NPOI.SS.Util.CellRangeAddress GetMergedCellRange(ICell cell)
        {
            for (int i = 0; i < Sheet.NumMergedRegions; i++)
            {
                NPOI.SS.Util.CellRangeAddress range = Sheet.GetMergedRegion(i);
                if (range.FirstRow == cell.RowIndex && range.FirstColumn == cell.ColumnIndex)
                {
                    return range;
                }
            }
            return null;
        }
        /// <summary>
        /// 设置列头行号
        /// </summary>
        /// <param name="headerString">列头中包含的字符串</param>
        /// <returns>true-找到列头行并设置,false-没找到行</returns>
        public bool SetHeaderRow(string headerString)
        {
            List<ICell> headerCells = GetCellByContent(headerString);
            if (headerCells.Count > 0)
            {
                Sheet = headerCells[0].Sheet;
                HeaderRow = headerCells[0].RowIndex;
                //如果列头单元格是合并单元格，设置数据首行加合并数
                if (headerCells[0].IsMergedCell)
                    FirstDataRow = GetMergedCellRange(headerCells[0]).LastRow + 1;
                else
                    FirstDataRow = HeaderRow + 1;
                DataRowCount = Sheet.PhysicalNumberOfRows - FirstDataRow;

                return true;
            }
            else
                return false;
        }
        /// <summary>
        /// 允许通过索引访问单元格(从列头开始)
        /// 已知可能问题：自定义日期时间类型单元获取时可能按数值型获取
        /// 原因是cellstyle.dataformat不能确定
        /// </summary>
        public object this[string SheetName, int DataRowNum, int CellNum]
        {
            
            get
            {
                for(int i =0;i<Sheets.Count;i++)
                {
                    if(Sheets[i].SheetName == SheetName)
                    {
                        return this[i, DataRowNum, CellNum];
                    }
                }
                return null;
            }
            set
            {
                for (int i = 0; i < Sheets.Count; i++)
                {
                    if (Sheets[i].SheetName == SheetName)
                    {
                        this[i, DataRowNum, CellNum] = value;
                    }
                }

            }
        }
        /// <summary>
        /// 允许通过索引访问单元格(从列头开始)
        /// </summary>
        /// <param name="SheetName"></param>
        /// <param name="DataRowNum"></param>
        /// <param name="CellName"></param>
        /// <returns></returns>
        public object this[string SheetName, int DataRowNum, string CellName]
        {
            get
            {
                try
                {
                    int CellNum = GetCellByContent(CellName)[0].ColumnIndex;
                    for (int i = 0; i < Sheets.Count; i++)
                    {
                        if (Sheets[i].SheetName == SheetName)
                        {
                            return this[i, DataRowNum, CellNum];
                        }
                    }
                }
                catch
                {
                    return null;
                }
                return null;
            }
            set
            {
                try
                {
                    int CellNum = GetCellByContent(CellName)[0].ColumnIndex;
                    for (int i = 0; i < Sheets.Count; i++)
                    {
                        if (Sheets[i].SheetName == SheetName)
                        {
                            this[i, DataRowNum, CellNum] = value;
                        }
                    }
                }
                catch( Exception ex)
                {
                    throw ex;
                }

            }
        }
        /// <summary>
        /// 允许通过索引访问单元格(从列头开始)
        /// </summary>
        /// <param name="SheetNum"></param>
        /// <param name="DataRowNum"></param>
        /// <param name="CellName"></param>
        /// <returns></returns>
        public object this[int SheetNum, int DataRowNum, string CellName]
        {
            get
            {
                object ob = null;
                try
                {
                    ICell cell = GetCellByContent(DataRowNum, CellName);
                    if(cell !=null)
                    { 
                        ob = this[SheetNum, DataRowNum, cell.ColumnIndex];
                    }
                }
                catch
                {
                    return null;
                }
                return ob;
            }
            set
            {
                try
                {
                    int CellNum = GetCellByContent(CellName)[0].ColumnIndex;
                    this[SheetNum, DataRowNum, CellNum] = value;
                }
                catch
                {
                    throw;
                }
            }
        }
        public void InsertRow(int rowNumber,int count)
        {
            Sheet.ShiftRows(rowNumber, this.Sheet.PhysicalNumberOfRows, count);
            //插入行时 形状同时下移
            if (Sheet is HSSFSheet)
            {
                var shapeContainer = Sheet.DrawingPatriarch as HSSFShapeContainer;
                if (null != shapeContainer)
                {
                    var shapeList = shapeContainer.Children;
                    foreach (var shape in shapeList)
                    {
                        if (shape is HSSFSimpleShape && shape.Anchor is HSSFClientAnchor && ((HSSFClientAnchor)shape.Anchor).Row1 >= rowNumber)
                        {
                            ((HSSFClientAnchor)shape.Anchor).Row1 += count;
                            ((HSSFClientAnchor)shape.Anchor).Row2 += count;
                        }
                    }
                }
            }
        }
        public void DeleteRow(int rowNumber)
        {
            Sheet.RemoveRow(Sheet.GetRow(rowNumber));

        }
        /// <summary>
        /// 允许通过索引访问单元格(从列头开始)
        /// </summary>
        /// <param name="SheetNum"></param>
        /// <param name="DataRowNum"></param>
        /// <param name="CellNum"></param>
        /// <param name="NewRow">总是插入新行</param>
        public object this[int SheetNum, int DataRowNum, int CellNum]
        {
            get
            {
                if (Sheets[SheetNum] == null) return null;
                if (Sheets[SheetNum].GetRow(DataRowNum + FirstDataRow) == null) return null;
                ICell cell = Sheets[SheetNum].GetRow(DataRowNum + FirstDataRow).GetCell(CellNum);
                if (cell == null) return null;


                switch (cell.CellType)
                {
                    case CellType.Boolean:
                        return cell.BooleanCellValue;
                    case CellType.Error:
                        return cell.ErrorCellValue;
                    case CellType.Formula:
                        return cell.CellFormula;
                    case CellType.Numeric:
                        if (DateUtil.IsCellDateFormatted(cell))
                            return cell.DateCellValue;

                        if ((cell.CellStyle.DataFormat > 13 && cell.CellStyle.DataFormat < 32) | (cell.CellStyle.DataFormat == 177 && cell.CellStyle.DataFormat == 178))
                        {
                            return cell.DateCellValue;
                        }
                        return cell.NumericCellValue;

                    default:
                        return cell.ToString();
                }
            }
            set
            {
                if (value == null) return;
                IRow row = Sheets[SheetNum].GetRow(DataRowNum + FirstDataRow);
                if (row == null ) row = Sheets[SheetNum].CreateRow(DataRowNum + FirstDataRow);
                ICell cell = row.GetCell(CellNum);
                if (cell == null)
                {
                    cell = row.CreateCell(CellNum);
                }
                switch (value.GetType().Name)
                {
                    case "Boolean":
                        cell.SetCellValue((bool)value);
                        break;
                    case "Int32":
                    case "Int16":
                        cell.SetCellValue((int)value);
                        break;
                    case "Decimal":
                    case "Single":
                    case "Double":
                        cell.SetCellValue((double)value);
                        break;
                    case "DateTime":
                        cell.SetCellValue((DateTime)value);
                        break;
                    default:
                        cell.SetCellValue(value.ToString());
                        break;
                        //throw new Exception("不能对空或未知类型单元格赋值");
                }
            }

        }
        /// <summary>
        /// 合并单元格.
        /// </summary>
        /// <param name="firstRow">The first row.</param>
        /// <param name="lastRow">The last row.</param>
        /// <param name="firstCol">The first col.</param>
        /// <param name="lastCol">The last col.</param>
        /// <param name="showLastValue">true:use right& bottom cell Value in top left cell.真值时使用右下单元格代替左上单元格值</param>
        public void MergedCell(int firstRow, int lastRow, int firstCol, int lastCol,bool showLastValue)
        {
            if (firstRow > -1 && lastRow > -1 && firstCol > -1 && lastCol > -1)
            {
                int MergedCount = Sheet.NumMergedRegions;
                int oldFirstRow = firstRow, oldFirstCol = firstCol;
                for (int i = MergedCount - 1; i >= 0; i--)
                {
                    if (Sheet.GetMergedRegion(i).FirstRow <= firstRow
                        && Sheet.GetMergedRegion(i).LastRow >= firstRow
                        && Sheet.GetMergedRegion(i).FirstColumn <= firstCol
                        && Sheet.GetMergedRegion(i).LastColumn >= firstCol)
                    {
                        oldFirstRow = Sheet.GetMergedRegion(i).FirstRow;
                        oldFirstCol = Sheet.GetMergedRegion(i).FirstColumn;
                        Sheet.RemoveMergedRegion(i);

                    }
                }
                if(showLastValue)
                {
                    this[0, oldFirstRow, oldFirstCol] = this[0, lastRow, lastCol];
                }
                Sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(oldFirstRow,
                    lastRow, oldFirstCol, lastCol));
            }
        }
        /// <summary>
        /// Reads the excel.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="FirstHeaderText">The  Header Row First Column Text</param>
        /// <returns>List&lt;T&gt;.</returns>
        /// <exception cref="Exception">行: + i.ToString() + ,列: + j.ToString() + ,值: + cell.ToString() +  :不是预期的类型： + item.PropertyType.Name</exception>
        /// <exception cref="System.Exception">行: + i.ToString() + ,列: + j.ToString() + ,值: + cell.ToString() +  :不是预期的类型： + item.PropertyType.Name</exception>
        public List<T> ReadExcel<T>(string FirstHeaderText)
        {
            List<T> list = new List<T>();
            System.Reflection.PropertyInfo[] properties;
            T t = default(T);
            t = Activator.CreateInstance<T>();
            properties = t.GetType().GetProperties();
            if (!SetHeaderRow(FirstHeaderText))
            {
                throw new Exception("Header Row Not Found!");
            }
            for (int i = FirstDataRow; i < Sheet.PhysicalNumberOfRows; i++)
            {
                foreach (System.Reflection.PropertyInfo item in properties)
                {
                    ICell cell = null;
                    //如果有列头
                    //看看列头是否包含属性名字
                    IRow row = Sheet.GetRow(HeaderRow);
                    foreach (ICell cel in row.Cells)
                    {
                        //有的话把该列作为取数据的单元格
                        if (cel.CellType != CellType.Blank && cel.CellType != CellType.Unknown)
                        {
                            if (cel.ToString().ToLower().Replace(" ", "").Contains(item.Name.ToLower()))
                            {
                                cell = Sheet.GetRow(i).GetCell(cel.ColumnIndex);
                                break;
                            }
                        }
                    }
                    if (cell != null)
                    {
                        if (item.PropertyType.Name == "DateTime")
                            cell.CellStyle = cellStyle2;
                        else
                            cell.CellStyle = cellStyle1;
                        try
                        {
                            object value = valueType(item.PropertyType, cell.ToString());
                            item.SetValue(t, value, null);
                        }
                        catch
                        {
                            throw new Exception("行:" + cell.RowIndex.ToString() + ",列:" + cell.ColumnIndex.ToString() + ",值:" + cell.ToString() + " :不是预期的类型：" + item.PropertyType.Name);
                        }
                    }
                }
                list.Add(t);
            }
            return list;
        }
        /// <summary>
        /// Values the type.
        /// </summary>
        /// <param name="t">The t.</param>
        /// <param name="value">The value.</param>
        /// <returns>System.Object.</returns>
        /// <exception cref="System.Exception">
        /// </exception>
        /// <exception cref="Exception"></exception>
        internal protected static object valueType(Type t, string value)
        {
            object o = null;
            try
            {
                switch (t.Name)
                {
                    case "Boolean":
                        o = Boolean.Parse(value);
                        break;
                    case "Decimal":
                        o = decimal.Parse(value);
                        break;
                    case "Int":
                    case "Int32":
                    case "Int64":
                        o = int.Parse(value);
                        break;
                    case "Single":
                    case "Float":
                        o = float.Parse(value);
                        break;
                    case "DateTime":
                        DateTime dt1;
                        if (DateTime.TryParseExact(value, "M/d/yy", null, DateTimeStyles.None, out dt1))
                            o = (object)dt1;
                        else
                        {
                            o = DateTime.MinValue;
                            throw new Exception();
                        }
                        break;
                    default:
                        o = value;
                        break;
                }
            }
            catch
            {
                throw new Exception(value + ":不是预期的类型：" + t.Name);
            }
            return o;
        }
    }
    /// <summary>
    /// NPOI Extent
    /// </summary>
    public static class NPOIExtent
        {
        /// <summary>
        /// Reads to exist DataTable from the excel file.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="FirstDateRow">FirstDateRow</param>
        /// <param name="dataTable">The data table.</param>
        /// <exception cref="System.Exception">行: + i.ToString() + ,列: + j.ToString() + ,值: + cell.ToString() +  :不是预期的类型： + dataTable.Columns[j].DataType.Name</exception>
        /// <exception cref="Exception">行: + i.ToString() + ,列: + j.ToString() + ,值: + cell.ToString() +  :不是预期的类型： + dataTable.Columns[j].DataType.Name</exception>
        public static void ReadExcel(System.Data.DataTable dataTable,string filePath, int sheetIndex, int FirstDateRow)
        {
            if (!string.IsNullOrEmpty(filePath))
            {
                NPOI.HSSF.UserModel.HSSFWorkbook book;
                NPOI.XSSF.UserModel.XSSFWorkbook book2007;
                NPOI.SS.UserModel.ISheet sheet;
                ICellStyle cellStyle1, cellStyle2;
                FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                try
                {
                    book = new NPOI.HSSF.UserModel.HSSFWorkbook(fs);
                    sheet = book.GetSheetAt(sheetIndex);
                    cellStyle1 = book.CreateCellStyle();
                    cellStyle1.DataFormat = 0;
                    cellStyle2 = book.CreateCellStyle();
                    cellStyle2.DataFormat = 14;
                }
                catch (NPOI.POIFS.FileSystem.OfficeXmlFileException)
                {
                    fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                    book2007 = new NPOI.XSSF.UserModel.XSSFWorkbook(fs);
                    sheet = book2007.GetSheetAt(sheetIndex);
                    cellStyle1 = book2007.CreateCellStyle();
                    cellStyle1.DataFormat = 0;
                    cellStyle2 = book2007.CreateCellStyle();
                    cellStyle2.DataFormat = 14;

                }
                for (int i = FirstDateRow; i < sheet.PhysicalNumberOfRows; i++)
                {
                    DataRow row = dataTable.NewRow();
                    NPOI.SS.UserModel.IRow iRow = sheet.GetRow(i);
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        NPOI.SS.UserModel.ICell cell = iRow.GetCell(j);
                        if (dataTable.Columns[j].DataType.Name == "DateTime")
                            cell.CellStyle = cellStyle2;
                        else
                            cell.CellStyle = cellStyle1;
                        try
                        {
                            object value = NPOIHelper.valueType(dataTable.Columns[j].DataType, cell.ToString());
                            row[j] = value;
                        }
                        catch
                        {
                            throw new Exception("行:" + i.ToString() + ",列:" + j.ToString() + ",值:" + cell.ToString() + " :不是预期的类型：" + dataTable.Columns[j].DataType.Name);
                        }

                    }
                    dataTable.Rows.Add(row);
                }

            }
        }
        /// <summary>
        /// Reads the excel.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="HeaderRowNum">HeaderRowNum, -1 is None</param>
        /// <returns>DataTable.</returns>
        public static DataTable ReadExcel(string filePath, int sheetIndex, int HeaderRowNum)
        {
            if (!string.IsNullOrEmpty(filePath))
            {
                NPOI.HSSF.UserModel.HSSFWorkbook book;
                NPOI.XSSF.UserModel.XSSFWorkbook book2007;
                NPOI.SS.UserModel.ISheet sheet;
                FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                try
                {
                    book = new NPOI.HSSF.UserModel.HSSFWorkbook(fs);
                    sheet = book.GetSheetAt(sheetIndex);
                }
                catch (NPOI.POIFS.FileSystem.OfficeXmlFileException)
                {
                    fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                    book2007 = new NPOI.XSSF.UserModel.XSSFWorkbook(fs);
                    sheet = book2007.GetSheetAt(sheetIndex);
                }

                DataTable dt = new DataTable(sheet.SheetName);
                if (HeaderRowNum < 0)
                {
                    foreach (NPOI.SS.UserModel.ICell cell in sheet.GetRow(0))
                    {
                        dt.Columns.Add("Column" + cell.ColumnIndex);
                    }
                }
                else
                {
                    foreach (NPOI.SS.UserModel.ICell cell in sheet.GetRow(HeaderRowNum))
                    {
                        dt.Columns.Add(cell.StringCellValue);
                    }

                }
                for (int i = HeaderRowNum + 1; i < sheet.PhysicalNumberOfRows; i++)
                {
                    DataRow row = dt.NewRow();
                    NPOI.SS.UserModel.IRow iRow = sheet.GetRow(i);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        NPOI.SS.UserModel.ICell cell = iRow.GetCell(j);
                        row[j] = cell.ToString();
                    }
                    dt.Rows.Add(row);
                }
                return dt;
            }
            return null;
        }
        /// <summary>
        /// Writes the excel.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">The list.</param>
        /// <param name="filePath">The file path.</param>
        public static void WriteExcel<T>(List<T> list, string filePath)
        {
            string[] filenamesplit = filePath.Split('.');
            if (filenamesplit[filenamesplit.Length -1].ToLower() == "xlsx")
                WriteExcel(list, filePath, ExcelFormat.Excel2007);
            else
                WriteExcel(list, filePath, ExcelFormat.Excel2003);

        }

        /// <summary>
        /// Writes the excel.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">The list.</param>
        /// <param name="filePath">The file path.</param>
        /// <param name="format">The format.</param>
        public static void WriteExcel<T>(List<T> list, string filePath, ExcelFormat format)
        {
            if (!string.IsNullOrEmpty(filePath) && null != list && list.Count > 0)
            {
                HSSFWorkbook book = null;
                XSSFWorkbook book1 = null;
                ISheet sheet;
                ICellStyle cellStyle1;
                ICellStyle cellStyle2;
                if (format == ExcelFormat.Excel2003)
                {
                    book = new HSSFWorkbook();
                    sheet = book.CreateSheet("Sheet1");
                    cellStyle1 = book.CreateCellStyle();
                    cellStyle1.DataFormat = 0;
                    cellStyle2 = book.CreateCellStyle();
                    cellStyle2.DataFormat = 14;

                }
                else
                {
                    book1 = new XSSFWorkbook();
                    sheet = book1.CreateSheet("Sheet1");
                    cellStyle1 = book1.CreateCellStyle();
                    cellStyle1.DataFormat = 0;
                    cellStyle2 = book1.CreateCellStyle();
                    cellStyle2.DataFormat = 14;

                }

                //创建属性的集合
                List<PropertyInfo> pList = new List<PropertyInfo>();
                //获得反射的入口

                Type type = typeof(T);
                //把所有的public属性加入到集合 并添加DataTable的列
                Array.ForEach<PropertyInfo>(type.GetProperties(), p =>
                {
                    pList.Add(p);
                });
                IRow row = sheet.CreateRow(0);
                for (int i = 0; i < pList.Count; i++)
                {
                    row.CreateCell(i).SetCellValue(pList[i].Name);
                }
                for (int i = 0; i < list.Count; i++)
                {
                    //创建一个DataRow实例
                    IRow row2 = sheet.CreateRow(i + 1);
                    //给row 赋值
                    for (int j = 0; j < pList.Count; j++)
                    {
                        ICell cell = row2.CreateCell(j);
                        cell.CellStyle = cellStyle1;
                        switch (pList[j].PropertyType.Name)
                        {
                            case "Int":
                            case "Int32":
                            case "Int64":
                                try
                                {
                                    cell.SetCellValue(int.Parse(pList[j].GetValue(list[i], null).ToString()));
                                }
                                catch
                                {
                                    //cell.SetCellValue();
                                }
                                break;
                            case "Single":
                            case "Double":
                            case "Float":
                            case "Decimal":
                                try
                                {
                                    cell.SetCellValue(double.Parse(pList[j].GetValue(list[i], null).ToString()));
                                }
                                catch
                                {
                                    //cell.SetCellValue("格式错误");
                                }
                                break;
                            case "Boolean":
                                cell.SetCellType(NPOI.SS.UserModel.CellType.Boolean);
                                try
                                {
                                    cell.SetCellValue(bool.Parse(pList[j].GetValue(list[i], null).ToString()));
                                }
                                catch
                                {
                                    //cell.SetCellValue("格式错误");
                                }
                                break;
                            case "DateTime":
                                cell.CellStyle = cellStyle2;
                                try
                                {
                                    cell.SetCellValue(DateTime.Parse(pList[j].GetValue(list[i], null).ToString()));
                                }
                                catch
                                {
                                    //cell.SetCellValue("格式错误");
                                }
                                break;
                            default:
                                try
                                {
                                    row2.CreateCell(j).SetCellValue(pList[j].GetValue(list[i], null).ToString());
                                }
                                catch
                                {

                                }
                                break;
                        }


                    }
                }
                // 写入到客户端  
                using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite))
                {
                    if (format == ExcelFormat.Excel2003)
                    {
                        book.Write(file);
                    }
                    else
                    {
                        book1.Write(file);
                    }
                    file.Close();
                }
            }
        }
        /// <summary>
        /// Enum ExcelFormat
        /// </summary>
        public enum ExcelFormat
        {
            /// <summary>
            /// The excel2003
            /// </summary>
            Excel2003,
            /// <summary>
            /// The excel2007
            /// </summary>
            Excel2007
        }

        /// <summary>
        /// Writes the excel.
        /// </summary>
        /// <param name="dt">The dt.</param>
        /// <param name="filePath">The file path.</param>
        public static void WriteExcel(DataTable dt, string filePath)
        {
            string[] array = filePath.Split('.');
            if (array[array.Length - 1] == "xls")
                WriteExcel(dt, filePath, ExcelFormat.Excel2003);
            else
                WriteExcel(dt, filePath, ExcelFormat.Excel2007);

        }
        /// <summary>
        /// Writes the excel.
        /// </summary>
        /// <param name="dt">The dt.</param>
        /// <param name="filePath">The file path.</param>
        /// <param name="format">The format.</param>
        public static void WriteExcel(DataTable dt, string filePath, ExcelFormat format)
        {
            if (!string.IsNullOrEmpty(filePath) && null != dt && dt.Rows.Count > 0)
            {
                HSSFWorkbook book = null;
                XSSFWorkbook book1 = null;
                ISheet sheet;
                ICellStyle cellStyle1;
                ICellStyle cellStyle2;
                if (dt.TableName == "") dt.TableName = "Sheet1";
                if (format == ExcelFormat.Excel2003)
                {
                    book = new HSSFWorkbook();
                    sheet = book.CreateSheet(dt.TableName);
                    cellStyle1 = book.CreateCellStyle();
                    cellStyle2 = book.CreateCellStyle();
                }
                else
                {
                    book1 = new XSSFWorkbook();
                    sheet = book1.CreateSheet(dt.TableName);
                    cellStyle1 = book1.CreateCellStyle();
                    cellStyle2 = book1.CreateCellStyle();

                }

                IRow row = sheet.CreateRow(0);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row.CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    NPOI.SS.UserModel.IRow row2 = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ICell cell = row2.CreateCell(j);
                        if ((dt.Rows[i][j]) == DBNull.Value) continue;
                        switch (dt.Columns[j].DataType.ToString())
                        {
                            case "System.Int":
                            case "System.Int32":
                            case "System.Int64":
                                cellStyle1.DataFormat = 0;
                                cell.CellStyle = cellStyle1;
                                try
                                {
                                    cell.SetCellValue(int.Parse(dt.Rows[i][j].ToString()));
                                }
                                catch
                                {
                                    cell.SetCellValue("格式错误");
                                }
                                break;
                            case "System.Single":
                            case "System.Double":
                            case "System.Decimal":
                                cellStyle1.DataFormat = 0;
                                cell.CellStyle = cellStyle1;
                                try
                                {
                                    cell.SetCellValue(double.Parse(dt.Rows[i][j].ToString()));
                                }
                                catch
                                {
                                    cell.SetCellValue("格式错误");
                                }
                                break;
                            case "System.Boolean":
                                //cell.SetCellType(NPOI.SS.UserModel.CellType.Boolean);
                                cell.SetCellValue((bool)((dt.Rows[i][j]) ?? false));
                                break;
                            case "System.DateTime":
                                cellStyle2.DataFormat = 14;
                                cell.CellStyle = cellStyle2;
                                try
                                {
                                    cell.SetCellValue(DateTime.Parse(dt.Rows[i][j].ToString()));
                                }
                                catch
                                {
                                    cell.SetCellValue("格式错误");
                                }
                                break;
                            default:
                                row2.CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
                                break;
                        }
                    }
                }
                // 写入到客户端  
                using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite))
                {
                    if (format == ExcelFormat.Excel2003)
                    {
                        book.Write(file);
                    }
                    else
                    {
                        book1.Write(file);
                    }
                    file.Close();
                }
            }
        }

        /// <summary>
        /// Gets all picture infos.
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        /// <returns>List&lt;PicturesInfo&gt;.</returns>
        public static List<PicturesInfo> GetAllPictureInfos(ISheet sheet)
        {
            return GetAllPictureInfos(sheet,null, null, null, null);
        }

        /// <summary>
        /// Gets all picture infos.
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        /// <param name="minRow">The minimum row.</param>
        /// <param name="maxRow">The maximum row.</param>
        /// <param name="minCol">The minimum col.</param>
        /// <param name="maxCol">The maximum col.</param>
        /// <param name="onlyInternal">if set to <c>true</c> [only internal].</param>
        /// <returns>List&lt;PicturesInfo&gt;.</returns>
        private static List<PicturesInfo> GetAllPictureInfos(HSSFSheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol, bool onlyInternal)
        {
            List<PicturesInfo> picturesInfoList = new List<PicturesInfo>();

            var shapeContainer = sheet.DrawingPatriarch as HSSFShapeContainer;
            if (null != shapeContainer)
            {
                var shapeList = shapeContainer.Children;
                foreach (var shape in shapeList)
                {
                    if (shape is HSSFPicture && shape.Anchor is HSSFClientAnchor)
                    {
                        var picture = (HSSFPicture)shape;
                        var anchor = (HSSFClientAnchor)shape.Anchor;

                        if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, onlyInternal))
                        {
                            picturesInfoList.Add(new PicturesInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, picture.PictureData.Data));
                        }
                    }
                }
            }

            return picturesInfoList;
        }

        /// <summary>
        /// Gets all picture infos.
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        /// <param name="minRow">The minimum row.</param>
        /// <param name="maxRow">The maximum row.</param>
        /// <param name="minCol">The minimum col.</param>
        /// <param name="maxCol">The maximum col.</param>
        /// <param name="onlyInternal">if set to <c>true</c> [only internal].</param>
        /// <returns>List&lt;PicturesInfo&gt;.</returns>
        /// <exception cref="System.Exception">未处理类型，没有为该类型添加：GetAllPicturesInfos()扩展方法！</exception>
        public static List<PicturesInfo> GetAllPictureInfos(ISheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol, bool onlyInternal = true)
        {
            if (sheet is HSSFSheet)
            {
                return GetAllPictureInfos((HSSFSheet)sheet, minRow, maxRow, minCol, maxCol, onlyInternal);
            }
            else if (sheet is XSSFSheet)
            {
                return GetAllPictureInfos((XSSFSheet)sheet, minRow, maxRow, minCol, maxCol, onlyInternal);
            }
            else
            {
                throw new Exception("未处理类型，没有为该类型添加：GetAllPicturesInfos()扩展方法！");
            }
        }
        /// <summary>
        /// Gets all picture infos.
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        /// <param name="minRow">The minimum row.</param>
        /// <param name="maxRow">The maximum row.</param>
        /// <param name="minCol">The minimum col.</param>
        /// <param name="maxCol">The maximum col.</param>
        /// <param name="onlyInternal">if set to <c>true</c> [only internal].</param>
        /// <returns>List&lt;PicturesInfo&gt;.</returns>
        private static List<PicturesInfo> GetAllPictureInfos(XSSFSheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol, bool onlyInternal)
        {
            List<PicturesInfo> picturesInfoList = new List<PicturesInfo>();
            try
            {
                XSSFDrawing drawing = sheet.GetDrawingPatriarch();
                var shapeList = drawing.GetShapes();
                foreach (XSSFPicture shape in shapeList)
                {
                    XSSFPicture picture = (XSSFPicture)shape;
                    IClientAnchor anchor = picture.ClientAnchor;

                    if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, onlyInternal))
                    {
                        picturesInfoList.Add(new PicturesInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, picture.PictureData.Data));
                    }
                }
            }
            catch
            {
                return null;
            }
            return picturesInfoList;
        }

        /// <summary>
        /// Determines whether [is internal or intersect] [the specified range minimum row].
        /// </summary>
        /// <param name="rangeMinRow">The range minimum row.</param>
        /// <param name="rangeMaxRow">The range maximum row.</param>
        /// <param name="rangeMinCol">The range minimum col.</param>
        /// <param name="rangeMaxCol">The range maximum col.</param>
        /// <param name="pictureMinRow">The picture minimum row.</param>
        /// <param name="pictureMaxRow">The picture maximum row.</param>
        /// <param name="pictureMinCol">The picture minimum col.</param>
        /// <param name="pictureMaxCol">The picture maximum col.</param>
        /// <param name="onlyInternal">if set to <c>true</c> [only internal].</param>
        /// <returns><c>true</c> if [is internal or intersect] [the specified range minimum row]; otherwise, <c>false</c>.</returns>
        private static bool IsInternalOrIntersect(int? rangeMinRow, int? rangeMaxRow, int? rangeMinCol, int? rangeMaxCol,
            int pictureMinRow, int pictureMaxRow, int pictureMinCol, int pictureMaxCol, bool onlyInternal)
        {
            int _rangeMinRow = rangeMinRow ?? pictureMinRow;
            int _rangeMaxRow = rangeMaxRow ?? pictureMaxRow;
            int _rangeMinCol = rangeMinCol ?? pictureMinCol;
            int _rangeMaxCol = rangeMaxCol ?? pictureMaxCol;

            if (onlyInternal)
            {
                return (_rangeMinRow <= pictureMinRow && _rangeMaxRow >= pictureMaxRow &&
                        _rangeMinCol <= pictureMinCol && _rangeMaxCol >= pictureMaxCol);
            }
            else
            {
                return ((Math.Abs(_rangeMaxRow - _rangeMinRow) + Math.Abs(pictureMaxRow - pictureMinRow) >= Math.Abs(_rangeMaxRow + _rangeMinRow - pictureMaxRow - pictureMinRow)) &&
                (Math.Abs(_rangeMaxCol - _rangeMinCol) + Math.Abs(pictureMaxCol - pictureMinCol) >= Math.Abs(_rangeMaxCol + _rangeMinCol - pictureMaxCol - pictureMinCol)));
            }
        }
    }
    /// <summary>
    /// Class PicturesInfo.
    /// </summary>
    public class PicturesInfo
    {
        /// <summary>
        /// Gets or sets the minimum row.
        /// </summary>
        /// <value>The minimum row.</value>
        public int MinRow { get; set; }
        /// <summary>
        /// Gets or sets the maximum row.
        /// </summary>
        /// <value>The maximum row.</value>
        public int MaxRow { get; set; }
        /// <summary>
        /// Gets or sets the minimum col.
        /// </summary>
        /// <value>The minimum col.</value>
        public int MinCol { get; set; }
        /// <summary>
        /// Gets or sets the maximum col.
        /// </summary>
        /// <value>The maximum col.</value>
        public int MaxCol { get; set; }
        /// <summary>
        /// Gets the picture data.
        /// </summary>
        /// <value>The picture data.</value>
        public Byte[] PictureData { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="PicturesInfo"/> class.
        /// </summary>
        /// <param name="minRow">The minimum row.</param>
        /// <param name="maxRow">The maximum row.</param>
        /// <param name="minCol">The minimum col.</param>
        /// <param name="maxCol">The maximum col.</param>
        /// <param name="pictureData">The picture data.</param>
        public PicturesInfo(int minRow, int maxRow, int minCol, int maxCol, Byte[] pictureData)
        {
            this.MinRow = minRow;
            this.MaxRow = maxRow;
            this.MinCol = minCol;
            this.MaxCol = maxCol;
            this.PictureData = pictureData;
        }
    }
}

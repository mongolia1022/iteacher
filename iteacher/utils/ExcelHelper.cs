using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace JbkConsole
{
    /// <summary>
    /// Excel导入导出 ，基于NOPI2.1.3版本
    /// </summary>
    public class ExcelHelper
    {
        public bool istie;
        private IWorkbook _workBook;
        //文件名
        private string _fileName = DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

        /// <summary>
        /// 由DataTable导出Excel
        /// </summary>
        /// <param name="fileName"> 文件名 </param>
        /// <param name="table"> 数据表 </param>
        /// <param name="hasTitle"> 第一行是否是标题 </param>
        /// <returns></returns>
        public ExcelHelper DataTableToExcel(string fileName, DataTable table, bool hasTitle)
        {
            _fileName = fileName;
            if (fileName.Contains(".xlsx"))
            {
                _workBook = new XSSFWorkbook();
            }
            else if (fileName.Contains(".xls"))
            {
                _workBook = new HSSFWorkbook();
            }
            if (table == null) return this;
            var rowsCount = table.Rows.Count;
            if (rowsCount == 0) return this;
            ISheet sheet = InitSheetTitle(table, "Sheet1", hasTitle);
            var bodyStyle = BodyStyle();
            var dateStyle = BodyStyle();
            var formate = _workBook.CreateDataFormat();
            dateStyle.DataFormat = formate.GetFormat("yyyy-Mm-dd HH:mm:ss");

            var rows = table.Rows;
            //行计数
            var count = 1;
            var sheetCount = 1;
            foreach (DataRow row in rows)
            {
                try
                {
                    if (count > 65534 && !fileName.Contains(".xlsx"))
                    {
                        sheetCount++;
                        sheet = InitSheetTitle(table, "Sheet" + sheetCount, hasTitle);
                        count = 1;
                    }
                    IRow newRow = sheet.CreateRow(count);
                    InitRowData(newRow, row, bodyStyle, dateStyle);
                    count++;
                    //Console.WriteLine(count);
                }
                catch (Exception ex)
                {
                    //Console.WriteLine(ex.StackTrace);
                }
            }
            return this;
        }

        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="fileName"> 文件名 </param>
        /// <param name="tables"> table列表 </param>
        /// <param name="firstIsTitle"> DataTable的列名是否要导入 </param>
        /// <returns> 导入数据行数(包含列名那一行) </returns>
        public ExcelHelper DataTableToExcelWithMultSheet(string fileName, Dictionary<string,DataTable> tables, bool firstIsTitle)
        {
            if (tables.Count == 0) return this;
            if (fileName.Contains(".xlsx")) // 2007版本
                _workBook = new XSSFWorkbook();
            else if (fileName.Contains(".xls")) // 2003版本
                _workBook = new HSSFWorkbook();
            try
            {
                var headStyle = TitleStyle();
                var bodyStyle = BodyStyle();
                foreach (var p in tables)
                {
                    ISheet sheet;
                    if (_workBook != null)
                    {
                        sheet = InitSheetTitle(p.Value, p.Key, true);
                    }
                    else
                    {
                        return this;
                    }
                    int count;
                    int j;
                    if (firstIsTitle) //写入DataTable的列名
                    {
                        IRow row = sheet.CreateRow(0);
                        for (j = 0; j < p.Value.Columns.Count; ++j)
                        {
                            var cell = row.CreateCell(j);
                            cell.CellStyle = headStyle;
                            cell.SetCellValue(p.Value.Columns[j].ColumnName);
                        }
                        count = 1;
                    }
                    else
                    {
                        count = 0;
                    }

                    int i;
                    for (i = 0; i < p.Value.Rows.Count; ++i)
                    {
                        IRow row = sheet.CreateRow(count);
                        for (j = 0; j < p.Value.Columns.Count; ++j)
                        {
                            var cell = row.CreateCell(j);
                            cell.CellStyle = bodyStyle;
                            cell.SetCellValue(p.Value.Rows[i][j].ToString());
                        }
                        ++count;
                    }
                }
                _fileName = fileName;
                return this;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return this;
            }
            finally
            {
                foreach (var table in tables)
                {
                    table.Value.Dispose();
                }
            }
        }

        /// <summary>
        /// 由对象列表导出Excel
        /// </summary>
        /// <param name="fileName"> 文件名 </param>
        /// <param name="list"> 数据列表 </param>
        /// <param name="config"> 标题配置 </param>
        /// <returns></returns>
        public ExcelHelper ObjectToExcel(string fileName, List<object> list, List<OutPutConfig> config)
        {
            _fileName = fileName;
            if (fileName.Contains(".xlsx"))
            {
                _workBook = new XSSFWorkbook();
            }
            else if (fileName.Contains(".xls"))
            {
                _workBook = new HSSFWorkbook();
            }
            if (_workBook == null) return this;
            if (list == null) return this;
            var rowsCount = list.Count;
            if (rowsCount == 0) return this;
            ISheet sheet = InitSheetTitle(config, "Sheet1");
            var bodyStyle = BodyStyle();
            var dateStyle = BodyStyle();
            var formate = _workBook.CreateDataFormat();
            dateStyle.DataFormat = formate.GetFormat("yyyy-Mm-dd HH:mm:ss");
            //行计数
            var count = 1;
            var sheetCount = 1;
            Type type = null;
            for (int i = 0; i < rowsCount; i++)
            {
                if (list[i] != null)
                {
                    type = list[i].GetType();
                    break;
                }
            }
            foreach (var item in list)
            {
                try
                {
                    if (item == null) continue;
                    if (count > 65534 && !fileName.Contains(".xlsx"))
                    {
                        sheetCount++;
                        sheet = InitSheetTitle(config, "Sheet" + sheetCount);
                        count = 1;
                    }
                    IRow newRow = sheet.CreateRow(count);
                    InitRowData(newRow, item, config, type, bodyStyle, dateStyle);
                    count++;
                    Console.WriteLine(count);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.StackTrace);
                }
            }
            return this;
        }

        /// <summary>
        /// 执行导出本地导出
        /// </summary>
        public void ExcuteLocal()
        {
            if (_workBook == null) return;
            var fs = new FileStream(_fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            try
            {
                _workBook.Write(fs);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                fs.Close();
                fs.Dispose();
            }

        }

        ///// <summary>
        ///// 执行Web导出
        ///// </summary>
        //public void ExcuteWeb()
        //{
        //    if (_workBook == null) return;
        //    var file = new MemoryStream();
        //    try
        //    {
        //        _workBook.Write(file);
        //        HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
        //        HttpContext.Current.Response.AddHeader("Content-Disposition",
        //            string.Format("attachment;filename= {0}", _fileName));
        //        HttpContext.Current.Response.Clear();
        //        HttpContext.Current.Response.BinaryWrite(file.GetBuffer());
        //        HttpContext.Current.Response.End();
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine(ex.Message);
        //    }
        //    finally
        //    {
        //        file.Close();
        //        file.Dispose();
        //    }
        //}

        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="sheetName"> excel工作薄sheet的名称 </param>
        /// <param name="hasTitle"> 第一行是否是列名 </param>
        /// <returns> 返回的DataTable </returns>
        public DataTable ExcelToDataTable(string fileName, string sheetName, bool hasTitle)
        {
            var data = new DataTable();
            var fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            IWorkbook workbook = null;
            if (fileName.Contains(".xlsx"))
            {
                workbook = new XSSFWorkbook(fs);
            }
            else if (fileName.Contains(".xls"))
            {
                workbook = new HSSFWorkbook(fs);
            }
            if (workbook == null) return new DataTable();
            try
            {
                var sheet = !string.IsNullOrEmpty(sheetName)
                    ? (workbook.GetSheet(sheetName) ?? workbook.GetSheetAt(0))
                    : workbook.GetSheetAt(0);
                if (sheet == null) return data;

                data = GetTable(sheet, hasTitle);
                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return new DataTable();
            }
            finally
            {
                fs.Close();
            }
        }

        /// <summary>
        /// 获取Excel中所有工作区间
        /// </summary>
        /// <param name="fileName"> 文件名 </param>
        /// <param name="hasTitle"> 是否含有标题 </param>
        /// <returns></returns>
        public List<DataTable> ExcelToDataTable(string fileName, bool hasTitle)
        {
            if (string.IsNullOrEmpty(fileName)) return new List<DataTable>();
            var list = new List<DataTable>();
            var fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.Read);
            IWorkbook workbook = null;
            if (fileName.Contains(".xlsx"))
            {
                workbook = new XSSFWorkbook(fs);
            }
            else if (fileName.Contains(".xls"))
            {
                workbook = new HSSFWorkbook(fs);
            }
            if (workbook == null) return new List<DataTable>();
            int count = workbook.NumberOfSheets;
            for (int i = 0; i < count; i++)
            {
                var table = GetTable(workbook.GetSheetAt(i), hasTitle);
                if (table == null) continue;
                list.Add(table);
            }
            return list;
        }

        //获取DataTable
        private DataTable GetTable(ISheet sheet, bool hasTitle)
        {
            var table = new DataTable();
            IRow firstRow = sheet.GetRow(0);
            int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

            int startRow;
            if (hasTitle)
            {
                for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                {
                    ICell cell = firstRow.GetCell(i);
                    if (cell == null) continue;
                    var cellValue = cell.StringCellValue;
                    if (cellValue == null) continue;
                    var column = new DataColumn(cellValue);
                    table.Columns.Add(column);
                }
                startRow = sheet.FirstRowNum + 1;
            }
            else
            {
                startRow = sheet.FirstRowNum;
            }

            //最后一列的标号
            int rowCount = sheet.LastRowNum;
            for (int i = startRow; i <= rowCount; ++i)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue; //没有数据的行默认是null　　　　　　　

                DataRow dataRow = table.NewRow();
                for (int j = row.FirstCellNum; j < cellCount; ++j)
                {
                    if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                        dataRow[j] = row.GetCell(j).ToString();
                }
                table.Rows.Add(dataRow);
            }
            return table;
        }

        //初始化行
        private void InitRowData(IRow newRow, object obj, List<OutPutConfig> config,
            Type type, ICellStyle style, ICellStyle dateStyle)
        {
            for (int i = 0, len = config.Count; i < len; i++)
            {
                var val = type.GetProperty(config[i].FieldName).GetValue(obj, null);
                SetCellValue(newRow, val, i, style, dateStyle);
            }
        }

        //初始化行
        private void InitRowData(IRow newRow, DataRow oldRow, ICellStyle style, ICellStyle dateStyle)
        {
            var cols = oldRow.ItemArray;
            for (int i = 0, len = cols.Count(); i < len; i++)
            {
                if (cols[i] == null) continue;
                SetCellValue(newRow, cols[i], i, style, dateStyle);
            }
        }

        //设置单元格的值
        private void SetCellValue(IRow newRow, object obj, int i, ICellStyle style, ICellStyle dateStyle)
        {
            ICell cell;
            switch (obj.GetType().ToString())
            {
                case "System.DateTime":
                    cell = newRow.CreateCell(i);
                    cell.SetCellValue(Convert.ToDateTime(obj));
                    cell.CellStyle = dateStyle;
                    break;
                case "System.Boolean":
                    cell = newRow.CreateCell(i, CellType.Numeric);
                    cell.SetCellValue(Convert.ToBoolean(obj));
                    cell.CellStyle = style;
                    break;
                case "System.Int16"://整型  
                case "System.Int32":
                case "System.Int64":
                case "System.Byte":
                    cell = newRow.CreateCell(i, CellType.Numeric);
                    cell.SetCellValue(Convert.ToInt64(obj));
                    cell.CellStyle = style;
                    break;
                case "System.Decimal"://浮点型  
                case "System.Double":
                    cell = newRow.CreateCell(i, CellType.Numeric);
                    cell.SetCellValue(Convert.ToDouble(obj));
                    cell.CellStyle = style;
                    break;
                default:
                    cell = newRow.CreateCell(i, CellType.String);
                    cell.SetCellValue(obj.ToString());
                    cell.CellStyle = style;
                    break;
            }
        }

        //初始化标题
        private ISheet InitSheetTitle(DataTable table, string sheet, bool hasTitle)
        {
            ISheet sheet1 = _workBook.CreateSheet(sheet);
            IRow newRow = sheet1.CreateRow(0);
            var cols = table.Columns;
            var oldRow = table.Rows[0];
            var colsCount = table.Columns.Count;
            var titleStyle = TitleStyle();
            ICell cell;
            for (int i = 0; i < colsCount; i++)
            {
                sheet1.SetColumnWidth(i, 20 * 300);
                if (!hasTitle)
                {
                    cell = newRow.CreateCell(i);
                    cell.SetCellValue(!string.IsNullOrEmpty(cols[i].Caption) ? cols[i].Caption : cols[i].ColumnName);
                    cell.CellStyle = titleStyle;
                    continue;
                }

                //cell = HSSFCellUtil.CreateCell(newRow, i, );
                cell = newRow.CreateCell(i);
                cell.SetCellValue(oldRow[i] != null ? oldRow[i].ToString() : "");
                cell.CellStyle = titleStyle;
            }
            return sheet1;
        }

        //初始化标题
        private ISheet InitSheetTitle(List<OutPutConfig> config, string sheet)
        {
            ISheet sheet1 = _workBook.CreateSheet(sheet);
            IRow newRow = sheet1.CreateRow(0);
            var colsCount = config.Count;
            var titleStyle = TitleStyle();
            ICell cell;
            for (int i = 0; i < colsCount; i++)
            {
                sheet1.SetColumnWidth(i, 20 * 300);

                //cell = HSSFCellUtil.CreateCell(newRow, i, config[i].DisplayName);
                cell = newRow.CreateCell(i);
                cell.SetCellValue(config[i].DisplayName);
                cell.CellStyle = titleStyle;
            }
            return sheet1;
        }

        //初始化内容样式
        private ICellStyle BodyStyle()
        {
            IFont bodyFont = _workBook.CreateFont();
            //bodyFont.Color = HSSFColor.OLIVE_GREEN.index;
            bodyFont.Color = HSSFColor.OliveGreen.Index;
            bodyFont.Boldweight = (short)FontBoldWeight.Normal; //设置粗体
            bodyFont.FontHeightInPoints = 12;
            bodyFont.FontName = "宋体";
            bodyFont.IsStrikeout = false;
            var bodyStyle = _workBook.CreateCellStyle();
            bodyStyle.SetFont(bodyFont);
            //边框
            bodyStyle.BorderBottom = BorderStyle.Thin;
            bodyStyle.BorderLeft = BorderStyle.Thin;
            bodyStyle.BorderRight = BorderStyle.Thin;
            bodyStyle.BorderTop = BorderStyle.Thin;
            //居中
            bodyStyle.Alignment = HorizontalAlignment.Left;
            return bodyStyle;
        }

        //初始化标题样式
        private ICellStyle TitleStyle()
        {
            IFont titleFont = _workBook.CreateFont();
            titleFont.Color = HSSFColor.OliveGreen.Index;
            titleFont.Boldweight = (short)FontBoldWeight.Bold; //设置粗体
            titleFont.FontHeightInPoints = 12;
            titleFont.FontName = "宋体";
            titleFont.IsStrikeout = false;

            ICellStyle titleStyle = _workBook.CreateCellStyle();
            titleStyle.SetFont(titleFont);
            //边框
            titleStyle.BorderBottom = BorderStyle.Thin;
            titleStyle.BorderLeft = BorderStyle.Thin;
            titleStyle.BorderRight = BorderStyle.Thin;
            titleStyle.BorderTop = BorderStyle.Thin;
            //居中
            titleStyle.Alignment = HorizontalAlignment.Center;
            return titleStyle;
        }
    }

    /// <summary>
    /// 导出配置类
    /// </summary>
    public class OutPutConfig
    {
        public string FieldName { get; set; }

        public string DisplayName { get; set; }
    }
}
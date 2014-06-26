using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Script.Serialization;

namespace ExportXLSFromDataTable
{
    public class ExportDataTable
    {
        public static Stream RenderDataTableToExcelByFrozen(DataTable SourceTable, int FrozenRow = 0, int FrozenColumn = 0)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            MemoryStream ms = new MemoryStream();
            HSSFSheet sheet = workbook.CreateSheet() as HSSFSheet;
            sheet.CreateFreezePane(FrozenColumn, FrozenRow);

            #region 哈哈
            //foreach (var item in collection)
            //{
            //    HSSFRow headerRow = sheet.CreateRow(0) as HSSFRow;
            //    ICellStyle headStyle = workbook.CreateCellStyle();
            //    headStyle.Alignment = HorizontalAlignment.CENTER;
            //    HSSFFont font = workbook.CreateFont() as HSSFFont;
            //    font.FontHeightInPoints = 10;
            //    font.Boldweight = 700;
            //    headStyle.SetFont(font); 
            //}

            // handling header.
            //foreach (DataColumn column in SourceTable.Columns)
            //{
            //    headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
            //    headerRow.GetCell(column.Ordinal).CellStyle = headStyle;
            //}

            // handling value.
            //int rowIndex = 1;
            //foreach (DataRow row in SourceTable.Rows)
            //{
            //    HSSFRow dataRow = sheet.CreateRow(rowIndex) as HSSFRow;

            //    foreach (DataColumn column in SourceTable.Columns)
            //    {
            //        dataRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
            //    }

            //    rowIndex++;
            //}
            #endregion
            workbook.Write(ms);
            ms.Flush();
            ms.Position = 0;

            sheet = null;
            //headerRow = null;
            workbook = null;

            return ms;
        }

        public static Stream DataSource(int FrozenColumn = 2, int FrozenRow = 2)
        {
            string jsonString = @"{'Columns':[[{'field':'OrgName','title':'单位','width':100,'align':'center','colspan':1,'rowspan':2,'hidden':false,'editor':'text'},{'field':'DeptName','title':'部门','width':100,'align':'center','colspan':1,'rowspan':2,'hidden':false,'editor':'text'},{'field':null,'title':'11-30','width':80,'align':'center','colspan':2,'rowspan':1,'hidden':false,'editor':'text'},{'field':null,'title':'31-40','width':80,'align':'center','colspan':2,'rowspan':1,'hidden':false,'editor':'text'},{'field':null,'title':'41-50','width':80,'align':'center','colspan':2,'rowspan':1,'hidden':false,'editor':'text'},{'field':null,'title':'51-60','width':80,'align':'center','colspan':2,'rowspan':1,'hidden':false,'editor':'text'},{'field':null,'title':'61-70','width':80,'align':'center','colspan':2,'rowspan':1,'hidden':false,'editor':'text'},{'field':null,'title':'71-80','width':80,'align':'center','colspan':2,'rowspan':1,'hidden':false,'editor':'text'},{'field':null,'title':'81-90','width':80,'align':'center','colspan':2,'rowspan':1,'hidden':false,'editor':'text'},{'field':null,'title':'101-110','width':80,'align':'center','colspan':2,'rowspan':1,'hidden':false,'editor':'text'},{'field':null,'title':'未知年龄','width':100,'align':'center','colspan':2,'rowspan':1,'hidden':false,'editor':'text'}],[{'field':'Age30','title':'人数','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'},{'field':'Percent30','title':'比例','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'},{'field':'Age40','title':'人数','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'},{'field':'Percent40','title':'比例','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'},{'field':'Age50','title':'人数','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'},{'field':'Percent50','title':'比例','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'},{'field':'Age60','title':'人数','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'},{'field':'Percent60','title':'比例','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'},{'field':'Age70','title':'人数','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'},{'field':'Percent70','title':'比例','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'},{'field':'Age80','title':'人数','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'},{'field':'Percent80','title':'比例','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'},{'field':'Age90','title':'人数','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'},{'field':'Percent90','title':'比例','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'},{'field':'Age110','title':'人数','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'},{'field':'Percent110','title':'比例','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'},{'field':'Undefinde','title':'人数','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'},{'field':'Percent','title':'比例','width':50,'align':'center','colspan':1,'rowspan':1,'hidden':false,'editor':'text'}]],'jsonData':{'rows':[{'OrgName':'铁峰煤业有限公司','DeptName':'公司部室','Age30':1,'Percent30':'100.00%','Age40':4,'Percent40':'100.00%','Age50':1,'Percent50':'100.00%','Age60':0,'Percent60':'0.00%','Age70':0,'Percent70':'0.00%','Age80':0,'Percent80':'0.00%','Age90':0,'Percent90':'0.00%','Age110':0,'Percent110':'0.00%','Undefinde':0,'Percent':'0.00%'}],'footer':[{'Age30':1,'Age40':4,'Age50':1,'Age60':0,'Age70':0,'Age80':0,'Age90':0,'Age110':0,'Undefinde':0,'OrgName':'总数：','DeptName':6}]}}";
            JavaScriptSerializer js = new JavaScriptSerializer();
            Dictionary<string, object> JsonObj = js.Deserialize<Dictionary<string, object>>(jsonString);
            HSSFWorkbook workbook = new HSSFWorkbook();
            MemoryStream ms = new MemoryStream();
            HSSFSheet sheet = workbook.CreateSheet() as HSSFSheet;
            sheet.CreateFreezePane(FrozenColumn, FrozenRow);

            int rowIndex = 0;
            #region 获取表头
            IEnumerable columnList = JsonObj["Columns"] as IEnumerable;//List<List<EasyColumn>>;
            ICellStyle headStyle = null;
            HSSFFont font = null;
            IEnumerable colList = null;
            int totalCol = 0;//总列数
            foreach (var item in columnList)
            {
                HSSFRow headerRow = sheet.CreateRow(rowIndex) as HSSFRow;
                headStyle = workbook.CreateCellStyle();
                headStyle.Alignment = HorizontalAlignment.CENTER;
                font = workbook.CreateFont() as HSSFFont;
                font.FontHeightInPoints = 14;
                headStyle.SetFont(font);
                int colIndex = totalCol;
                colList = item as IEnumerable;
                foreach (var column in colList)
                {
                    Dictionary<string, object> eaColumn = column as Dictionary<string, object>;
                    headerRow.CreateCell(colIndex).SetCellValue(eaColumn["title"].ToString());
                    #region 合并列
                    if (eaColumn["colspan"].ToString() != "1")
                    {
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, colIndex, colIndex + Convert.ToInt32(eaColumn["colspan"]) - 1);
                        sheet.AddMergedRegion(cellRangeAddress);
                    }
                    #endregion

                    #region 合并行
                    if (eaColumn["rowspan"].ToString() != "1")
                    {
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex + Convert.ToInt32(eaColumn["rowspan"]) - 1, colIndex, colIndex);
                        sheet.AddMergedRegion(cellRangeAddress);
                        headStyle.VerticalAlignment = VerticalAlignment.CENTER;
                        totalCol++;
                    }
                    #endregion
                    headerRow.GetCell(colIndex).CellStyle = headStyle;
                    colIndex += Convert.ToInt32(eaColumn["colspan"]);
                }
                rowIndex++;
            }
            #endregion


            #region 数据行
            Dictionary<string, object> jsonList = JsonObj["jsonData"] as Dictionary<string, object>;//List<List<EasyColumn>>;
            IEnumerable dataList = jsonList["rows"] as IEnumerable;
            foreach (var item in dataList)
            {
                Dictionary<string, object> row = item as Dictionary<string, object>;
                HSSFRow dataRow = sheet.CreateRow(rowIndex) as HSSFRow;
                int i = 0;
                foreach (var key in row.Keys)
                {
                    dataRow.CreateCell(i).SetCellValue(row[key].ToString());
                    i++;
                }

                rowIndex++;
            }
            #endregion

            workbook.Write(ms);
            ms.Flush();
            ms.Position = 0;

            sheet = null;
            //headerRow = null;
            workbook = null;
            return ms;
        }
    }


    public class EasyColumn
    {
        public string field { get; set; }
        public string title { get; set; }
        private int _width = 80;
        public int width
        {
            get { return _width; }
            set { _width = value; }
        }
        private string _align = "center";
        public string align
        {
            get { return _align; }
            set { _align = value; }
        }
        private int _colspan = 1;

        public int colspan
        {
            get { return _colspan; }
            set { _colspan = value; }
        }
        private int _rowspan = 1;

        public int rowspan
        {
            get { return _rowspan; }
            set { _rowspan = value; }
        }
        private bool _hidden = false;

        public bool hidden
        {
            get { return _hidden; }
            set { _hidden = value; }
        }
        private string _editor = "text";
        public string editor
        {
            get { return _editor; }
            set { _editor = value; }
        }
    }
}
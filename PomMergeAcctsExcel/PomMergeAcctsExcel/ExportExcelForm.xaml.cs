using System;
using System.Windows;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.XSSF.UserModel;
using System.Diagnostics;
using Microsoft.Win32;
using System.Data;
using NPOI.HSSF.Util;
using NPOI.SS.Util;
using System.Collections.Generic;
using System.Windows.Forms;
using NPOI.SS.UserModel;
namespace PomMergeAcctsExcel
{
    /// <summary>
    /// ExportExcelForm.xaml 的交互逻辑
    /// </summary>
    public partial class ExportExcelForm
    {
        public ExportExcelForm()
        {



        }






        public void text2()
        {

            try
            {
                Debug.Print(DateTime.Now.ToString("yyyyMMdd HH:mm:ss fff"));
                //创建Excel文件名称
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel|*.xlsx|Excel-2003|*.xls";
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                //saveFileDialog.FileName = "test" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                saveFileDialog.Title = "选择导出路径";


                exportExcel(saveFileDialog);

            }
            catch (Exception)
            {

                throw;
            }

        }

        private void exportExcel(object SaveFileDialog)
        {
            try
            {
                SaveFileDialog SaveFileDialoga = SaveFileDialog as SaveFileDialog;
                string pathname = "";
                string FileName = "sheet1";
                if (SaveFileDialoga.ShowDialog() == DialogResult.OK)
                {
                    pathname = SaveFileDialoga.FileName;
                }
                else
                {
                    return;
                }




                //txtCount.Text = "0/0";
                //txtTotal.Text = "0";

                // DataTable dt = GetRandTable();

                //DataTable dt = AnalyseTable(SqlHelper.ExecuteDataset(SqlHelper.connectionString, CommandType.Text, QuerySql).Tables[0], lGridTarget);
                DataTable dt = new DataTable();
                for (int i = 0; i < 10; i++)
                {
                    dt.Columns.Add(i.ToString());
                }
                for (int i = 0; i < 10; i++)
                {
                    DataRow dr = dt.NewRow();
                    for (int j = 0; j < 10; j++)
                    {
                        dr[j] = j.ToString();
                    }
                    dt.Rows.Add(dr);
                }
                FileStream fs = File.Create(pathname);

                //创建工作薄
                IWorkbook workbook = new XSSFWorkbook();
                //txtState.Text = "开始创建excel数据";
                int sheetIndex = 0, CurrentrowIndex = 0;
                //创建sheet
                ISheet sheet = workbook.CreateSheet("1234" + sheetIndex);

                ICellStyle Style = workbook.CreateCellStyle();
                Style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;
                Style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.CENTER;
                Style.BorderTop = NPOI.SS.UserModel.BorderStyle.THIN;
                Style.BorderRight = NPOI.SS.UserModel.BorderStyle.THIN;
                Style.BorderLeft = NPOI.SS.UserModel.BorderStyle.THIN;
                Style.BorderBottom = NPOI.SS.UserModel.BorderStyle.THIN;
                Style.DataFormat = 0;
                //sheet.SetColumnWidth(0, 100 * 256);
                //依次创建行和列
                //for (int i = 0; i < 40000; i++)
                //{
                //    CurrentrowIndex++; ;
                //    Debug.Print(i + "" + DateTime.Now.ToString("yyyyMMdd HH:mm:ss fff"));
                //    if (i != 0 && i % 20000 == 0)
                //    {
                //        sheetIndex++;
                //        CurrentrowIndex = 0;
                //        sheet = workbook.CreateSheet("sheet" + sheetIndex);

                //    }
                //    IRow row = sheet.CreateRow(CurrentrowIndex);
                //    row.HeightInPoints = 30;
                //    for (int j = 0; j < 50; j++)
                //    {

                //        ICell cell = row.CreateCell(j);
                //        cell.SetCellValue(i * 10 + j);

                //        cell.CellStyle = Style;
                //    }

                //}

                CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, dt.Columns.Count - 1);
                sheet.AddMergedRegion(cellRangeAddress);

                //((HSSFSheet)sheet).SetEnclosedBorderOfRegion(cellRangeAddress, BorderStyle.THIN, NPOI.HSSF.Util.HSSFColor.BLACK.index);

                IRow Titlerow = sheet.CreateRow(CurrentrowIndex);
                ICell Titlecell = Titlerow.CreateCell(0);
                Titlecell.SetCellValue("1234");

                Titlerow.HeightInPoints = 35;
                HSSFColor co = new HSSFColor.BLACK();

                IFont fon = GetFontStyle(workbook, "宋体", co, 16);
                ICellStyle TitleStyle = workbook.CreateCellStyle();
                TitleStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;
                TitleStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.CENTER;
                //TitleStyle.BorderTop = BorderStyle.THIN;
                //TitleStyle.BorderRight = BorderStyle.THIN;
                //TitleStyle.BorderLeft = BorderStyle.THIN;
                //TitleStyle.BorderBottom = BorderStyle.THIN;
                TitleStyle.DataFormat = 0;
                TitleStyle.SetFont(fon);
                Titlecell.CellStyle = TitleStyle;

                //IFont ColumnHeaderfon = GetFontStyle(workbook, "宋体", co, 10);
                //ICellStyle ColumnHeaderStyle = workbook.CreateCellStyle();
                //ColumnHeaderStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;
                //ColumnHeaderStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.CENTER;
                //ColumnHeaderStyle.BorderTop = BorderStyle.THIN;
                //ColumnHeaderStyle.BorderRight = BorderStyle.THIN;
                //ColumnHeaderStyle.BorderLeft = BorderStyle.THIN;
                //ColumnHeaderStyle.BorderBottom = BorderStyle.THIN;
                //ColumnHeaderStyle.DataFormat = 0;
                //ColumnHeaderStyle.SetFont(ColumnHeaderfon);

                IFont ColumnHeaderfon = GetFontStyle(workbook, "宋体", co, 12);
                ICellStyle ColumnHeaderStyle = workbook.CreateCellStyle();
                ColumnHeaderStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;
                ColumnHeaderStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.CENTER;
                ColumnHeaderStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.THIN;
                ColumnHeaderStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.THIN;
                ColumnHeaderStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.THIN;
                ColumnHeaderStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.THIN;
                ColumnHeaderStyle.DataFormat = 0;
                ColumnHeaderStyle.SetFont(ColumnHeaderfon);
                CurrentrowIndex++;
                IRow Headerrow = sheet.CreateRow(CurrentrowIndex);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    ICell cell = Headerrow.CreateCell(i);
                    cell.SetCellValue(dt.Columns[i].ColumnName);
                    cell.CellStyle = ColumnHeaderStyle;
                }

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    CurrentrowIndex++;
                    //txtCount.Text = (i + 1) + "/" + dt.Rows.Count;

                    if (i != 0 && i % 20000 == 0)
                    {
                        sheetIndex++;
                        CurrentrowIndex = 0;
                        sheet = workbook.CreateSheet(FileName + sheetIndex);

                    }
                    IRow row = sheet.CreateRow(CurrentrowIndex);
                    //row.HeightInPoints = 30;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {

                        ICell cell = row.CreateCell(j);
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                        cell.CellStyle = Style;
                    }

                }
                //向excel文件中写入数据并保保存
                //txtState.Dispatcher.BeginInvoke(updateAction, txtState, "正在保存文件");
                //txtState.Text = "正在保存文件";
                workbook.Write(fs);
                fs.Close();
                //txtState.Dispatcher.BeginInvoke(updateAction, txtState, "导出完毕");
                //txtState.Text = "导出完毕";

                Debug.Print(DateTime.Now.ToString("yyyyMMdd HH:mm:ss fff"));
                MessageBox.Show("导出完毕");

            }
            catch (Exception ex)
            {
                return;
            }

        }
        Random ran = new Random();
        private DataTable GetRandTable()
        {
            DataTable dt = new DataTable();
            //int rowCount = ran.Next(1, 100);
            //int ColumnCount = ran.Next(1, 100);
            int rowCount = 10;
            int ColumnCount = 10;
            for (int i = 0; i < ColumnCount; i++)
            {
                dt.Columns.Add();
            }
            for (int i = 0; i < rowCount; i++)
            {
                dt.Rows.Add();
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    dt.Rows[i][j] = i * 10 + j;
                }
            }
            return dt;
        }
        /// <summary>
        /// 获取字体样式
        /// </summary>
        /// <param name="hssfworkbook">Excel操作类</param>
        /// <param name="fontname">字体名</param>
        /// <param name="fontcolor">字体颜色</param>
        /// <param name="fontsize">字体大小</param>
        /// <returns></returns>
        public static IFont GetFontStyle(IWorkbook hssfworkbook, string fontfamily, HSSFColor fontcolor, int fontsize)
        {
            IFont font1 = hssfworkbook.CreateFont();
            if (string.IsNullOrEmpty(fontfamily))
            {
                font1.FontName = fontfamily;
            }
            if (fontcolor != null)
            {
                font1.Color = fontcolor.GetIndex();
            }
            //font1.IsItalic = true;
            font1.FontHeightInPoints = (short)fontsize;
            return font1;
        }




    }
}

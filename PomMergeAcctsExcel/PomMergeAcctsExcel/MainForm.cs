using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Reflection;
namespace PomMergeAcctsExcel
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog opf = new OpenFileDialog();
            if (opf.ShowDialog() == DialogResult.OK)
            {
                txtSrcFileName.Text = opf.FileName;
            }

            return;

            //using (ExcelHelper excelHelper = new ExcelHelper(SrcFileName))
            //{
            //    DataTable dt = excelHelper.ExcelToDataTable("sheet1", true);
            //}
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Excel 2007版本|*.xlsx|Excel 2003版本|*.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                txtDestFileName.Text = sfd.FileName;
            }

            //return;
            //ExportExcelForm form11 = new ExportExcelForm();
            //form11.text2();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (txtSrcFileName.Text == null || txtSrcFileName.Text.Equals(""))
            {
                MessageBox.Show("请选择您想要合并的Excel文件");
                btnSrcSelect.Focus();
                return;
            }

            if (txtDestFileName.Text == null || txtDestFileName.Text.Equals(""))
            {
                MessageBox.Show("请选择合并后Excel文件保存的路径");
                btnDestSelect.Focus();
                return;
            }

            btnExecute.Enabled = false;
            ExcelHelper helper = new ExcelHelper(txtSrcFileName.Text, txtDestFileName.Text, this);
            

            int iRslt = helper.ReadExcelMergeIntoDic();
            if (iRslt != 1)
            {
                MessageBox.Show("读取文件失败，请检查您的文件内容是否正确");
                return;
            }

            iRslt = helper.WriteToExcelFromDic();
            if (iRslt != 1)
            {
                MessageBox.Show("文件写入失败，请确定是否有权限将文件写入目标路径");
                return;
            }

            MessageBox.Show("文件合并成功！");
            btnExecute.Enabled = true;

            pb1.Value = 0;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Contact contact = new Contact();
            foreach (System.Reflection.PropertyInfo prop in contact.GetType().GetProperties())
            {
                Console.WriteLine(prop.Name);

                prop.SetValue
            }
        }
    }
}

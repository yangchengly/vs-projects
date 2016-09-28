using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Data;
namespace PomMergeAcctsExcel
{
    public class ExcelHelper : IDisposable
    {
        private MainForm form1;
        private static String SEPERATOR = ";";
        private String srcFileName = null;
        private String dstFileName = null;


        private FileStream fsRead = null;
        private FileStream fsWrite = null;

        private bool disposed = false;
        private Dictionary<String, Contact> dic = new Dictionary<string, Contact>();

        public ExcelHelper(String srcFileName, String dstFileName, MainForm form1)
        {
            this.srcFileName = srcFileName;
            this.dstFileName = dstFileName;

            this.form1 = form1;
        }

        public int ReadExcelMergeIntoDic()
        {
            fsRead = new FileStream(srcFileName, FileMode.OpenOrCreate, FileAccess.Read);

            IWorkbook workbook = null;
            if (srcFileName.IndexOf(".xlsx") > 0) // 2007版本
                workbook = new XSSFWorkbook(fsRead);
            else if (srcFileName.IndexOf(".xls") > 0) // 2003版本
                workbook = new HSSFWorkbook(fsRead);

            ISheet sheet = null;
            try
            {
                sheet = workbook.GetSheet("Sheet1");


                int iLastRowNum = sheet.LastRowNum;
                for (int i = 0; i < iLastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null)
                        break;

                    Contact contact = new Contact()
                    {
                        id = row.GetCell(0) != null ? row.GetCell(0).ToString() : "",
                        custNo = row.GetCell(1) != null ? row.GetCell(1).ToString() : "",
                        cardHolder = row.GetCell(2) != null ? row.GetCell(2).ToString() : "",
                        sex = row.GetCell(3) != null ? row.GetCell(3).ToString() : "",
                        mvNo = row.GetCell(4) != null ? row.GetCell(4).ToString() : "",
                        daAcCnt = row.GetCell(5) != null ? row.GetCell(5).ToString() : "",
                        acNo = row.GetCell(6) != null ? row.GetCell(6).ToString() : "",
                        acTyp = row.GetCell(7) != null ? row.GetCell(7).ToString() : "",
                        cycle = row.GetCell(8) != null ? row.GetCell(8).ToString() : "",
                        lastStmtDte = row.GetCell(9) != null ? row.GetCell(9).ToString() : "",
                        duedate = row.GetCell(10) != null ? row.GetCell(10).ToString() : "",
                        stmtBalR = row.GetCell(11) != null ? row.GetCell(11).ToString() : "",
                        stmtBalU = row.GetCell(12) != null ? row.GetCell(12).ToString() : "",
                        totBalR = row.GetCell(13) != null ? row.GetCell(13).ToString() : "",
                        totBalU = row.GetCell(14) != null ? row.GetCell(14).ToString() : "",
                        pastDueR = row.GetCell(15) != null ? row.GetCell(15).ToString() : "",
                        pastDueU = row.GetCell(16) != null ? row.GetCell(16).ToString() : "",
                        totR = row.GetCell(17) != null ? row.GetCell(17).ToString() : "",
                        delqDays = row.GetCell(18) != null ? row.GetCell(18).ToString() : "",
                        afeeFlag = row.GetCell(19) != null ? row.GetCell(19).ToString() : "",
                        pfeeFlag = row.GetCell(20) != null ? row.GetCell(20).ToString() : "",
                        collId = row.GetCell(21) != null ? row.GetCell(21).ToString() : "",
                        coll = row.GetCell(22) != null ? row.GetCell(22).ToString() : "",
                        arDate = row.GetCell(23) != null ? row.GetCell(23).ToString() : "",
                        arTime = row.GetCell(24) != null ? row.GetCell(24).ToString() : "",
                        arPayment = row.GetCell(25) != null ? row.GetCell(25).ToString() : ""
                    };

                    if (!dic.ContainsKey(contact.custNo))
                    {
                        dic.Add(contact.custNo, contact);
                    }
                    else
                    {
                        Contact contactInDic = dic[contact.custNo];                        

                        contactInDic.acNo = contactInDic.acNo + SEPERATOR + contact.acNo;
                        contactInDic.acTyp = contactInDic.acTyp + SEPERATOR + contact.acTyp;
                        contactInDic.lastStmtDte = contactInDic.lastStmtDte + SEPERATOR + contact.lastStmtDte;
                        contactInDic.duedate = contactInDic.duedate + SEPERATOR + contact.duedate;
                        contactInDic.stmtBalR = contactInDic.stmtBalR + SEPERATOR + contact.stmtBalR;
                        contactInDic.stmtBalU = contactInDic.stmtBalU + SEPERATOR + contact.stmtBalU;
                        contactInDic.totBalR = contactInDic.totBalR + SEPERATOR + contact.totBalR;
                        contactInDic.totBalU = contactInDic.totBalU + SEPERATOR + contact.totBalU;
                        contactInDic.pastDueR = contactInDic.pastDueR + SEPERATOR + contact.pastDueR;
                        contactInDic.pastDueU = contactInDic.pastDueU + SEPERATOR + contact.pastDueU;
                        contactInDic.totR = contactInDic.totR + SEPERATOR + contact.totR;
                        contactInDic.delqDays = contactInDic.delqDays + SEPERATOR + contact.delqDays;
                        contactInDic.afeeFlag = contactInDic.afeeFlag + SEPERATOR + contact.afeeFlag;
                        contactInDic.pfeeFlag = contactInDic.pfeeFlag + SEPERATOR + contact.pfeeFlag;
                        contactInDic.stmtBalU = contactInDic.stmtBalU + SEPERATOR + contact.stmtBalU;
                        contactInDic.stmtBalU = contactInDic.stmtBalU + SEPERATOR + contact.stmtBalU;

                        contactInDic.stmtBalU = contactInDic.stmtBalU + SEPERATOR + contact.stmtBalU;
                        contactInDic.stmtBalU = contactInDic.stmtBalU + SEPERATOR + contact.stmtBalU;
                        contactInDic.stmtBalU = contactInDic.stmtBalU + SEPERATOR + contact.stmtBalU;
                        contactInDic.stmtBalU = contactInDic.stmtBalU + SEPERATOR + contact.stmtBalU;
                        contactInDic.stmtBalU = contactInDic.stmtBalU + SEPERATOR + contact.stmtBalU;
                    }

                    this.form1.pb1.Value = i * 50 / iLastRowNum;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                fsRead.Close();
            }

            return 1;
        }

        public int WriteToExcelFromDic()
        {
            fsWrite = new FileStream(dstFileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);

            IWorkbook workbook = null;
            if (dstFileName.IndexOf(".xlsx") > 0) // 2007版本
                workbook = new XSSFWorkbook();
            else if (dstFileName.IndexOf(".xls") > 0) // 2003版本
                workbook = new HSSFWorkbook();

            ISheet sheet = workbook.CreateSheet("sheet1");
            int length = dic.Count;

            int i = 0;
            foreach (var value in dic.Values)
            {
                IRow row = sheet.CreateRow(i++);

                row.CreateCell(0).SetCellValue(value.id);
                row.CreateCell(1).SetCellValue(value.custNo);
                row.CreateCell(2).SetCellValue(value.cardHolder);
                row.CreateCell(3).SetCellValue(value.sex);
                row.CreateCell(4).SetCellValue(value.mvNo);
                row.CreateCell(5).SetCellValue(value.daAcCnt);
                row.CreateCell(6).SetCellValue(value.acNo);
                row.CreateCell(7).SetCellValue(value.acTyp);
                row.CreateCell(8).SetCellValue(value.cycle);
                row.CreateCell(9).SetCellValue(value.lastStmtDte);
                row.CreateCell(10).SetCellValue(value.duedate);
                row.CreateCell(11).SetCellValue(value.stmtBalR);
                row.CreateCell(12).SetCellValue(value.stmtBalU);
                row.CreateCell(13).SetCellValue(value.totBalR);
                row.CreateCell(14).SetCellValue(value.totBalU);
                row.CreateCell(15).SetCellValue(value.pastDueR);
                row.CreateCell(16).SetCellValue(value.pastDueU);
                row.CreateCell(17).SetCellValue(value.totR);
                row.CreateCell(18).SetCellValue(value.delqDays);
                row.CreateCell(19).SetCellValue(value.afeeFlag);
                row.CreateCell(20).SetCellValue(value.pfeeFlag);
                row.CreateCell(21).SetCellValue(value.collId);
                row.CreateCell(22).SetCellValue(value.coll);
                row.CreateCell(23).SetCellValue(value.arDate);
                row.CreateCell(24).SetCellValue(value.arTime);
                row.CreateCell(25).SetCellValue(value.arPayment);

                this.form1.pb1.Value = 50 + i * 50 / dic.Count;
            }

            workbook.Write(fsWrite);
            fsWrite.Close();

            return 1;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    if (fsRead != null)
                    {
                        fsRead.Close();
                    }

                    if (fsWrite != null)
                    {
                        fsWrite.Close();
                    }
                }

                fsRead = null;
                fsWrite = null;
                disposed = true;
            }
        }


    }
}
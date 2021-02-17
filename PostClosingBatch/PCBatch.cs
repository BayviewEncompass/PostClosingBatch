using EllieMae.Encompass.Automation;
using EllieMae.Encompass.BusinessObjects.Loans;
using EllieMae.Encompass.Client;
using EllieMae.Encompass.ComponentModel;
using EllieMae.Encompass.Query;
using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

namespace PostClosingBatch
{
    [Plugin]
    public class PCBatch
    {
        public PCBatch()
        {
            EncompassApplication.LoanOpened += EncompassApplication_LoanOpened;
        }
        private void EncompassApplication_LoanOpened(object sender, EventArgs e)
        {
            EncompassApplication.CurrentLoan.FieldChange += CurrentLoan_FieldChange;
        }


        private void CurrentLoan_FieldChange(object source, EllieMae.Encompass.BusinessObjects.Loans.FieldChangeEventArgs e)
        {
            if (e.FieldID == "CX.CHRIS.BATCH.UPDATE")
            {
                BatchUpdater();
            }

        }

        private void BatchUpdater()
        {


            Session newSession = new Session();
            newSession.Start("https://TEBE11147866.ea.elliemae.net$TEBE11147866", "admin", "Y0uT@lk!ngT0M3?");
            String currentUser = EncompassApplication.Session.UserID;

            //Excel prep and call
            Microsoft.Office.Interop.Excel.Application userApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook userWorkbook = userApp.Workbooks.Open(@"H:\Encompass Support\Batch_Updater3.csv");
            Microsoft.Office.Interop.Excel._Worksheet userWorksheet = userWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range userRange = userWorksheet.UsedRange;
            //userWorksheet.Cells.NumberFormat = "General";

            //set shortcut for loan call
            Loan loan = EncompassApplication.CurrentLoan;

            //row/column setup
            int rCnt = 1;
            int cCnt = 1;
            int rowCount = userRange.Rows.Count;
            int colCount = userRange.Columns.Count;

            string r1c2string = "";
            string r1c3string = "";
            string r1c4string = "";
            string r1c5string = "";
            string r1c6string = "";

            double colData1 = 0;
            string colDatSt = colData1.ToString();

            cCnt++;
            var r1c2 = userWorksheet.Cells[rCnt, cCnt].Value2;
            r1c2string = r1c2.ToString();
            cCnt++;
            var r1c3 = userWorksheet.Cells[rCnt, cCnt].Value2;
            if (r1c3 != null)
            {
                r1c3string = r1c3.ToString();
                cCnt++;
            }
            var r1c4 = userWorksheet.Cells[rCnt, cCnt].Value2;
            if (r1c4 != null)
            {
                r1c4string = r1c4.ToString();
                cCnt++;
            }
            var r1c5 = userWorksheet.Cells[rCnt, cCnt].Value2;
            if (r1c5 != null)
            {
                r1c5string = r1c5.ToString();
                cCnt++;
            }
            var r1c6 = userWorksheet.Cells[rCnt, cCnt].Value2;
            if (r1c6 != null)
            {
                r1c6string = r1c6.ToString();
                cCnt++;
            }


            cCnt = 1;

            try
            {
                rCnt = 2;
                cCnt = 2;
                var colData2 = userWorksheet.Cells[rCnt, cCnt].Text;
                string colData2string = "";
                cCnt++;
                var colData3 = userWorksheet.Cells[rCnt, cCnt].Text;
                string colData3string = "";
                cCnt++;
                var colData4 = userWorksheet.Cells[rCnt, cCnt].Text;
                string colData4string = "";
                cCnt++;
                var colData5 = userWorksheet.Cells[rCnt, cCnt].Text;
                string colData5string = "";
                cCnt++;
                var colData6 = userWorksheet.Cells[rCnt, cCnt].Text;
                string colData6string = "";
                cCnt = 1;


                for (rCnt = 2; rCnt <= rowCount; rCnt++)
                {
                    if (userWorksheet.Cells[rCnt, cCnt].Value2 is null)
                    {
                        break;
                    }
                    colData1 = (double)userWorksheet.Cells[rCnt, cCnt].Value2;
                    colDatSt = colData1.ToString();
                    if (colData1.ToString().Substring(0, 1) == "6")
                    {

                        colDatSt = "00" + colDatSt;
                        colData1 = Convert.ToDouble(colDatSt);
                    }
                    cCnt++;

                    colData2 = userWorksheet.Cells[rCnt, cCnt].Text;
                    colData2string = colData2.ToString();

                    //Additional columns as needed
                    cCnt++;
                    colData3 = userWorksheet.Cells[rCnt, cCnt].Text;
                    if (colData3 != "")
                    {
                        colData3string = colData3.ToString();
                        cCnt++;
                    }
                    colData4 = userWorksheet.Cells[rCnt, cCnt].Text;
                    if (colData4 != "")
                    {
                        colData4string = colData4.ToString();
                        cCnt++;
                    }
                    colData5 = userWorksheet.Cells[rCnt, cCnt].Text;
                    if (colData5 != "")
                    {
                        colData5string = colData5.ToString();
                        cCnt++;
                    }
                    colData6 = userWorksheet.Cells[rCnt, cCnt].Text;
                    if (colData6 != "")
                    {
                        colData6string = colData6.ToString();
                        cCnt++;
                    }

                    cCnt = 1;

                    StringFieldCriterion cri = new StringFieldCriterion();
                    cri.FieldName = "Loan.LoanNumber";
                    cri.Value = colDatSt;
                    BatchUpdate batch = new BatchUpdate(cri);
                    batch.Fields.Add(r1c2string, colData2string);
                    if (colData3string != "")
                    {
                        batch.Fields.Add(r1c3string, colData3string);
                    }
                    if (colData4string != "")
                    {
                        batch.Fields.Add(r1c4string, colData4string);
                    }
                    if (colData5string != "")
                    {
                        batch.Fields.Add(r1c5string, colData5string);
                    }
                    if (colData6string != "")
                    {
                        batch.Fields.Add(r1c6string, colData6string);
                    }
                    EncompassApplication.Session.Loans.SubmitBatchUpdate(batch);
                }




            
            
            }
            catch (Exception)
            {

                throw;
            }

            userRange.Delete(XlDeleteShiftDirection.xlShiftUp);
            string fileName = @"H:\Encompass Support\Batch_Updater3.csv";
            string folder = System.IO.Path.GetDirectoryName(fileName);
            if (System.IO.Directory.Exists(folder))
            {
                
                userWorkbook.Close(true, fileName, null);
            }

            


            userApp.Quit();
            Marshal.ReleaseComObject(userWorksheet);
            Marshal.ReleaseComObject(userWorkbook);
            Marshal.ReleaseComObject(userApp);
            newSession.End();



        }
    }
}
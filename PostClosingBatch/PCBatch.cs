using EllieMae.EMLite.RemotingServices;
using EllieMae.Encompass.Automation;
using EllieMae.Encompass.BusinessObjects.Loans;
using EllieMae.Encompass.Client;
using EllieMae.Encompass.ComponentModel;
using EllieMae.Encompass.Query;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PostClosingBatch;
using EllieMae.Encompass.Collections;
using Excel = Microsoft.Office.Interop.Excel;
using Session = EllieMae.Encompass.Client.Session;
using EllieMae.EMLite.ClientServer;
using Newtonsoft.Json;

namespace PostClosingBatch
{
    [Plugin]
    public class PCBatch
    {
        private static string acctExec = "BatchUpdatePC.json";
        private static BatchUpdateDB file;
        public static string RepName = null;
        public static string Col1 = null;
        public static string Col2 = null;
        public static string Col3 = null;
        public static string Col4 = null;
        public static string Col5 = null;
        //public static string BatchUpdatePC = null;
        public static BatchUpdateDB CDO => file ?? DownloadCDO();
        public PCBatch()
        {
            EllieMae.Encompass.Automation.EncompassApplication.LoanOpened += EncompassApplication_LoanOpened;
        }
        private void EncompassApplication_LoanOpened(object sender, EventArgs e)
        {
            EllieMae.Encompass.Automation.EncompassApplication.CurrentLoan.FieldChange += CurrentLoan_FieldChange;

        }

        private static BatchUpdateDB DownloadCDO()
        {
            file = JsonConvert.DeserializeObject<BatchUpdateDB>(Encoding.UTF8.GetString(EncompassApplication.Session.DataExchange.GetCustomDataObject(acctExec).Data));
            return file;

        }


        private void CurrentLoan_FieldChange(object source, EllieMae.Encompass.BusinessObjects.Loans.FieldChangeEventArgs e)
        {
            if (e.FieldID == "CX.CHRIS.BATCH.UPDATE")
            {
                DownloadCDO();
                
                BatchUpdater();
            }

        }

        private void BatchUpdater()
        {
            Loan loan = EncompassApplication.CurrentLoan;
            //EncompassApplication.Session.Start()
            //EncompassApplication.CurrentLoan.Session.StartInstance()
            //var sessionStart = new EncompassApplication.CurrentLoan.Session;
            Session sessionStart = new Session();
            sessionStart.Start("https://TEBE11147866.ea.elliemae.net$TEBE11147866", "admin", "Y0uT@lk!ngT0M3?");
            String currentUser = EncompassApplication.Session.UserID;
            string repAddress = loan.Fields["CX.PC.BATCH.NAME"].Value.ToString() + ".csv";

                //Excel prep and call
            Microsoft.Office.Interop.Excel.Application userApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook userWorkbook = userApp.Workbooks.Open(@"\\ftwfs02\Groups\LLS-Scanned\Bulk Updating\Encompass Bulk Update Templates\" + repAddress);
            Microsoft.Office.Interop.Excel._Worksheet userWorksheet = userWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range userRange = userWorksheet.UsedRange;
            //userWorksheet.Cells.NumberFormat = "General";
            try
            {
                //set shortcut for loan call
                //Loan loan = EncompassApplication.CurrentLoan;

            //row/column setup
            int rCnt = 1;
            int cCnt = 1;
            int rowCount = userRange.Rows.Count;
            int colCount = userRange.Columns.Count;

            List<BatchUpdatePC> bUpdate = CDO.BatchUpdatePC.ToList();



            string r1c2string = "";
            string r1c3string = "";
            string r1c4string = "";
            string r1c5string = "";
            string r1c6string = "";
            //var batchT = loan.Fields["CX.PC.BATCH.TYPE"].Value;
            //string batchType = batchT.ToString();
            DateTime dateRan = new DateTime();
            dateRan = DateTime.Now;
            
            
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
                else
                {
                    r1c3string = "null";
                    cCnt++;
                }
            var r1c4 = userWorksheet.Cells[rCnt, cCnt].Value2;
            if (r1c4 != null)
            {
                r1c4string = r1c4.ToString();
                cCnt++;
            }
                else
                {
                    r1c4string = "null";
                    cCnt++;
                }
                var r1c5 = userWorksheet.Cells[rCnt, cCnt].Value2;
            if (r1c5 != null)
            {
                r1c5string = r1c5.ToString();
                cCnt++;
            }
                else
                {
                    r1c5string = "";
                    cCnt++;
                }
                var r1c6 = userWorksheet.Cells[rCnt, cCnt].Value2;
            if (r1c6 != null)
            {
                r1c6string = r1c6.ToString();
                cCnt++;
            }
                else
                {
                    r1c6string = "";
                    cCnt++;
                }


                cCnt = 1;

                foreach (BatchUpdatePC item in bUpdate)
                {
                    if (item.RepName.ToLower() == loan.Fields["CX.PC.BATCH.NAME"].Value.ToString().ToLower())
                    {
                        if (r1c2string.ToLower() != item.Col2.ToLower() | r1c3string.ToLower() != item.Col3.ToLower() | r1c4string.ToLower() != item.Col4.ToLower() | r1c5string.ToLower() != item.Col5.ToLower())
                        {
                            MessageBox.Show("This does not appear to be the correct format for this report.  Please review each column in the Excel spreadsheet for accuracy.");
                            loan.Fields["CX.PC.BATCH.IRON.MOUNTAIN"].Value = "Fail";
                        }
                        else
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
                               batch.Fields.Add("CX.PC.BATCH.COLL.TRACK", currentUser);
                               batch.Fields.Add("CX.PC.BATCH.COL.DATE", dateRan);
                               loan.Fields["CX.PC.BATCH.IRON.MOUNTAIN"].Value = "Fail";
                                EllieMae.Encompass.Automation.EncompassApplication.Session.Loans.SubmitBatchUpdate(batch);
                                loan.Fields["CX.PC.BATCH.IRON.MOUNTAIN"].Value = "Success";
                            }
                        }
                    }

                }


            }
            catch (Exception)
            {

                MessageBox.Show("Something went wrong.  Please click YES to save on the next popup, rebuild the spreadsheet and run the Batch Update again."); 
            }
            finally
            {
                
                string fileName = @"\\ftwfs02\Groups\LLS-Scanned\Bulk Updating\Encompass Bulk Update Templates\" + repAddress;
                //string folder = System.IO.Path.GetDirectoryName(fileName);
                //if (System.IO.Directory.Exists(folder))
                //{
                //MessageBox.Show(loan.Fields["CX.PC.BATCH.IRON.MOUNTAIN"].Value.ToString());
                    userWorkbook.Close(true, fileName, null);
                //}
                userApp.Quit();
                Marshal.ReleaseComObject(userWorksheet);
                Marshal.ReleaseComObject(userWorkbook);
                Marshal.ReleaseComObject(userApp);
                sessionStart.End();
               

            }


            //userRange.Delete(XlDeleteShiftDirection.xlShiftUp);
            //string fileName = @"H:\Encompass Support\Batch_Updater2.csv";
            //string folder = System.IO.Path.GetDirectoryName(fileName);
            //if (System.IO.Directory.Exists(folder))
            //{

            //    userWorkbook.Close(true, fileName, null);
            //}

            //userApp.Quit();
            //Marshal.ReleaseComObject(userWorksheet);
            //Marshal.ReleaseComObject(userWorkbook);
            //Marshal.ReleaseComObject(userApp);
            //newSession.End();


        }
    }
}
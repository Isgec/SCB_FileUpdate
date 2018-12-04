using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Globalization;
using System.Timers;
using Microsoft.Office.Interop;
using System.Diagnostics;

namespace SCBFILEUpdate_
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private static void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            //this.close();

            Environment.Exit(0);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    string[] txtFiles = Directory.GetFiles(fbd.SelectedPath, "*.txt");
                    foreach (string currentFile in txtFiles)
                    {
                        string fileName = Path.GetFileName(currentFile);
                        using (StreamReader sr = new StreamReader(File.Open(currentFile, FileMode.Open, FileAccess.Read)))
                        {
                            using (SqlConnection con = new SqlConnection("Server = 192.9.200.45; Database = inforerpdb; User Id = dev1; Password = dev1; "))
                            {
                                char[] fieldSep = new char[] { ',' };
                                char[] delim = new char[] { '"' };
                                con.Open();
                                string line = "";
                                string cmdTxt = "";
                                int count = 0;
                                int TotalHRecords = 0;

                                string cmdCountTxt = "select count(*) from ttfisg052200 ";
                                using (SqlCommand cmd = new SqlCommand(cmdCountTxt, con))
                                {
                                    count = (int)(cmd.ExecuteScalar());
                                    TotalHRecords = count + 1;
                                }

                                while ((line = sr.ReadLine()) != "" && line != null)
                                {
                                    string[] parts = line.Split(fieldSep);
                                    int nFields = parts.Length;
                                    int RecordExist;
                                    int TotalDRecords = 0;
                                    if (parts[0].Trim(delim) == "H")
                                    {
                                        string recordexist = "select count(*) from ttfisg052200 where t_fnam in ('" + fileName.Trim(delim) + "')";
                                        using (SqlCommand cmd = new SqlCommand(recordexist, con))
                                        {
                                            RecordExist = (int)(cmd.ExecuteScalar());
                                        }
                                        if (RecordExist == 0)
                                        {
                                            DateTime DtofPreparation = DateTime.ParseExact((parts[3].Trim(delim) + parts[4].Trim(delim)), "yyMMddHHmm", CultureInfo.InvariantCulture);

                                            cmdTxt = @"INSERT INTO ttfisg052200 (t_hsrn,t_date,t_user,t_t_no,t_fnam,t_flid,t_rtyp,t_sdid,t_rpid,t_dprp,t_utrn,
                                     t_ftyp,t_brcd,t_cus1,t_cus2,t_cus3,t_cus4,t_cus5,t_adr1,t_adr2,t_adr3,t_adr4,t_adr5,t_pcod
                                     ,t_ccod,t_cfr1,t_cfr2,t_cfr3,t_cfr4,t_t_rt,t_Refcntd,t_Refcntu,t_uerp,t_udon,t_udby)";

                                            cmdTxt += "VALUES (" + TotalHRecords + ",'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' ,'" + "3194" + "' ," + TotalHRecords + ",'" + fileName + "'," + TotalHRecords + "";
                                            cmdTxt += String.Format(@",'{0}','{1}','{2}','{3}','{4}' ,'{5}','{6}','{7}','{8}','{9}','{10}',
                                                '{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}','{25}','{26}','{27}','{28}'
                                                )", parts[0].Trim(delim), parts[1].Trim(delim), parts[2].Trim(delim),
                                                   DtofPreparation, parts[5].Trim(delim),
                                                    parts[6].Trim(delim), parts[7].Trim(delim), parts[8].Trim(delim),
                                                    parts[9].Trim(delim), parts[10].Trim(delim), parts[11].Trim(delim),
                                                    parts[12].Trim(delim), parts[13].Trim(delim), parts[14].Trim(delim),
                                                    parts[15].Trim(delim), parts[16].Trim(delim), parts[17].Trim(delim),
                                                    parts[18].Trim(delim), parts[19].Trim(delim), parts[20].Trim(delim),
                                                    parts[21].Trim(delim), parts[22].Trim(delim), parts[23].Trim(delim), "H", 0, 0, 1, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), "3194");

                                            using (SqlCommand cmd = new SqlCommand(cmdTxt, con))
                                            {
                                                cmd.ExecuteNonQuery();
                                            }
                                        }

                                    }
                                    else if (parts[0].Trim(delim) == "D")
                                    {
                                        //string recordexist = "select count(*) from ttfisg053200 where t_cusr  in ('" + parts[1].Trim(delim) + "')";
                                        //using (SqlCommand cmd = new SqlCommand(recordexist, con))
                                        //{
                                        //    RecordExist = (int)(cmd.ExecuteScalar());
                                        //}
                                        //if (RecordExist == 0)
                                        //{
                                        string cmdTotalDCountTxt = "select count(*) from ttfisg053200 ";
                                        using (SqlCommand cmd = new SqlCommand(cmdTotalDCountTxt, con))
                                        {
                                            count = (int)(cmd.ExecuteScalar());
                                            TotalDRecords = count + 1;
                                        }

                                        cmdTxt = @"INSERT INTO ttfisg053200 (t_hsrn,t_dsrn,t_rtyp,t_cusr,t_stat,t_stds,t_batr
                                                 ,t_payr,t_pidt,t_pydt,t_dedt,t_pamt,t_pcur,t_scur,t_tcur,t_exrt,t_ecur,t_stdc
                                                 ,t_chqn,t_cqdt,t_acno,t_acnm,t_benm,t_bad1,t_bad2,t_bad3,t_bad4,t_flr1,t_flr2
                                                 ,t_fl2a,t_flr3,t_flr4,t_flr5,t_flr6,t_flr7,t_flr8,t_flr9,t_flid,t_Refcntd,t_Refcntu,t_uerp,t_udat,t_udby) ";

                                        cmdTxt += "VALUES ( " + TotalHRecords + "," + TotalDRecords + ",";
                                        for (int i = 0; i < nFields; ++i)
                                        {
                                            string field = parts[i].Trim(delim);
                                            if (i == 0)
                                                cmdTxt += String.Format("'{0}'", field);
                                            else
                                                cmdTxt += String.Format(", '{0}'", field);
                                        }
                                        cmdTxt += "," + TotalDRecords + ", 0,0,1, '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' ,'" + "3194" + "' )";
                                        using (SqlCommand cmd = new SqlCommand(cmdTxt, con))
                                        {
                                            cmd.ExecuteNonQuery();
                                        }
                                        lblInsert.Text = "" + fileName.Trim(delim) + " records have been inserted!";
                                        //}
                                        //else
                                        //{
                                        //    lblInsert.Text = "All the records have already been inserted!";
                                        //}
                                    }
                                    else
                                    {
                                        cmdTxt = string.Empty;
                                    }

                                }
                                con.Close();
                            }
                        }
                    }
                }
            }
            System.Timers.Timer timer = new System.Timers.Timer(5000);
            timer.Elapsed += Timer_Elapsed;
            timer.Start();
        }
        public void btnExport_Click(object sender, EventArgs e)
        {
            DataTable dtnew = new DataTable();
            using (SqlConnection con = new SqlConnection("Server = 192.9.200.45; Database = inforerpdb; User Id = dev1; Password = dev1; "))
            {
                string sExcelFilePath = "";
               //// string scmd= @"Select t_paym as " + "Payment Type" + ",'' as " + "Payment Date" + ",t_bpid as " + "Payee ID" + ",t_amnt as " + "Payment Amount" + ",t_docn as " + "Customer Ref" + "," + "Payment Detail" + " =( case t_tadv
               //    string scmd= @"Select t_paym as "Payment Type",'' as "Payment Date",t_bpid as "Payee ID",t_amnt as "Payment Amount",t_docn as "Customer Ref","Payment Detail" = ( case t_tadv
               //      when 1 then 'Purchase Invoice'
               //     when 2 then 'Purchase Credit Note'
               //    when 3 then 'Sales Invoice'
               //    when 4 then 'Sales Credit Note'
               //    when 5 then 'Advance Payment'
               //    when 6 then 'Unallocated Payment'
               //    when 7 then 'Standing Order'
               //    when 8 then 'Stand-alone Payment'
               //     ELSE Null
               //     END
               //     ) ,
                  
               //    '' as "Delivery Method" ,'' as "Delivery To",'' as "Print Location",'' as "Payble Location",'' as "POP Code",'' as "Supplier Name",'' as "Supplier eMail id",'' as "Cheque Number" from ttfisg051200 where t_fnam = '' and t_pmtf = '2'";
                using (SqlCommand cmd = new SqlCommand("SP_SCBExportToCSV",con))
                {

                    using (SqlDataAdapter sda = new SqlDataAdapter())
                    {
                        cmd.Connection = con;
                        cmd.CommandType = CommandType.StoredProcedure;
                        sda.SelectCommand = cmd;
                        using (DataTable dt = new DataTable())
                        {
                            sda.Fill(dt);
                            string sFilename = DateTime.Now.ToString("yyyyMMdd HHmmss") + ".csv";

                            sExcelFilePath = @"C:\host2host\documents\payments\out\h2h_payments\" + sFilename;
                            string sExcelFilePath2 = @"C:\host2host\documents\payments\out2\h2h_payments\" + sFilename;
                            // ExportToExcel(dt, sExcelFilePath, sExcelFilePath2, sFilename);
                            ExportCSV(dt, sExcelFilePath, sExcelFilePath2, sFilename);
                            System.Timers.Timer timer = new System.Timers.Timer(5000);
                            timer.Elapsed += Timer_Elapsed;
                            timer.Start();
                        }
                    }
                }
            }
        }

        // public void ExportToExcel(DataTable DataTable, string ExcelFilePath, string CopyExcelPath,string filename)
        // public void ExportToExcel(DataTable DataTable, string ExcelFilePath, string filename)
        //{
        //    using (SqlConnection con = new SqlConnection("Server = 192.9.200.45; Database = inforerpdb; User Id = dev1; Password = dev1; "))
        //    {
        //        int Filecount = 0;
        //        SqlCommand cmd2 = new SqlCommand();
        //        cmd2.Connection = con;
        //        SqlCommand cmd3 = new SqlCommand();
        //        cmd3.Connection = con;

        //        con.Open();
        //        cmd2.CommandText = @"SELECT count(distinct t_fnam)FROM ttfisg051200 where t_fnam not in ('')";
        //        Filecount = (int)cmd2.ExecuteScalar();
        //        Filecount++;

        //        try
        //        {
        //            if (DataTable.Rows.Count > 0)
        //            {
        //                cmd3.CommandText = @"update ttfisg051200 set t_fnam = '" + filename + "' , t_flid = " + Filecount + ", t_pmtf= '1' where t_fnam=''";
        //                cmd3.ExecuteNonQuery();

        //                try
        //                {
        //                    int ColumnsCount;

        //                    if (DataTable == null || (ColumnsCount = DataTable.Columns.Count) == 0)
        //                        throw new Exception("ExportToExcel: Null or empty input table!\n");

        //                    // load excel, and create a new workbook
        //                    Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
        //                    Excel.Workbooks.Add();

        //                    // single worksheet
        //                    Microsoft.Office.Interop.Excel._Worksheet Worksheet = Excel.ActiveSheet;

        //                    object[] Header = new object[ColumnsCount];

        //                    // column headings               
        //                    //for (int i = 0; i < ColumnsCount; i++)
        //                    //    Header[i] = DataTable.Columns[i].ColumnName;
        //                    Header[0] = "Payment Type";
        //                    Header[1] = "Payment Date";
        //                    Header[2] = "Payee ID";
        //                    Header[3] = "Payment Amount";
        //                    Header[4] = "Customer Ref";
        //                    Header[5] = "Payment Detail";
        //                    Header[6] = "Delivery Method";
        //                    Header[7] = "Delivery To";
        //                    Header[8] = "Print Location";
        //                    Header[9] = "Payble Location";
        //                    Header[10] = "POP Code";
        //                    Header[11] = "Supplier Name";
        //                    Header[12] = "Supplier eMail id";
        //                    Header[13] = "Cheque Number";

        //                    Microsoft.Office.Interop.Excel.Range HeaderRange =
        //                                    Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[1, 1]),
        //                                    (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[1, ColumnsCount]));

        //                    HeaderRange.Value = Header;
        //                    HeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
        //                    HeaderRange.Font.Bold = true;

        //                    // DataCells
        //                    int RowsCount = DataTable.Rows.Count;
        //                    object[,] Cells = new object[RowsCount, ColumnsCount];

        //                    for (int j = 0; j < RowsCount; j++)
        //                        for (int i = 0; i < ColumnsCount; i++)
        //                            Cells[j, i] = DataTable.Rows[j][i];

        //                    Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[2, 1]),
        //                        (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[RowsCount + 1, ColumnsCount])).Value = Cells;
        //                    Worksheet.Columns[1].ColumnWidth = 13;
        //                    Worksheet.Columns[2].ColumnWidth = 13;
        //                    Worksheet.Columns[3].ColumnWidth = 11;
        //                    Worksheet.Columns[4].ColumnWidth = 16;
        //                    Worksheet.Columns[5].ColumnWidth = 14;
        //                    Worksheet.Columns[6].ColumnWidth = 20;
        //                    Worksheet.Columns[7].ColumnWidth = 15;
        //                    Worksheet.Columns[8].ColumnWidth = 11;
        //                    Worksheet.Columns[9].ColumnWidth = 13;
        //                    Worksheet.Columns[10].ColumnWidth = 14;
        //                    Worksheet.Columns[11].ColumnWidth = 9;
        //                    Worksheet.Columns[12].ColumnWidth = 13;
        //                    Worksheet.Columns[13].ColumnWidth = 13;
        //                    Worksheet.Columns[14].ColumnWidth = 15;
        //                    // check fielpath
        //                    if (ExcelFilePath != null && ExcelFilePath != "")
        //                    {
        //                        try
        //                        {
        //                            Worksheet.SaveAs(ExcelFilePath);
        //                          //  Worksheet.SaveAs(CopyExcelPath);
        //                            Excel.Quit();
        //                            lblInsert.Text = "" + filename + " saved successfully!";
        //                            //System.Windows.MessageBox.Show("Excel file saved!");
        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
        //                                + ex.Message);
        //                        }
        //                    }
        //                    else    // no filepath is given
        //                    {
        //                        Excel.Visible = true;
        //                    }
        //                }

        //                catch (Exception ex)
        //                {
        //                    throw new Exception("ExportToExcel: \n" + ex.Message);
        //                }
        //            }
        //            else
        //            {
        //                lblInsert.Text = "Excel file for the present data has been already generated! ";
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            throw new Exception("ExportToExcel failed due to: \n" + ex.Message);
        //        }
        //        con.Close();
        //    }

        //}


        protected void ExportCSV(DataTable dt, string scsvPath, string scsvcopypath, string csvFilename)
        {
            try { 

            using (SqlConnection con = new SqlConnection("Server = 192.9.200.45; Database = inforerpdb; User Id = dev1; Password = dev1; "))
            {
                int Filecount = 0;
                SqlCommand cmd2 = new SqlCommand();
                cmd2.Connection = con;
                SqlCommand cmd3 = new SqlCommand();
                cmd3.Connection = con;

                con.Open();
                cmd2.CommandText = @"SELECT count(distinct t_fnam)FROM ttfisg051200 where t_fnam not in ('')";
                Filecount = (int)cmd2.ExecuteScalar();
                Filecount++;

                try
                {
                    if (dt.Rows.Count > 0)
                    {
                        cmd3.CommandText = @"update ttfisg051200 set t_fnam = '" + csvFilename + "' , t_flid = " + Filecount + ", t_pmtf= '1' where t_fnam=''";
                        cmd3.ExecuteNonQuery();

                        //Build the CSV file data as a Comma separated string.
                        try
                        {
                            string csv = string.Empty;

                            foreach (DataColumn column in dt.Columns)
                            {
                                //Add the Header row for CSV file.
                                csv += column.ColumnName + ',';
                            }

                                //Microsoft.Office.Interop.Excel.Worksheet(" + csvFilename + ").Columns("A:M").AutoFit;
                                //Add new line.
                                csv += "\r\n";

                            foreach (DataRow row in dt.Rows)
                            {
                                foreach (DataColumn column in dt.Columns)
                                {
                                    //Add the Data rows.
                                    csv += row[column.ColumnName].ToString().Replace(",", ";") + ',';
                                }

                                //Add new line.
                                csv += "\r\n";
                            }
                            File.WriteAllText(scsvPath, csv.ToString());
                            File.WriteAllText(scsvcopypath, csv.ToString());
                            lblInsert.Text = "" + csvFilename + " saved successfully!";
                        }
                        catch (Exception ex)

                        {
                            cmd3.CommandText = @"update ttfisg051200 set t_fnam = '' , t_flid = '', t_pmtf= '2' where t_fnam='" + csvFilename + "'";
                            cmd3.ExecuteNonQuery();
                            lblInsert.Text = "Error occurs while saving the CSV file, Kindly check log file for details!";
                           // throw new Exception("ExportToCSV: \n" + ex.Message);
                            WriteToErrorLog(ex.Message,ex.StackTrace,ex.Source);
                          //  WriteToEventLog(ex.Source,ex.TargetSite,ex.GetHashCode,ex.Message);

                        }
                        //Download the CSV file.
                        //Response.Clear();
                        //Response.Buffer = true;
                        //Response.AddHeader("content-disposition", "attachment;filename=SqlExport.csv");
                        //Response.Charset = "";
                        //Response.ContentType = "application/text";
                        //Response.Output.Write(csv);
                        //Response.Flush();
                        //Response.End();
                        //        }
                        //    }
                        //}
                    }
                    else
                    {
                        lblInsert.Text = "CSV file for the present data has been already generated! ";
                    }

                }
                catch (Exception ex)
                {
                    lblInsert.Text = "Error occurs while saving the CSV file, Kindly check log file for details!";
                    // throw new Exception("ExportToCSV: \n" + ex.Message);
                    WriteToErrorLog(ex.Message, ex.StackTrace, ex.Source);
                    // throw new Exception("ExportToCSV: \n" + ex.Message);
                }
                con.Close();
            }
            }
            catch (Exception ex)
            {
                lblInsert.Text = "Error occurs while saving the CSV file, Kindly check log file for details!";
                // throw new Exception("ExportToCSV: \n" + ex.Message);
                WriteToErrorLog(ex.Message, ex.StackTrace, ex.Source);
                // throw new Exception("ExportToCSV: \n" + ex.Message);
            }
        }

        public void WriteToErrorLog(string msg, string stkTrace, string title)

        {
            FileStream fs1 = new FileStream(@"C:\host2host\documents\payments\out\h2h_payments\errlog.txt", FileMode.Append, FileAccess.Write);

            StreamWriter s1 = new StreamWriter(fs1);

            s1.Write("Title: " + title + System.Environment.NewLine);

            s1.Write("Message: " + msg + System.Environment.NewLine);

            s1.Write("StackTrace: " + stkTrace + System.Environment.NewLine);

            s1.Write("Date/Time: " + DateTime.Now.ToString() + System.Environment.NewLine);

            s1.Write("============================================" + System.Environment.NewLine);

            s1.Close();

            fs1.Close();

        }
      

    }

}
//}
//}

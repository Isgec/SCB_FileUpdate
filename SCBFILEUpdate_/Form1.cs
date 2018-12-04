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
using System.Text.RegularExpressions;

namespace SCBFILEUpdate_
{
public partial class Form1 : Form
{
    public Form1()
    {
        InitializeComponent();
    }
        //private void Form1_Shown(Object sender, EventArgs e)
        //{
        //    btnExport.PerformClick();
        //}
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            this.btnExport_Click(null, null);
            this.button1_Click(null, null);
        }
        private static void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
          Environment.Exit(0);
        }

    private void button1_Click(object sender, EventArgs e)
    {
        try
        {
            string[] txtFiles = Directory.GetFiles(@"C:\host2host\documents\payments\in\rpt\", "*.rpt");
            foreach (string currentFile in txtFiles)
            {
                string fileName = Path.GetFileName(currentFile);
                using (StreamReader sr = new StreamReader(File.Open(currentFile, FileMode.Open, FileAccess.Read)))
                {
                    //using (SqlConnection con = new SqlConnection("Server=192.9.200.45; Database=inforerpdb; User Id=dev1; Password=dev1;"))
                    using (SqlConnection con = new SqlConnection("Server=192.9.200.129; Database=inforerpdb; User Id=dev1; Password=Dev1@12345;"))
                    {
                        char[] fieldSep = new char[] { ',' };
                        char[] delim = new char[] { '"' };
                        con.Open();
                        string line = "";
                        string cmdTxt = "";
                        int count = 0;
                        int TotalHRecords = 0;

                        while ((line = sr.ReadLine()) != "" && line != null)
                        {
                                line= line.Replace("’", "");
                                line = line.Replace("'", "");
                                string[] parts = Regex.Split(line, ",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");
                            int nFields = parts.Length;
                            int RecordExist;
                            int TotalDRecords = 0;
                            if (parts[0].Trim(delim) == "H")
                            {
                                string recordexist = "select count(*) from ttfisg052200 where t_fnam in ('" + fileName.Trim(delim) + "')";
                                using (SqlCommand cmd = new SqlCommand(recordexist, con))
                                {
                                    if (con.State != ConnectionState.Open)
                                    {
                                        con.Open();
                                    }
                                    RecordExist = (int)(cmd.ExecuteScalar());
                                }
                                if (RecordExist == 0)
                                {
                                    string cmdCountTxt = "select count(*) from ttfisg052200 ";
                                    using (SqlCommand cmd = new SqlCommand(cmdCountTxt, con))
                                    {
                                        if (con.State != ConnectionState.Open)
                                        {
                                            con.Open();
                                        }
                                        count = (int)(cmd.ExecuteScalar());
                                        TotalHRecords = count + 1;
                                    }

                                    DateTime DtofPreparation = DateTime.ParseExact((parts[3].Trim(delim) + parts[4].Trim(delim)), "yyMMddHHmm", CultureInfo.InvariantCulture);

                                    cmdTxt = @"INSERT INTO ttfisg052200 (t_hsrn,t_date,t_user,t_t_no,t_fnam,t_flid,t_rtyp,t_sdid,t_rpid,t_dprp,t_utrn,
                                                t_ftyp,t_brcd,t_cus1,t_cus2,t_cus3,t_cus4,t_cus5,t_adr1,t_adr2,t_adr3,t_adr4,t_adr5,t_pcod
                                                ,t_ccod,t_cfr1,t_cfr2,t_cfr3,t_cfr4,t_t_rt,t_Refcntd,t_Refcntu,t_uerp,t_udon,t_udby)";

                                    cmdTxt += "VALUES (" + TotalHRecords + ",'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' ,'" + "3194" + "' ," + TotalHRecords + ",'" + fileName + "'," + TotalHRecords + "";
                                    cmdTxt += String.Format(@",'{0}','{1}','{2}','{3}','{4}' ,'{5}','{6}','{7}','{8}','{9}','{10}',
                                                '{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}',
                                                '{25}','{26}','{27}','{28}')", parts[0].Trim(delim), parts[1].Trim(delim), parts[2].Trim(delim),
                                            DtofPreparation.ToString("yyyy-MM-dd HH:mm:ss"), parts[5].Trim(delim),
                                            parts[6].Trim(delim), parts[7].Trim(delim), parts[8].Trim(delim),
                                            parts[9].Trim(delim), parts[10].Trim(delim), parts[11].Trim(delim),
                                            parts[12].Trim(delim), parts[13].Trim(delim), parts[14].Trim(delim),
                                            parts[15].Trim(delim), parts[16].Trim(delim), parts[17].Trim(delim),
                                            parts[18].Trim(delim), parts[19].Trim(delim), parts[20].Trim(delim),
                                            parts[21].Trim(delim), parts[22].Trim(delim), parts[23].Trim(delim), "H", 0, 0, 2, "", "");

                                    {
                                        SqlTransaction tran = con.BeginTransaction();
                                        try
                                        {
                                            if (con.State != ConnectionState.Open)
                                            {
                                                con.Open();
                                            }
                                            SqlCommand cmd = new SqlCommand(cmdTxt, con, tran);
                                            cmd.ExecuteNonQuery();
                                            tran.Commit();
                                        }
                                        catch (Exception ex)
                                        {
                                               
                                            WriteToErrorLog(ex.Message, ex.StackTrace, ex.Source);
                                            tran.Rollback();
                                        }
                                    }
                                }
                                else
                                {
                                    lblInsert.Text = "All the records for file " + fileName + " already been inserted!";
                                }
                            }
                            else if (parts[0].Trim(delim) == "D")
                            {
                                if (TotalHRecords != 0)
                                {
                                    string cmdTotalDCountTxt = "select count(*) from ttfisg053200 ";
                                    using (SqlCommand cmd = new SqlCommand(cmdTotalDCountTxt, con))
                                    {
                                        if (con.State != ConnectionState.Open)
                                        {
                                            con.Open();
                                        }
                                        count = (int)(cmd.ExecuteScalar());
                                        TotalDRecords = count + 1;
                                    }
                                        cmdTxt = @"INSERT INTO ttfisg053200 (t_hsrn,t_dsrn,t_rtyp,t_cusr,t_stat,t_stds,t_batr
                                            ,t_payr,t_pidt,t_pydt,t_dedt,t_pamt,t_pcur,t_scur,t_tcur,t_exrt,t_ecur,t_stdc
                                            ,t_chqn,t_cqdt,t_acno,t_acnm,t_benm,t_bad1,t_bad2,t_bad3,t_bad4,t_bcty,t_bpos,t_bccr,t_binf,t_flr1,t_flr2
                                            ,t_fl2a,t_flr3,t_flr4,t_flr5,t_flid,t_Refcntd,t_Refcntu,t_uerp,t_udat,t_udby,t_flr6,t_flr7,t_flr8,t_flr9) ";

                                    cmdTxt += "VALUES ( " + TotalHRecords + "," + TotalDRecords + ",";

                                       
                                        for (int i = 0; i < nFields; ++i)
                                        {
                                            //t_pidt--6,t_pydt--7,t_dedt--8,t_cqdt17,t_udat
                                            string field = "";
                                            if (parts[i].Trim(delim).ToString() == "")
                                            {
                                            }
                                            else
                                            {
                                               field = parts[i].Trim(delim);
                                            }
                                           
                                            if (i == 0)
                                                cmdTxt += String.Format("'{0}'", field);
                                            else
                                                cmdTxt += String.Format(", '{0}'", field);
                                        }
                                        cmdTxt += "," + TotalDRecords + ", 0,0,2, '' ,'',0,0,0,0 )";
                                        {
                                            SqlTransaction tran = con.BeginTransaction();
                                        try
                                        {
                                            if (con.State != ConnectionState.Open)
                                            {
                                                con.Open();
                                            }
                                            SqlCommand cmd = new SqlCommand(cmdTxt, con, tran);
                                            cmd.ExecuteNonQuery();
                                            tran.Commit();
                                        }
                                        catch (Exception ex)
                                        {
                                                WriteToErrorLog(fileName.Trim(delim), cmdTxt, "");
                                                WriteToErrorLog(ex.Message, ex.StackTrace, ex.Source);
                                            tran.Rollback();
                                        }

                                    }
                                    lblInsert.Text = "" + fileName.Trim(delim) + " records have been inserted!";
                                }
                                else
                                {

                                }
                            }
                            else if (parts[0].Trim(delim) == "T")
                            {
                                if (TotalHRecords != 0)
                                {
                                    char[] fieldSep1 = new char[] { ',' };
                                    char[] delim1 = new char[] { };
                                    parts = line.Split(fieldSep1);
                                    string sTUpdate = @"update ttfisg052200 set t_t_rt= 'T', t_t_no= " + parts[1].Trim(delim1) + " where t_fnam='" + fileName.Trim(delim) + "'";
                                    {
                                        SqlTransaction tran = con.BeginTransaction();
                                        try
                                        {
                                            if (con.State != ConnectionState.Open)
                                            {
                                                con.Open();
                                            }
                                            SqlCommand cmd = new SqlCommand(sTUpdate, con, tran);
                                            cmd.ExecuteNonQuery();
                                            tran.Commit();
                                        }
                                        catch (Exception ex)
                                        {
                                            WriteToErrorLog(ex.Message, ex.StackTrace, ex.Source);
                                            tran.Rollback();
                                        }
                                    }
                                }
                                else
                                {

                                }
                            }
                            else
                            {
                                if (con.State != ConnectionState.Closed)
                                {
                                    con.Close();
                                }
                                  
                            }

                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            WriteToErrorLog(ex.Message, ex.StackTrace, ex.Source);
            System.Timers.Timer timer = new System.Timers.Timer(5000);
            timer.Elapsed += Timer_Elapsed;
            timer.Start();
            lblInsert.Text = "Error occurs while loading Bank file into BaaN LN Table!";
        }
        finally
        {
               
            System.Timers.Timer timer = new System.Timers.Timer(5000);
            timer.Elapsed += Timer_Elapsed;
            timer.Start();
        }

    }
    public void btnExport_Click(object sender, EventArgs e)
    {
        DataTable dtnew = new DataTable();
      // using (SqlConnection con = new SqlConnection("Server=192.9.200.45; Database=inforerpdb; User Id=dev1; Password=dev1;"))
       using (SqlConnection con = new SqlConnection("Server=192.9.200.129; Database=inforerpdb; User Id=dev1; Password=Dev1@12345;"))
        {

            using (SqlCommand cmd = new SqlCommand("SP_SCBExportToCSV", con))
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

                        string sCSVFilePath = @"C:\host2host\documents\payments\out\isgec_h2h_payments\" + sFilename;
                        string sCSVFilePath2 = @"C:\host2host\documents\payments\out\h2h_payments\" + sFilename;
                        ExportCSV(dt, sCSVFilePath, sCSVFilePath2, sFilename);
                        System.Timers.Timer timer = new System.Timers.Timer(5000);
                        timer.Elapsed += Timer_Elapsed;
                        timer.Start();
                    }
                }
            }
        }
    }

    protected void ExportCSV(DataTable dt, string scsvPath, string scsvcopypath, string csvFilename)
    {
        try
        {
          // using (SqlConnection con = new SqlConnection("Server=192.9.200.45; Database=inforerpdb; User Id=dev1; Password=dev1;"))
            using (SqlConnection con = new SqlConnection("Server=192.9.200.129; Database=inforerpdb; User Id=dev1; Password=Dev1@12345;"))
            {
                int Filecount = 0;
                SqlCommand cmd2 = new SqlCommand();
                cmd2.Connection = con;
                SqlCommand cmd3 = new SqlCommand();
                cmd3.Connection = con;

                if (con.State != ConnectionState.Open)
                {
                    con.Open();
                }
                cmd2.CommandText = @"SELECT count(distinct t_fnam)FROM ttfisg051200 where t_fnam not in ('')";
                Filecount = (int)cmd2.ExecuteScalar();
                Filecount++;

                try
                {
                    if (dt.Rows.Count > 0)
                    {
                        cmd3.CommandText = @"update ttfisg051200 set t_fnam = '" + csvFilename + "' , t_flid = " + Filecount + " where t_fnam=''";
                        cmd3.ExecuteNonQuery();

                        try
                        {
                            string csv = string.Empty;
                            foreach (DataColumn column in dt.Columns)
                            {
                                csv += column.ColumnName + ',';
                            }
                            csv += "\r\n";
                            foreach (DataRow row in dt.Rows)
                            {
                                foreach (DataColumn column in dt.Columns)
                                {
                                    csv += row[column.ColumnName].ToString().Replace(",", ";") + ',';
                                }
                                csv += "\r\n";
                            }
                                
                            File.WriteAllText(scsvPath, csv.ToString());
                            File.WriteAllText(scsvcopypath, csv.ToString());

                            lblInsert.Text = "" + csvFilename + " saved successfully!";
                        }
                        catch (Exception ex)

                        {

                            cmd3.CommandText = @"update ttfisg051200 set t_fnam = '' , t_flid = '' where t_fnam='" + csvFilename + "'";
                            cmd3.ExecuteNonQuery();
                            File.WriteAllText(scsvPath, "");
                            lblInsert.Text = "Error occurs while saving the CSV file, Kindly check log file for details!";
                            WriteToErrorLog(ex.Message, ex.StackTrace, ex.Source);

                        }
                    }
                    else
                    {
                        lblInsert.Text = "CSV file for the present data has been already generated! ";
                        string sRepeatedEntry = "CSV file for the present data has been already generated! ";
                        WriteToErrorLog(sRepeatedEntry, "", "");
                    }

                }
                catch (Exception ex)
                {
                    lblInsert.Text = "Error occurs while saving the CSV file, Kindly check log file for details!";
                    WriteToErrorLog(ex.Message, ex.StackTrace, ex.Source);
                }
                con.Close();
            }
        }
        catch (Exception ex)
        {
            lblInsert.Text = "Error occurs while saving the CSV file, Kindly check log file for details!";
            WriteToErrorLog(ex.Message, ex.StackTrace, ex.Source);
        }
    }

    public void WriteToErrorLog(string msg, string stkTrace, string title)

    {
        FileStream fs1 = new FileStream(@"C:\host2host\errlog.txt", FileMode.Append, FileAccess.Write);

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


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb ;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Runtime.InteropServices;
using AXmsCtrl;
using System.IO;
using Microsoft.Win32;
namespace SMSAlertSystem
{
    public partial class frmSmsBlaster : Form
    {

      OleDbConnection CN ;
      SmsProtocolGsm objGsmOut ;
        //TimeSpan varDateTime = DateTime.Now.AddMinutes(1).TimeOfDay; 
        //TimeSpan varDateTime = DateTime.Now.TimeOfDay; 
        String spath = Application.StartupPath + "\\SMSLog.txt";
        bool ServerRunning = false;

        [DllImport("ODBCCP32.DLL")]
        private static extern int SQLConfigDataSource(int hwndParent, int fRequest, string lpszDriver, string lpszAttributes);



        public frmSmsBlaster()
        {
            InitializeComponent();
        }

        private void frmSmsBlaster_Load(object sender, EventArgs e)
        {

            
            try
            { 
                 RegistryKey  add = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);

            
                if ((add.GetValue("SMSAuto")== null) ) 
                {
                    add.SetValue("SMSAuto", Application.ExecutablePath.ToString());
                        }

            }
                    
           catch(Exception ex)
            {

                string smsg = ex.ToString();
            }
            
                  
           
     

        string varConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + Application.StartupPath + @"\db\SMSAlert.accdb';Jet OLEDB:Database Password='rdshmhadminff1c28c';"; 

        CN = new OleDbConnection() ;
        CN.ConnectionString = varConnectionString;

        //'===============CREAT DSN
        CreateDSN();
        //'====================CREATE DSN

  
            ////var STR = System.IO.File.ReadAllText(@"D:\SMSAlert\sms_File.json");
                // var str = OBJ.GetyEmployee();
                //  JObject obj1 = JObject.Parse(STR);
            ////DataTable myDataSet = JsonConvert.DeserializeObject<DataTable>(STR);
            /////DG1.DataSource = myDataSet;


       // Timer1.Interval = GetSMSTimeInterval();
        //Timer1.Enabled = true;

        }

        private void CreateDSN()
        {
            int ii = 0;
            char cc = (char)ii;

            //string sDriver = "Microsoft Access Driver (*.mdb,*.accdb)";
              string sDriver = "Microsoft Access Driver (*.mdb, *.accdb)";
            System.Text.StringBuilder sAttributes = new System.Text.StringBuilder();
            int ODBC_ADD_SYS_DSN = 4;
            long intResult;

            sAttributes.Append("DSN=SMSAlert");
            sAttributes.Append("\0");
            sAttributes.Append("DBQ=");
            sAttributes.Append(Application.StartupPath);
            sAttributes.Append(@"\db\SMSAlert.accdb");
            //sAttributes.Append("\0");

           // sAttributes.Append("\0");

            //SQLConfigDataSource(

            intResult = SQLConfigDataSource(0,ODBC_ADD_SYS_DSN,sDriver,sAttributes.ToString());

            
            sAttributes = null;
        }

        private void startSmsBlast()
        {


            objGsmOut = new SmsProtocolGsm();

            lblStatus.Text = "Checking Device Connectivity!";
            Application.DoEvents();

            for (int l = 0; l < objGsmOut.GetDeviceCount(); l++)
            {
                objGsmOut.Device = objGsmOut.GetDevice(l);
                break;
            }

            if (objGsmOut.Device == "" || CheckDeviceConnected() != true)
            {
                lblStatus.Text = "No Device Connected!";
                objGsmOut = null;
                return;
            }
            else
            {
                lblStatus.Text = "Device Connectivity Success!";
                Application.DoEvents();
            }



            lblStatus.Text = "Checking DB Connectivity!";
            Application.DoEvents();

            try
            {
                CN.Open();
                lblStatus.Text = "DB Connectivity Success!";
                Application.DoEvents();
                CN.Close();
            }
            catch (Exception ex)
            {
                lblStatus.Text = ex.Message.ToString();
                ;
                return;
            }



            lblStatus.Text = "Start Sending Previous Pending SMS!";
            string varBlastPendingSMS = BlastPendingSMS();
            if (varBlastPendingSMS != "OK")
            {
                lblStatus.Text = varBlastPendingSMS;
                return;
            }
            else
            {
                lblStatus.Text = "Pending SMS Blasting Completed Successfully!";
            }


            DG1.DataSource = null;
            DG1.Rows.Clear();



            lblStatus.Text = "Start Fetching Records!";
            Application.DoEvents();
            string varFilePath = Application.StartupPath + "\\"; //@"D:\SMSAlert\" ;
            string varFileName = "send_sms.json"; //AddZeroIfRequired(DateTime.Now .Day) + "_"  + AddZeroIfRequired(DateTime.Now .Month)  + "_" +  DateTime.Now .Year  +  "_SMS_File.json" ;
            string STR = "";
            DataTable myDataSet = null;

            try
            {
                 STR = System.IO.File.ReadAllText(varFilePath + "" + varFileName);
            }
            catch (Exception ex)
            {
                lblStatus.Text = ex.Message.ToString();
                return;
            }

            try
            {
                myDataSet = JsonConvert.DeserializeObject<DataTable>(STR);
                myDataSet.Columns.Add("Status", typeof(string));
              
            }
            catch (Exception ex)
            {
                lblStatus.Text = ex.Message.ToString();
                WriteErrorLog(spath, ex.Message + "      " + ex.ToString());
                return;
            }



            lblStatus.Text = "Start Sending Fresh SMS!";
            string varImportData = SendSMSAndImportData(myDataSet, varFileName);
            if (varImportData != "OK")
            {
                lblStatus.Text = varImportData;
                return;
            }
            else
            {
                lblStatus.Text = "SMS Blasting Completed Successfully!";
            }


            objGsmOut = null;

        }

        private string SendSMSAndImportData(DataTable DBDataTable,string varFileName)
        {
            string varReturn = "OK";
            try
            {
                //int varMaxTransID = GetMaxSMSTransId(varFileName);
                //DataRow[] varDataRow = DBDataTable.Select("TransID > " + GetMaxSMSTransId(varFileName));
                DataRow[] varDataRow = DBDataTable.Select("TransID <> '' " );
                DG1.DataSource = varDataRow.CopyToDataTable() ;

                OleDbCommand DBCommand ;
                string SQLString = "";

                for (int x = 0; x < varDataRow.Length; x++)
                {

                    DG1.Rows[x].Selected = true ;
                   

                    if (CheckIfTransIDAlreadyExist(varDataRow[x]["TransID"].ToString()) == false)
                    {
                        lblStatus.Text = "Sending (" + (x + 1) + " / " + varDataRow.Length + ") Sms Text To  Mr/Miss : " + varDataRow[x]["title"].ToString() + ". Contact Number : " + varDataRow[x]["Contact_Number"].ToString();
                        Application.DoEvents();

                        if (SendSMSText(varDataRow[x]["Contact_Number"].ToString(), varDataRow[x]["text"].ToString()) == true)
                        {
                            DG1[4, x].Value = "Success";
                            SQLString = "insert into SmsAlertLog values(" + varDataRow[x]["TransID"].ToString() + ",'" + varFileName + "','" + varDataRow[x]["title"].ToString() + "','" + varDataRow[x]["Contact_Number"].ToString() + "','" + varDataRow[x]["text"].ToString() + "',1,'Success','" + DateTime.Now + "')";
                        }
                        else
                        {
                            DG1[4, x].Value = "Failed";
                            SQLString = "insert into SmsAlertLog values(" + varDataRow[x]["TransID"].ToString() + ",'" + varFileName + "','" + varDataRow[x]["title"].ToString() + "','" + varDataRow[x]["Contact_Number"].ToString() + "','" + varDataRow[x]["text"].ToString() + "',0,'Failure','" + DateTime.Now + "')";
                        }

                        Application.DoEvents();

                        DBCommand = new OleDbCommand(SQLString, CN);

                        try
                        {
                            CN.Open();
                        }
                        catch (Exception ex)
                        {

                            WriteErrorLog(spath, ex.Message + "      " + ex.ToString());
                        }

                        DBCommand.ExecuteNonQuery();
                        DBCommand.Connection.Close();

                        try
                        {
                            CN.Close();
                        }
                        catch (Exception ex)
                        {

                            WriteErrorLog(spath, ex.Message + "      " + ex.ToString());

                        }
                    }
                    else
                    {
                        DG1[4, x].Value = "Existed";
                    }


                }


            }
            catch (Exception ex)
            {
                varReturn = ex.Message .ToString() ;
            }


            return varReturn;
        }

        private int GetMaxSMSTransId(string varFileName)
        {
            OleDbDataAdapter  DBAdapter ;
            DataSet DBDataSet ;
            string SQLString ;
            int varReturn = 0; 

            SQLString = "SELECT MAX(TransID) FROM SmsAlertLog where FileName = '" + varFileName + "'" ;

            DBAdapter = new OleDbDataAdapter(SQLString, CN);
            DBDataSet = new DataSet() ;
            DBAdapter.Fill(DBDataSet, "SmsAlertLog" ) ;

            if(DBDataSet.Tables[0].Rows[0][0] == DBNull.Value  )
            {
                varReturn =  0; 
            }
            else
            {
                varReturn = int.Parse(DBDataSet.Tables[0].Rows[0][0].ToString()) ; 
            }

            return varReturn ;
        }

        private string AddZeroIfRequired(int varInt)
        {
            if (varInt > 9)
                return varInt.ToString() ;
            else
                return "0" + varInt;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            DG1.DataSource = null  ;
            startSmsBlast();
        }

        private bool CheckDeviceConnected()
        {
           SmsMessage objSmsMessage = new SmsMessage();

           try
           {
               objGsmOut.DeviceSpeed = 9600;

               AXmsCtrl.SmsConstants objSmsConstants = new AXmsCtrl.SmsConstants();

               objSmsMessage.Format = objSmsConstants.asMESSAGEFORMAT_UNICODE_MULTIPART;

               objSmsMessage.Data = "Testing Deivce Connectivity!";

               objSmsMessage.Recipient = "+92" + SendSMSNo();
               objGsmOut.Send(objSmsMessage);

               if (objGsmOut.LastError == 0 || objGsmOut.LastError == 30350)
                   return true;
               else
                   return false;
           }
           catch (Exception ex)
           {
                WriteErrorLog(spath, ex.Message + "      " + ex.ToString());
                return false;
           }

       
        }

        private int SendSMSNo()
        {
            OleDbCommand cmd = new OleDbCommand();
            CN.Open();

            cmd.Connection = CN;
           
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT ClientMobileNo from SMSAlertNoConfig";
       int MobileNo=   int.Parse(cmd.ExecuteScalar().ToString());


            return MobileNo;

           
           

        }

        private bool SendSMSText(string varNumber , string varMsg)
        {
            SmsMessage objSmsMessage = new SmsMessage();

            try
            {
                objGsmOut.DeviceSpeed = 9600;

                AXmsCtrl.SmsConstants objSmsConstants = new AXmsCtrl.SmsConstants();

                objSmsMessage.Format = objSmsConstants.asMESSAGEFORMAT_UNICODE_MULTIPART;

                objSmsMessage.Data = varMsg;

                objSmsMessage.Recipient = "+92" + varNumber;
                objGsmOut.Send(objSmsMessage);

                if (objGsmOut.LastError == 0 || objGsmOut.LastError == 30350)
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                WriteErrorLog(spath, ex.Message + "      " + ex.ToString());
                return false;
            }


        }


        private bool CheckIfTransIDAlreadyExist(string varTransID)
        {
            OleDbDataAdapter DBAdapter;
            DataSet DBDataSet;
            string SQLString;
            bool varReturn = false;

            SQLString = "SELECT count(*) FROM SmsAlertLog where TransID = '" + varTransID + "'";

            DBAdapter = new OleDbDataAdapter(SQLString, CN);
            DBDataSet = new DataSet();
            DBAdapter.Fill(DBDataSet, "SmsAlertLog");

            if (DBDataSet.Tables[0].Rows[0][0].ToString()  != "0")
            {
                varReturn = true;
            }
            

            return varReturn;
        }


        private string BlastPendingSMS()
        {
            string varReturn = "OK";
          //  string varFileName = "";



            try
            {
                //int varMaxTransID = GetMaxSMSTransId(varFileName);
                //DataRow[] varDataRow = DBDataTable.Select("TransID > " + GetMaxSMSTransId(varFileName));
                DataSet DBDataSet = GetTodaysPendingSMS();

                if (DBDataSet != null && DBDataSet.Tables[0].Rows.Count > 0)
                {

                    DataRow[] varDataRow = DBDataSet.Tables[0].Select("TransID <> '' ");
                    DG1.DataSource = varDataRow.CopyToDataTable();


                    OleDbCommand DBCommand;
                    string SQLString = "";

                    for (int x = 0; x < varDataRow.Length; x++)
                    {

                        DG1.Rows[x].Selected = true;


                        lblStatus.Text = "Sending Previous Pending (" + (x + 1) + " / " + varDataRow.Length + ") Sms Text To  Mr/Miss : " + varDataRow[x]["title"].ToString() + ". Contact Number : " + varDataRow[x]["ContactNumber"].ToString();
                        Application.DoEvents();

                        if (SendSMSText(varDataRow[x]["ContactNumber"].ToString(), varDataRow[x]["MsgText"].ToString()) == true)
                        {

                            DG1[4, x].Value = "Success";
                            SQLString = "update SmsAlertLog set status = 1, StatusMsg = 'Success' where TransID = '" + varDataRow[x]["TransID"].ToString() + "'";
                        }
                        else
                        {
                            DG1[4, x].Value = "Failure";
                        }


                        Application.DoEvents();

                        DBCommand = new OleDbCommand(SQLString, CN);

                        try
                        {
                            CN.Open();
                        }
                        catch (Exception ex)
                        {
                            WriteErrorLog(spath, ex.Message + "      " + ex.ToString());
                        }

                        DBCommand.ExecuteNonQuery();
                        DBCommand.Connection.Close();

                        try
                        {
                            CN.Close();
                        }
                        catch (Exception ex)
                        {

                            WriteErrorLog(spath, ex.Message + "      " + ex.ToString());
                        }



                    }
                }

            }
            catch (Exception ex)
            {
                varReturn = ex.Message.ToString();
                WriteErrorLog(spath, ex.Message + "      " + ex.ToString());
            }


           return varReturn;
        }



        private DataSet GetTodaysPendingSMS()
        {
            
        OleDbDataAdapter DBAdapter ;
        DataSet DBDataSet ;


        DBAdapter = new OleDbDataAdapter("select TransID, title , ContactNumber , MsgText , 'Pending' as Status from SmsAlertLog where Status = 0 and day(SendDateTime) = " + DateTime.Now.Day + " and month(SendDateTime) = " + DateTime.Now.Month + " and year(SendDateTime) = " + DateTime.Now.Year, CN);
        DBDataSet = new DataSet() ;

        DBAdapter.Fill(DBDataSet);

      
         return DBDataSet;

   // Follow up with this Case ID
  //Case #38263509
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            
            //string StartTime = CboStartHours.Text + ":" + CboStartMinuts.Text;
            //string EndTime = CboEndHours.Text + ":" + CboEndMinuts.Text;

            ////TimeSpan varStartCurrentTime = System.DateTime.Now.TimeOfDay ;
           ////TimeSpan varEndCurrentTime = System.DateTime.Now.AddMinutes(60).TimeOfDay ;

            //TimeSpan start = new TimeSpan(int.Parse(CboStartHours.Text) , int.Parse(CboStartMinuts.Text) , 0);
            //TimeSpan end = new TimeSpan(int.Parse(CboEndHours.Text), int.Parse(CboEndMinuts.Text), 0);
           // TimeSpan varNow = DateTime.Now.TimeOfDay;

          
             //if ((varNow >= start) && (varNow < end))
             //if ((varDateTime >= varStartCurrentTime) && (varDateTime <= varEndCurrentTime) && ServerRunning == false)
            if (ServerRunning == false)
             {
                   
                            try
                            {
                                ServerRunning = true;
                                lblStatus.Text = "Timer Starts Successfully";
                                Cursor = Cursors.WaitCursor;
                                DG1.DataSource = null;
                                startSmsBlast();
                               
                            }
                
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message + " ", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                               WriteErrorLog(spath, ex.Message + "      " + ex.ToString());
                }

                            finally
                            {
                                Cursor = Cursors.Default;
                                //button1.Enabled = true;
                                ServerRunning = false;
                            }
               
            }
        }

        private int GetSMSTimeInterval()
        {
            OleDbDataAdapter DBAdapter;
            DataSet DBDataSet;
            string SQLString;
            int varReturn = 0;

            SQLString = "SELECT SMSMinutes FROM SMSConfiguration";

            DBAdapter = new OleDbDataAdapter(SQLString, CN);
            DBDataSet = new DataSet();
            DBAdapter.Fill(DBDataSet, "SMSConfiguration");

            if (DBDataSet.Tables[0].Rows.Count > 0)
            {
                varReturn = (int.Parse(DBDataSet.Tables[0].Rows[0][0].ToString()) * 60000);
            }
            else
            {
                varReturn = 60;
            }

            return varReturn;
        }

        private void WriteErrorLog(string spath, String LogsMessage)
        {
            System.IO.StreamWriter myStreamWriter = null;
            try
            {

                if (!(File.Exists(spath) == true))
                {
                    myStreamWriter = File.CreateText(spath);
                    myStreamWriter.WriteLine(System.DateTime.Now + "   " + LogsMessage);

                }
                else
                {
                    myStreamWriter = File.AppendText(spath);
                    myStreamWriter.WriteLine("------------------------------------------------");
                    myStreamWriter.WriteLine(System.DateTime.Now + "   " + LogsMessage);

                }

            }
            catch (Exception ex)
            {
                string smsg = ex.ToString();
            }
            finally
            {
                if (!(myStreamWriter == null))
                {
                    myStreamWriter.Flush();
                    myStreamWriter.Close();
                }

            }

        }
    }
}

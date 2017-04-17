using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Windows.Automation;

namespace LabelPrintMesMon
{
    public partial class frmMain : Form
    {
        private static string sConnString = "Data Source=10.217.170.10;Initial Catalog=aoidb;Persist Security Info=True;User ID=smt;Password=smt123;Connection Timeout=5"; //10.217.170.17
        //private static string sConnString = "Data Source=192.168.0.18;Initial Catalog=aoidb;Persist Security Info=True;User ID=smt;Password=smt123;Connection Timeout=5"; //10.217.170.17        
        private string sWO = "";
        private bool bMatikan = false;
        private string sCode = "";
        private int iWOQty = 0;
        private string sSerialFrom = "";
        private string sSerialToxx = "";
        private string CompName = "";

        public void SaveToLog(string strMessage)
        {

            string sYear = DateTime.Now.Year.ToString();
            string sMonth = DateTime.Now.Month.ToString();
            string sDay = DateTime.Now.Day.ToString();
            string sErrorTime = sYear + sMonth + sDay;
            string FileName = @Application.StartupPath + @"\" + sErrorTime + ".txt";
            try
            {

                // Store the script names and test results in a output text file.
                using (StreamWriter writer = new StreamWriter(new FileStream(FileName, FileMode.Append)))
                {
                    writer.Write("\r\n{0}{1}", DateTime.Now.ToString(), strMessage);
                }
            }
            catch (Exception ie)
            {
                return;
            }
        }


        private void SaveDataToRemote()
        {
            SqlConnection RemoteConnection = new SqlConnection(sConnString);
            string sStatus = "";
            try
            {
                RemoteConnection.Open();
            }
            catch (Exception ie)
            {
                //MessageBox.Show(ie.Message);
                //lblStatus.Text  = ie.Message.Substring(0, 60);
                SaveToLog("SaveDataToRemote Conn Open " + ie.Message.ToString());
                return;
            }

            try
            {
                SqlCommand RemoteCommand = new SqlCommand("InsertToWOSerial", RemoteConnection);
                RemoteCommand.CommandType = CommandType.StoredProcedure;
                RemoteCommand.Parameters.AddWithValue("@WONumber", sWO);
                RemoteCommand.Parameters.AddWithValue("@Code", sCode);
                RemoteCommand.Parameters.AddWithValue("@WOQqty", iWOQty);
                RemoteCommand.Parameters.AddWithValue("@SerialFrom", sSerialFrom);
                RemoteCommand.Parameters.AddWithValue("@SerialToxx", sSerialToxx);
                RemoteCommand.ExecuteNonQuery();
                //SqlDataReader reader = RemoteCommand.ExecuteReader();
            }
            catch (Exception ie)
            {

                //if (!ie.Message.ToLower().Contains("a duplicate value cannot be inserted"))
                //    MessageBox.Show(ie.Message);
                SaveToLog("SaveDataToRemote Execute " + ie.Message.ToString());
            }            

        }


        private void SaveMonitor(string strHasil)
        {
            SqlConnection RemoteConnection = new SqlConnection(sConnString);
            string sStatus = "";
            try
            {
                RemoteConnection.Open();
            }
            catch (Exception ie)
            {
                //MessageBox.Show(ie.Message);
                //lblStatus.Text  = ie.Message.Substring(0, 60);
                SaveToLog("SaveMonitor Conn Open " + ie.Message.ToString());
                return;
            }

            try
            {
                SqlCommand RemoteCommand = new SqlCommand("Insert Into MonitorProg(status,hasil,NamaKomp) Values(GetDate(),'"+ strHasil + "','" + CompName + "')", RemoteConnection);
                RemoteCommand.CommandType = CommandType.Text;
                RemoteCommand.ExecuteNonQuery();
                //SqlDataReader reader = RemoteCommand.ExecuteReader();
            }
            catch (Exception ie)
            {

                //if (!ie.Message.ToLower().Contains("a duplicate value cannot be inserted"))
                //    MessageBox.Show(ie.Message);
                SaveToLog("SaveMonitor Execute " + ie.Message.ToString());
            }

        }


        private void AddToList(string spWO,string spCode, int ipWOQty, string spSerialFrom, string spSerialToxx)
        {
            DateTime dtWaktu = DateTime.Now;
            ListViewItem item1 = new ListViewItem(spWO);
            item1.SubItems.Add(spCode);
            item1.SubItems.Add(ipWOQty.ToString());
            item1.SubItems.Add(spSerialFrom);
            item1.SubItems.Add(spSerialToxx);
            item1.SubItems.Add(dtWaktu.ToString("dd/MMM HH:mm:ss"));
            if (listView1.Items.Count > 10)
                listView1.Items[0].Remove();
            listView1.Items.Add(item1);
        }

        private void ChekData()
        {
             string spWO = "";
             string spCode = "";
             int ipWOQty = 0;
             string spSerialFrom = "";
             string spSerialToxx = "";
            PropertyCondition typeCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
            PropertyCondition nameCondition = new PropertyCondition(AutomationElement.NameProperty, "LABEL Print"); //
            AndCondition andCondition = new AndCondition(typeCondition, nameCondition);
            AutomationElement mainWnd = AutomationElement.RootElement.FindFirst(TreeScope.Children, andCondition);

            try
            {

                if (mainWnd != null)
                {
                    lblStatus.Text = "Running Label Print MES Found";

                    // FIND WORK ORDER
                    //1. Find First ControlType.Button With AutomationId= ArrowToggle
                    //2. Next Find Sibling, next after this is W/O (


/* OLD WAY                    
                    PropertyCondition type1Condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                    PropertyCondition Id1Condition = new PropertyCondition(AutomationElement.AutomationIdProperty, "ArrowToggle"); //             


                    AndCondition and1Condition = new AndCondition(type1Condition, Id1Condition);
                    AutomationElement elm1Found = mainWnd.FindFirst(TreeScope.Descendants, and1Condition);

                    if (elm1Found != null)
                    {
                        AutomationElement elm2Found = TreeWalker.ControlViewWalker.GetNextSibling(elm1Found);

                        if (elm2Found != null)
                        {
                            spWO = (string)elm2Found.GetCurrentPropertyValue(ValuePattern.ValueProperty);
                        }
                    }
 */


                    PropertyCondition type1Condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
                    PropertyCondition Id1Condition = new PropertyCondition(AutomationElement.AutomationIdProperty, "EditControl"); //             


                    AndCondition and1Condition = new AndCondition(type1Condition, Id1Condition);
                    AutomationElement elm1Found = mainWnd.FindFirst(TreeScope.Descendants, and1Condition);

                    if (elm1Found != null)
                    {
                        spWO = (string)elm1Found.GetCurrentPropertyValue(ValuePattern.ValueProperty);
                    }

//                    MessageBox.Show("WO:" + spWO.ToString());

                    if (spWO.Length == 0)
                        return;


                    PropertyCondition type2Condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
                    PropertyCondition Id2Condition = new PropertyCondition(AutomationElement.AutomationIdProperty, "txtModelSuffix"); //             


                    AndCondition and2Condition = new AndCondition(type2Condition, Id2Condition);
                    AutomationElement elm2Found = mainWnd.FindFirst(TreeScope.Descendants, and2Condition);

                    if (elm2Found != null)
                    {
                        spCode = (string)elm2Found.GetCurrentPropertyValue(ValuePattern.ValueProperty);
                    }
                   // MessageBox.Show("Prod Code:" + spCode.ToString());

                    PropertyCondition type3Condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
                    PropertyCondition Id3Condition = new PropertyCondition(AutomationElement.AutomationIdProperty, "txtWOQty"); //             


                    AndCondition and3Condition = new AndCondition(type3Condition, Id3Condition);
                    AutomationElement elm3Found = mainWnd.FindFirst(TreeScope.Descendants, and3Condition);

                    if (elm3Found != null)
                    {
                        string sTest = "";
                        int intParsed=0;
                        sTest = (string)elm3Found.GetCurrentPropertyValue(ValuePattern.ValueProperty);
                        if (int.TryParse(sTest.Trim(), out intParsed))
                        {
                            ipWOQty = intParsed;
                        }
                    }

                    //MessageBox.Show("WO Qty:" + ipWOQty.ToString());


                    // NEW WAY USING Find nearest, then use next sibling method
                    PropertyCondition type21Condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                    PropertyCondition Id21Condition = new PropertyCondition(AutomationElement.AutomationIdProperty, "textBlock44"); //             


                    AndCondition and21Condition = new AndCondition(type21Condition, Id21Condition);
                    AutomationElement elm21Found = mainWnd.FindFirst(TreeScope.Descendants, and21Condition);

                    if (elm21Found != null)
                    {
                        //MessageBox.Show(elm21Found.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty).ToString());
                        //MessageBox.Show(elm21Found.Current.Name);
                        AutomationElement elm22Found = TreeWalker.ControlViewWalker.GetNextSibling(elm21Found);

                        if (elm22Found == null)
                            return;
                        //MessageBox.Show(elm22Found.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty).ToString());
                        //MessageBox.Show(elm22Found.Current.Name);


                        AutomationElement elm23Found = TreeWalker.ControlViewWalker.GetNextSibling(elm22Found);

                        if (elm23Found == null)
                            return;

                        //MessageBox.Show(elm23Found.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty).ToString());
                        //MessageBox.Show(elm23Found.Current.Name);


                        AutomationElement elm24Found = TreeWalker.ControlViewWalker.GetNextSibling(elm23Found);

                        if (elm24Found == null)
                            return;

                        //MessageBox.Show(elm24Found.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty).ToString());
                        //MessageBox.Show(elm24Found.Current.Name);

                        AutomationElement elm25Found = TreeWalker.ControlViewWalker.GetNextSibling(elm24Found);

                        if (elm25Found == null)
                            return;

                        //MessageBox.Show(elm25Found.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty).ToString());
                        //MessageBox.Show(elm25Found.Current.Name);

                        AutomationElement elm26Found = TreeWalker.ControlViewWalker.GetNextSibling(elm25Found);

                        if (elm26Found == null)
                            return;

                        //MessageBox.Show(elm26Found.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty).ToString());
                        //MessageBox.Show(elm26Found.Current.Name);


                        AutomationElement elm27Found = TreeWalker.ControlViewWalker.GetNextSibling(elm26Found);


                        if (elm27Found == null)
                            return;

                        //MessageBox.Show(elm27Found.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty).ToString());
                        //MessageBox.Show(elm27Found.Current.Name);

                        AutomationElement elm28Found = TreeWalker.ControlViewWalker.GetNextSibling(elm27Found);

                        if (elm28Found == null)
                            return;

                        if (elm28Found.Current.Name.Length >= 3)
                        {
                            if (elm28Found.Current.Name.Substring(0, 1) == "A" || elm28Found.Current.Name.Substring(0, 1) == "E" || elm28Found.Current.Name.Substring(0, 1) == "I")
                            {
                                spSerialToxx = elm28Found.Current.Name;
                                AutomationElement elm29Found = TreeWalker.ControlViewWalker.GetNextSibling(elm28Found);
                                if (elm29Found != null)
                                    spSerialFrom = elm29Found.Current.Name;
                            }
                        }

                       // MessageBox.Show("Serial:" + spSerialFrom.ToString() + " To " + spSerialToxx.ToString());

                        //MessageBox.Show(elm28Found.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty).ToString());
                        // MessageBox.Show(elm28Found.Current.Name);
                        //spWO = (string)elm2Found.GetCurrentPropertyValue(ValuePattern.ValueProperty);

                    }

                    //MessageBox.Show(spWO + "," + spCode + "," + ipWOQty.ToString() + "," + spSerialFrom + "," + spSerialToxx);


                    // FIND SERIAL
                    //1. Find First ControlType.Button With AutomationId= ArrowToggle
                    //2. Next Find Sibling, next after this is W/O (
                    // OLD WAY USING FindAll Method... very slow and take big RAM
                    /*PropertyCondition type3Condition = new PropertyCondition(AutomationElement.ControlTypeProperty,ControlType.Text);
                    PropertyCondition Name3Condition = new PropertyCondition(AutomationElement.NameProperty, "Start Serial"); //             
                    AndCondition and3Condition = new AndCondition(type3Condition, Name3Condition);
                    AutomationElementCollection elmColFound = mainWnd.FindAll(TreeScope.Descendants, and3Condition);
                    foreach(AutomationElement elmTester in elmColFound)
                    {
                        if (elmTester != null)
                        {
                            AutomationElement elm4Found = TreeWalker.ControlViewWalker.GetNextSibling(elmTester);
                            if (elm4Found != null)
                            {
                                if (elm4Found.Current.Name.Substring(0, 3) == "MHM")
                                {
                                    spSerialToxx = elm4Found.Current.Name;
                                    AutomationElement elm5Found = TreeWalker.ControlViewWalker.GetNextSibling(elm4Found);
                                    if (elm5Found != null)
                                        spSerialFrom = elm5Found.Current.Name;                                    
                                    break;
                                }
                            }

                        }
                            
                    }*/

                    if (spWO.Length == 0 || spSerialFrom.Length == 0 || spSerialToxx.Length == 0)
                    {
                        return;
                    }


                    if (spWO != sWO || spSerialFrom != sSerialFrom || spSerialToxx != sSerialToxx)
                    {
                        AddToList(spWO, spCode, ipWOQty, spSerialFrom, spSerialToxx);
                        sWO = spWO;
                        sSerialFrom = spSerialFrom;
                        sSerialToxx = spSerialToxx;
                        sCode = spCode;
                        iWOQty = ipWOQty;
                        SaveDataToRemote();
                    }



                } //MainWindow
                else
                    lblStatus.Text = "Running MES Not Found";
            }
            catch (Exception ie)
            {
                //MessageBox.Show(ie.Message);
                SaveToLog("ChekData "+ie.Message.ToString());
                return;
            }

        }

        public frmMain()
        {
            InitializeComponent();
            listView1.Items.Clear();
            CompName = Environment.MachineName;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            ChekData();
            SaveMonitor(sWO + sSerialFrom + sSerialToxx);
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            ChekData();
            SaveMonitor(sWO + sSerialFrom + sSerialToxx);
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!bMatikan)
            {
                e.Cancel = true;
                WindowState = FormWindowState.Minimized;
            } 
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bMatikan = true;
        }
    }
}

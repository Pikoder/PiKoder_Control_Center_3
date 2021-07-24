
using System;
using System.Windows.Forms;
using NativeWifi;
using System.IO.Ports;
using System.Drawing;
using System.Globalization;
using System.Net.NetworkInformation;


// This is a program for evaluating the PiKoder platform - please refer to http://pikoder.com for more details.
// 
// Copyright 2015-2020 Gregor Schlechtriem
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
// http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

namespace PCCpro
{
    public partial class PCCpro : Form
    {

        private const string LogFileName = "\\PPC_LogFile.log";
        private const bool LogOutput = false;

        private PiKoderCommunicationAbstractionLayer myPCAL = new PiKoderCommunicationAbstractionLayer();

        private WlanClient wlanClient = new WlanClient();

        private bool boolErrorFlag; // global flag for errors in communication
        private bool IOSwitching = false; // SSC feature switch (implemented w/ release 2.7)
        private bool FastChannelRetrieve = false; // SSC feature switch (implemented w/ release 2.8)
        private bool ProtectedSaveMode = false; // SSC feature switch (implemented w/ release 2.9)
        private bool HPMath = false; // SSC HP feature switch to high precision computing
        private bool bDataLoaded = false; // flag to avoid data updates while uploading data from Pikoder

        private string strPiKoderType = ""; // PiKoder type we are currently connected to
        private string myMessage = ""; // used for error messages while connecting
        private int PPMmode;

        /* rework needed regarding migration of USB2PPM */
        private int iDefaultMinValue = 1000; // default values for USB2PMM
        private int iDefaultMaxValue = 2000;
        private bool bUART2PPM_StartUpValues = false;
        private bool PPMModeLegacy = true;

        private int[] iChannelSetting = new int[9]; // contains the current output type (would be 1 for P(WM) and 2 for S(witch)

        public PCCpro()
        {
            InitializeComponent();
            boolErrorFlag = false;
            ObtainCurrentSSID();
            UpdateCOMPortList();
            AvailableCOMPorts.Select();     // set focus in view
            AvailableCOMPorts.Focus();
        }


        //****************************************************************************
        //   Function:
        //       public void WriteLog(ByVal TimeDate As Date, ByVal Message As String)
        //
        //   Summary:
        //       This function appends message to standard logfile.
        //
        //   Description:
        //
        //   Precondition:
        //       None
        //
        //   Parameters:
        //       string Message - message to be added
        //
        //   Return Values
        //       None
        //
        //   Remarks:
        //       None
        //***************************************************************************
        public void WriteLog(string Message)
        {
            System.IO.StreamWriter myStreamWriter = System.IO.File.AppendText(Application.StartupPath + LogFileName);

            myStreamWriter.Write($"[{DateTime.Now.ToLongTimeString()} {DateTime.Now.ToLongDateString()}");
            myStreamWriter.WriteLine("] :" + Message);
            myStreamWriter.Close();
        }

        //****************************************************************************
        //   Function:
        //       public void ObtainedSSID()
        //
        //   Summary:
        //       If we are connected to an AP then function will display the SSID of this very AP.
        //       If we are not connected, the field will be blank. This also means that we can only 
        //       connect to an AP to which we have connected earlier on a W10 level.
        //
        //       The listbox will be updated controlled by a timmer so that you could change the AP
        //       w/o ending and restarting the PCC.
        //
        //   Description:
        //
        //   Precondition:
        //       None
        //
        //   Parameters:
        //       None
        //
        //   Return Values
        //       None
        //
        //   Remarks:
        //       None
        //***************************************************************************
        private void ObtainCurrentSSID()
        {
            if (wlanClient.Interfaces[0].InterfaceState == Wlan.WlanInterfaceState.Disconnected)
            {
                if (ConnectedAP.Text != "")
                {
                    ConnectedAP.Text = "";
                }
                return;
            }

            if (wlanClient.Interfaces[0].InterfaceState == Wlan.WlanInterfaceState.Connected)
            {
                try
                {
                    Wlan.WlanAvailableNetwork[] networks = wlanClient.Interfaces[0].GetAvailableNetworkList(0);
                    // Wlan.WlanAvailableNetwork[] networks = wlanClient.Interfaces[0].GetAvailableNetworkList((Wlan.WlanGetAvailableNetworkFlags.IncludeAllManualHiddenProfiles));
                    Wlan.Dot11Ssid ssid = networks[0].dot11Ssid;
                    string networkName = System.Text.Encoding.ASCII.GetString(ssid.SSID, 0, (int)ssid.SSIDLength);
                    if (!ConnectedAP.Text.Equals(networkName))
                    {
                        ConnectedAP.Text = networkName;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }


        //****************************************************************************
        //   Function:
        //       private void UpdateCOMPortList()
        //
        //   Summary:
        //       This function updates the COM ports listbox.
        //
        //   Description:
        //       This function updates the COM ports listbox.  This function is launched 
        //       periodically based on its Interval attribute (set in the form editor under
        //       the properties window).
        //
        //   Precondition:
        //       None
        //
        //   Parameters:
        //       None
        //
        //   Return Values
        //       None
        //
        //   Remarks:
        //       None
        //***************************************************************************
        private void UpdateCOMPortList()
        {
            int i = 0;
            int k;
            bool foundDifference = false;
            bool foundDifferenceHere;
            int iBufferSelectedIndex;
            string myStringBuffer;  // used for temporary storage of COM port designator

            // define and initialize array for sorting COM ports 
            string[] myCOMPorts = new string[99];  // assume we have a max of 100 COM Ports
            iBufferSelectedIndex = AvailableCOMPorts.SelectedIndex;  //buffer selection

            foreach (string s in SerialPort.GetPortNames())     // check all entries
            {
                foundDifferenceHere = true;
                for (k = 0; k < AvailableCOMPorts.Items.Count; k++)
                {
                    if (AvailableCOMPorts.Items[k].Equals(s))
                    {
                        foundDifferenceHere = false;
                    }
                }
                foundDifference = foundDifference || foundDifferenceHere;
            }
            if (!foundDifference)
            {
                return;
            }

            // If something has changed, then clear the list
            AvailableCOMPorts.Items.Clear();
            for (i = 0; i < 99; i++)
            {
                myCOMPorts[i] = "";
            }
            // Connection setup - check for ports and display result
            foreach (string sp in SerialPort.GetPortNames())
            {
                myStringBuffer = "";
                i = 0;
                for (k = 0; k < sp.Length; k++)
                {
                    if (((sp[k] >= 'A') & (sp[k] <= 'Z')) | Char.IsDigit(sp[k]))
                    {
                        myStringBuffer += sp[k];
                        if (Char.IsDigit(sp[k]))
                        {
                            i = (i * 10) + (sp[k] - '0');
                        }
                    }
                }
                myCOMPorts[i] = myStringBuffer;
            }

            for (i = 0; i < 99; i++)
            {
                if (myCOMPorts[i] != "")
                {
                    AvailableCOMPorts.Items.Add(myCOMPorts[i]);
                }
            }

            AvailableCOMPorts.SelectedIndex = iBufferSelectedIndex;
        }


        //****************************************************************************
        //   Function:
        //       private void timer1_Tick(object sender, EventArgs e)
        //
        //   Summary:
        //
        //   Description:
        //
        //   Precondition:
        //       None
        //
        //   Parameters:
        //       object sender     - Sender of the event (this form)
        //       EventArgs e       - The event arguments
        //
        //   Return Values
        //       None
        //
        //   Remarks:
        //       None
        //***************************************************************************/
        private void Timer1_Tick(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                // check if we still have a connection on USB system level
                if (!myPCAL.PiKoderConnected())
                {
                    LostConnection();
                }
            }
            else
            {
                // Update the COM ports list so that we can detect
                //  new COM ports that have been added.
                UpdateCOMPortList();
                // and update SSID
                ObtainCurrentSSID();
            }
        }

        private void IndicateConnectionOk()
        {
            TextBox1.Text = "Parameters loaded ok.";
            ledBulb1.Color = Color.Green;
            ledBulb1.Blink(0);
            ledBulb1.On = true;
        }

        private void RetrievePiKoderType(ref string SerialInputString)
        {
            string strChannelBuffer = "";
            if (myPCAL.LinkConnected())
            {
                if (!boolErrorFlag)
                {
                    // check for identifier
                    myPCAL.GetStatusRecord(ref strChannelBuffer);
                    if (strChannelBuffer == "TimeOut")
                    {
                        boolErrorFlag = true;
                    }
                    else if (strChannelBuffer == "?")
                    {
                        boolErrorFlag = true;
                    }
                    SerialInputString = strChannelBuffer;
                }
            }
        }

        //****************************************************************************
        //   Function:
        //       private void saveButton_Click(object Sender, EventArgs e)
        //
        //   Summary:
        //
        //   Description:
        //       This method initiates the the saving of the configuration of the PiKoder/SSC 
        //       upon hitting the 'save Parameters' - Button.
        //
        //   Precondition:
        //       None
        //
        //   Parameters:
        //       object sender     - Sender of the event (this form)
        //       EventArgs e       - The event arguments
        //
        //   Return Values
        //       None
        //
        //   Remarks:
        //       None
        //***************************************************************************/
        private void saveButton_Click(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                bool iRetCode = myPCAL.SetPiKoderPreferences(ProtectedSaveMode);

            }
        }

        // Subroutine for starting heartbeat timer
        private void startHeartBeat(int heartBeatInterval)
        {
            tHeartBeat.Enabled = true;
            tHeartBeat.Interval = heartBeatInterval;
        }

        // Subroutine for stopping heartbeat timer
        private void stopHeartBeat()
        {
            if (tHeartBeat.Enabled)
            {
                tHeartBeat.Enabled = false;
            }
        }

        // Handle Heartbeat ticks: check connection
        private void tHeartBeat_Tick() // Handles tHeartBeat.Tick
        {
            if (myPCAL.PiKoderConnected())
            {
                return;
            }
            LostConnection();
        }

        //****************************************************************************
        //   Function:
        //       private void CleanUpUI()
        //
        //   Summary:
        //
        //   Description:
        //       Set UI to indicate that we have lost connection - either on USB system 
        //       level or due to a time out condition on application level
        //
        //   Precondition:
        //       None
        //
        //   Parameters:
        //
        //   Return Values
        //       None
        //
        //   Remarks:
        //       None
        //***************************************************************************/
        private void CleanUpUI()
        {

            ledBulb1.Color = Color.Red;     // indicate that connection is lost
            ledBulb1.On = false;
            TextBox1.Text = "";
            stopHeartBeat();
            Timer1.Enabled = true;
            TypeId.Text = "";    // reset type information
            strPiKoderType = "";
            HPMath = false;
            bDataLoaded = false;
            bUART2PPM_StartUpValues = false;
            PPMModeLegacy = true;

            // delete text fields
            strCH_1_Current.Text = "";
            strCH_2_Current.Text = "";
            strCH_3_Current.Text = "";
            strCH_4_Current.Text = "";
            strCH_5_Current.Text = "";
            strCH_6_Current.Text = "";
            strCH_7_Current.Text = "";
            strCH_8_Current.Text = "";

            espFW.Text = "";

            ch1_HScrollBar.Enabled = false;     // hide and disable sliders
            ch1_HScrollBar.Visible = false;
            ch2_HScrollBar.Enabled = false;
            ch2_HScrollBar.Visible = false;
            ch3_HScrollBar.Enabled = false;
            ch3_HScrollBar.Visible = false;
            ch4_HScrollBar.Enabled = false;
            ch4_HScrollBar.Visible = false;
            ch5_HScrollBar.Enabled = false;
            ch5_HScrollBar.Visible = false;
            ch6_HScrollBar.Enabled = false;
            ch6_HScrollBar.Visible = false;
            ch7_HScrollBar.Enabled = false;
            ch7_HScrollBar.Visible = false;
            ch8_HScrollBar.Enabled = false;
            ch8_HScrollBar.Visible = false;

            strCH_1_Neutral.ForeColor = Color.White;
            strCH_2_Neutral.ForeColor = Color.White;
            strCH_3_Neutral.ForeColor = Color.White;
            strCH_4_Neutral.ForeColor = Color.White;
            strCH_5_Neutral.ForeColor = Color.White;
            strCH_6_Neutral.ForeColor = Color.White;
            strCH_7_Neutral.ForeColor = Color.White;
            strCH_8_Neutral.ForeColor = Color.White;

            strCH_1_Min.ForeColor = Color.White;
            strCH_2_Min.ForeColor = Color.White;
            strCH_3_Min.ForeColor = Color.White;
            strCH_4_Min.ForeColor = Color.White;
            strCH_5_Min.ForeColor = Color.White;
            strCH_6_Min.ForeColor = Color.White;
            strCH_7_Min.ForeColor = Color.White;
            strCH_8_Min.ForeColor = Color.White;

            strCH_1_Max.ForeColor = Color.White;
            strCH_2_Max.ForeColor = Color.White;
            strCH_3_Max.ForeColor = Color.White;
            strCH_4_Max.ForeColor = Color.White;
            strCH_5_Max.ForeColor = Color.White;
            strCH_6_Max.ForeColor = Color.White;
            strCH_7_Max.ForeColor = Color.White;
            strCH_8_Max.ForeColor = Color.White;

            strCH_1_Max.Enabled = true;
            strCH_2_Max.Enabled = true;
            strCH_3_Max.Enabled = true;
            strCH_4_Max.Enabled = true;
            strCH_5_Max.Enabled = true;
            strCH_6_Max.Enabled = true;
            strCH_7_Max.Enabled = true;
            strCH_8_Max.Enabled = true;

            strCH_1_Min.Enabled = true;
            strCH_2_Min.Enabled = true;
            strCH_3_Min.Enabled = true;
            strCH_4_Min.Enabled = true;
            strCH_5_Min.Enabled = true;
            strCH_6_Min.Enabled = true;
            strCH_7_Min.Enabled = true;
            strCH_8_Min.Enabled = true;

            TimeOut.ForeColor = Color.White;
            strSSC_Firmware.Text = " ";

            miniSSCOffset.ForeColor = Color.White;

            PPM_Channels.ForeColor = Color.White;
            PPM_Mode.ForeColor = Color.White;

            ListBox1.ForeColor = Color.White;
            ListBox2.ForeColor = Color.White;
            ListBox3.ForeColor = Color.White;
            ListBox4.ForeColor = Color.White;
            ListBox5.ForeColor = Color.White;
            ListBox6.ForeColor = Color.White;
            ListBox7.ForeColor = Color.White;
            ListBox8.ForeColor = Color.White;

            // clean up USB2PPM specifics
            GroupBox11.Enabled = true;     // neutral positions
            GroupBox11.Text = "neutral";
            GroupBox11.Visible = true;

            // clean up PRO specifics
            GroupBox13.Text = "PPM-Channels";

            // clean up PPM stuff
            GroupBox13.Enabled = true;      // PPM mode
            GroupBox13.Visible = true;
            GroupBox17.Enabled = true;      // PPM mode
            GroupBox17.Visible = true;

            // clean up SSCe stuff
            GroupBox4.Enabled = true;      // Time Out
            GroupBox4.Visible = true;
            GroupBox7.Enabled = true;      // miniSSC offset
            GroupBox7.Visible = true;
            GroupBox8.Enabled = true;      // Safe
            GroupBox8.Visible = true;


            // close port
            myPCAL.MyForm_Dispose();

            // reset error flag
            boolErrorFlag = false;
        }

        private void LostConnection()
        {
            if (ConnectCOM.Checked)
            {
                ConnectCOM.Checked = false;   // this will take care of the CleanUp of the UI
                TextBox1.Text = "Lost connection to PiKoder.";
            }
        }

        private void RetrieveSSC_HPParameters()
        {
            string strChannelBuffer = "";
            HPMath = true;
            IOSwitching = false;  // Better safe than sorry
            bDataLoaded = false;   // Avoid overridding of channel type due to re-reading data after value change
            if (myPCAL.LinkConnected())
            {
                if (!boolErrorFlag)
                {
                    // request status information from SSC    
                    myPCAL.GetFirmwareVersion(ref strChannelBuffer);
                    if (strChannelBuffer != "TimeOut")
                    {
                        strSSC_Firmware.Text = strChannelBuffer;
                        if (Double.Parse(strChannelBuffer, CultureInfo.InvariantCulture) > 2.04)
                        {
                            MessageBox.Show("The PiKoder firmware version found is not supported! Please goto www.pikoder.com and upgrade PCC Control Center to the latest version.", "Error Message", MessageBoxButtons.OK);
                            System.Windows.Forms.Application.Exit();
                        }
                        else if (Double.Parse(strChannelBuffer, CultureInfo.InvariantCulture) == 2.04)
                        {
                            IOSwitching = true;
                        }
                        else if (Double.Parse(strChannelBuffer) < 2.03)
                        {
                            MessageBox.Show("The PiKoder firmware version found is not supported! Please goto www.pikoder.com and upgrade the PiKoder firmware to the latest version.", "Error Message", MessageBoxButtons.OK);
                            System.Windows.Forms.Application.Exit();
                        }
                        else // error message
                            boolErrorFlag = true;
                    }
                    RetrievePiKoderParameters();
                }
            }
        }

        private void RetrievePiKoderParameters()
        {
            string strChannelBuffer = "";
            bDataLoaded = false;    // Avoid overridding of channel type due to re-reading data after value change

            GroupBox3.Invalidate();
            Refresh();

            GroupBox8.Enabled = true;
            GroupBox8.Visible = true;
            GroupBox4.Enabled = true;
            GroupBox4.Visible = true;
            GroupBox7.Enabled = true;        // Save Parameters
            GroupBox7.Visible = true;
            if (TypeId.Text != "SSC PRO")
            {
                GroupBox13.Enabled = false;      // # PPM Channels
                GroupBox13.Visible = false;
            }
            else
            {
                GroupBox13.Text = "I2C Address";
                GroupBox13.Enabled = true;
            }
            GroupBox17.Enabled = false;
            // PPM mode
            GroupBox17.Visible = false;

            if (TypeId.Text.Contains("SSCe"))
            {
                GroupBox4.Enabled = false;      // Time Out
                GroupBox4.Visible = false;
                if (TypeId.Text.Contains("(free)"))
                {
                    GroupBox7.Enabled = false;      // Safe
                    GroupBox7.Visible = false;
                    GroupBox8.Enabled = false;      // miniSSC
                    GroupBox8.Visible = false;
                }
            }

            // retrieve channel settings
            RetrieveChannel1Information(HPMath);
            frmOuter.Invalidate();
            Refresh();

            RetrieveChannel2Information(HPMath);
            frmOuter.Invalidate();
            Refresh();

            RetrieveChannel3Information(HPMath);
            frmOuter.Invalidate();
            Refresh();

            RetrieveChannel4Information(HPMath);
            frmOuter.Invalidate();
            Refresh();

            RetrieveChannel5Information(HPMath);
            frmOuter.Invalidate();
            Refresh();

            RetrieveChannel6Information(HPMath);
            frmOuter.Invalidate();
            Refresh();

            RetrieveChannel7Information(HPMath);
            frmOuter.Invalidate();
            Refresh();

            RetrieveChannel8Information(HPMath);
            frmOuter.Invalidate();
            Refresh();

            // retrieve TimeOut
            if (TypeId.Text.Contains("SSCe"))
            {
            } else {
                if (!boolErrorFlag)
                {
                    myPCAL.GetTimeOut(ref strChannelBuffer);
                    if (strChannelBuffer != "TimeOut")
                    {
                        TimeOut.Value = int.Parse(strChannelBuffer);
                        TimeOut.ForeColor = Color.Black;
                    }
                }
            }

            // retrieve miniSSC offset
            if (TypeId.Text.Contains("SSCe (free)"))
            {
            }
            else
            {
                if (!boolErrorFlag)
                {
                    myPCAL.GetMiniSSCOffset(ref strChannelBuffer);
                    if (strChannelBuffer != "TimeOut")
                    {
                        miniSSCOffset.Value = int.Parse(strChannelBuffer);
                        miniSSCOffset.ForeColor = Color.Black;
                    }
                }

                // retrieve PRO parameters offset
                if (TypeId.Text == "SSC PRO")
                {
                    myPCAL.GetISCBaseAddress(ref strChannelBuffer);
                    if (strChannelBuffer != "TimeOut")
                    {
                        PPM_Channels.Value = int.Parse(strChannelBuffer);
                        PPM_Channels.ForeColor = Color.Black;
                    }
                }
            }
            IndicateConnectionOk();
            bDataLoaded = true;
        }

        private void RetrieveSSCParameters()
        {
            string strChannelBuffer = "";
            IOSwitching = false;
            FastChannelRetrieve = false;

            if (myPCAL.LinkConnected())
            {
                if (!boolErrorFlag)
                {
                    // request status information from SSC    
                    myPCAL.GetFirmwareVersion(ref strChannelBuffer);
                    if (strChannelBuffer != "TimeOut")
                    {
                        strSSC_Firmware.Text = strChannelBuffer;
                        if (Double.Parse(strChannelBuffer, CultureInfo.InvariantCulture) > 3.01)
                        {
                            MessageBox.Show("The PiKoder firmware version found is not supported! Please goto www.pikoder.com and upgrade PCC Control Center to the latest version.", "Error Message", MessageBoxButtons.OK);
                            Application.Exit();
                        }
                        else if (Double.Parse(strChannelBuffer, CultureInfo.InvariantCulture) >= 2.09)
                        {
                            ProtectedSaveMode = true;
                            FastChannelRetrieve = true;
                            IOSwitching = true;
                            RetrievePiKoderParameters();
                        }
                        else if (Double.Parse(strChannelBuffer) >= 2.09)
                        {
                            FastChannelRetrieve = true;
                            IOSwitching = true;
                            RetrievePiKoderParameters();
                        }
                        else if (Double.Parse(strChannelBuffer) >= 2.07)
                        {
                            IOSwitching = true;
                            RetrievePiKoderParameters();
                        }
                        else if (Double.Parse(strChannelBuffer) < 2.0)
                        {
                            MessageBox.Show("The PiKoder firmware version found is not supported! Please goto www.pikoder.com and upgrade the PiKoder firmware to the latest version.", "Error Message", MessageBoxButtons.OK);
                            Application.Exit();
                        }
                        else  // error message
                        {
                            boolErrorFlag = true;
                        }
                    }
                }
            }
        }

        private void RetrieveSSCeParameters()
        {
            string strChannelBuffer = "";
            IOSwitching = false;
            FastChannelRetrieve = false;

            if (myPCAL.LinkConnected())
            {
                if (!boolErrorFlag)
                {
                    // request status information from SSC    
                    myPCAL.GetFirmwareVersion(ref strChannelBuffer);
                    if (strChannelBuffer != "TimeOut")
                    {
                        strSSC_Firmware.Text = strChannelBuffer;
                        if (Double.Parse(strChannelBuffer, CultureInfo.InvariantCulture) > 1.01)
                        {
                            MessageBox.Show("The PiKoder firmware version found is not supported! Please goto www.pikoder.com and upgrade PCC Control Center to the latest version.", "Error Message", MessageBoxButtons.OK);
                            Application.Exit();
                        }
                        else if (Double.Parse(strChannelBuffer, CultureInfo.InvariantCulture) >= 1.00)
                        {
                            ProtectedSaveMode = true;
                            FastChannelRetrieve = true;
                            IOSwitching = true;
                            RetrievePiKoderParameters();
                        }
                        else  // error message
                        {
                            boolErrorFlag = true;
                        }
                    }
                }
            }
        }

        private void RetrieveSSCeDEMOParameters()
        {
            string strChannelBuffer = "";
            IOSwitching = false;
            FastChannelRetrieve = false;

            if (myPCAL.LinkConnected())
            {
                if (!boolErrorFlag)
                {
                    // request status information from SSC    
                    myPCAL.GetFirmwareVersion(ref strChannelBuffer);
                    if (Double.Parse(strChannelBuffer, CultureInfo.InvariantCulture) > 1.00)
                    {
                        MessageBox.Show("The PiKoder firmware version found is not supported! Please goto www.pikoder.com and upgrade PCC Control Center to the latest version.", "Error Message", MessageBoxButtons.OK);
                        Application.Exit();
                    }
                    else if (Double.Parse(strChannelBuffer, CultureInfo.InvariantCulture) == 1.00)
                    {
                        ProtectedSaveMode = true;
                        FastChannelRetrieve = true;
                        IOSwitching = true;
                    }
                    else  // error message
                    {
                        boolErrorFlag = true;
                    }
                    RetrievePiKoderParameters();
                }
            }
        }

        private void RetrieveSSC_PROParameters()
        {
            string strChannelBuffer = "";
            IOSwitching = false;
            FastChannelRetrieve = false;

            if (myPCAL.LinkConnected())
            {
                if (!boolErrorFlag)
                {
                    // request status information from SSC    
                    myPCAL.GetFirmwareVersion(ref strChannelBuffer);
                    if (strChannelBuffer != "TimeOut")
                    {
                        strSSC_Firmware.Text = strChannelBuffer;
                        if (Double.Parse(strChannelBuffer, CultureInfo.InvariantCulture) != 1.02)
                        {
                            MessageBox.Show("The PiKoder firmware version found is not supported! Please goto www.pikoder.com and upgrade PCC Control Center to the latest version.", "Error Message", MessageBoxButtons.OK);
                            Application.Exit();
                        }
                        ProtectedSaveMode = true;
                        FastChannelRetrieve = true;
                        IOSwitching = true;
                        RetrievePiKoderParameters();
                    }
                    else  // error message
                    {
                        boolErrorFlag = true;
                    }
                }
            }
        }

        private void RetrieveUART2PPMParameters()
        {
            string strChannelBuffer = "";
            IOSwitching = false;
            FastChannelRetrieve = false;
            if (myPCAL.LinkConnected())
            {
                if (!boolErrorFlag)
                {
                    // request status information from SSC    
                    myPCAL.GetFirmwareVersion(ref strChannelBuffer);
                    if (strChannelBuffer != "TimeOut")
                    {
                        strSSC_Firmware.Text = strChannelBuffer;
                        if (Double.Parse(strChannelBuffer, CultureInfo.InvariantCulture) > 2.06)
                        {
                            MessageBox.Show("The PiKoder firmware version found is not supported! Please goto www.pikoder.com and upgrade PCC Control Center to the latest version.", "Error Message", MessageBoxButtons.OK);
                            Application.Exit();
                        }
                    }
                    else    // error message
                    {
                        boolErrorFlag = true;
                    }

                }
                RetrievePiKoderParameters();
            }
        }

        private string FormatChannelValue(decimal iChannelInput)
        {
            string strChannelBuffer = "";
            decimal iChannelValue = iChannelInput;
            if (HPMath)
            {
                iChannelValue *= 5;
                if (iChannelValue < 10000)
                {
                    strChannelBuffer += "0";
                }
            }
            if (iChannelValue < 1000)
            {
                strChannelBuffer += "0";
            }
            return strChannelBuffer + Convert.ToString(iChannelValue);
        }

        private void strCH_1_Neutral_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    myPCAL.SetChannelNeutral(FormatChannelValue(strCH_1_Neutral.Value), 1);
                }
            }

        }

        private void strCH_2_Neutral_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    myPCAL.SetChannelNeutral(FormatChannelValue(strCH_2_Neutral.Value), 2);
                }
            }
        }

        private void strCH_3_Neutral_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    myPCAL.SetChannelNeutral(FormatChannelValue(strCH_3_Neutral.Value), 3);
                }
            }
        }

        private void strCH_4_Neutral_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    myPCAL.SetChannelNeutral(FormatChannelValue(strCH_4_Neutral.Value), 4);
                }
            }
        }

        private void strCH_5_Neutral_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    myPCAL.SetChannelNeutral(FormatChannelValue(strCH_5_Neutral.Value), 5);
                }
            }
        }

        private void strCH_6_Neutral_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    myPCAL.SetChannelNeutral(FormatChannelValue(strCH_6_Neutral.Value), 6);
                }
            }
        }
        private void strCH_7_Neutral_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    myPCAL.SetChannelNeutral(FormatChannelValue(strCH_7_Neutral.Value), 7);
                }
            }
        }

        private void strCH_8_Neutral_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    myPCAL.SetChannelNeutral(FormatChannelValue(strCH_8_Neutral.Value), 8);
                }
            }
        }

        private void strCH_1_Min_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    if (ch1_HScrollBar.Value < strCH_1_Min.Value)
                    {
                        ch1_HScrollBar.Value = Convert.ToInt32(strCH_1_Min.Value);
                        strCH_1_Current.Text = Convert.ToString(ch1_HScrollBar.Value);
                    }
                    ch1_HScrollBar.Minimum = Convert.ToInt32(strCH_1_Min.Value);
                    myPCAL.SetChannelLowerLimit(FormatChannelValue(strCH_1_Min.Value), 1);
                }
            }
        }

        private void strCH_2_Min_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    if (ch2_HScrollBar.Value < strCH_2_Min.Value)
                    {
                        ch2_HScrollBar.Value = Convert.ToInt32(strCH_2_Min.Value);
                        strCH_2_Current.Text = Convert.ToString(ch2_HScrollBar.Value);
                    }
                    ch2_HScrollBar.Minimum = Convert.ToInt32(strCH_2_Min.Value);
                    myPCAL.SetChannelLowerLimit(FormatChannelValue(strCH_2_Min.Value), 2);
                }
            }
        }

        private void strCH_3_Min_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    if (ch3_HScrollBar.Value < strCH_3_Min.Value)
                    {
                        ch3_HScrollBar.Value = Convert.ToInt32(strCH_3_Min.Value);
                        strCH_3_Current.Text = Convert.ToString(ch3_HScrollBar.Value);
                    }
                    ch3_HScrollBar.Minimum = Convert.ToInt32(strCH_3_Min.Value);
                    myPCAL.SetChannelLowerLimit(FormatChannelValue(strCH_3_Min.Value), 3);
                }
            }
        }

        private void strCH_4_Min_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    if (ch4_HScrollBar.Value < strCH_4_Min.Value)
                    {
                        ch4_HScrollBar.Value = Convert.ToInt32(strCH_4_Min.Value);
                        strCH_4_Current.Text = Convert.ToString(ch4_HScrollBar.Value);
                    }
                    ch4_HScrollBar.Minimum = Convert.ToInt32(strCH_4_Min.Value);
                    myPCAL.SetChannelLowerLimit(FormatChannelValue(strCH_4_Min.Value), 4);
                }
            }
        }

        private void strCH_5_Min_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    if (ch5_HScrollBar.Value < strCH_5_Min.Value)
                    {
                        ch5_HScrollBar.Value = Convert.ToInt32(strCH_5_Min.Value);
                        strCH_5_Current.Text = Convert.ToString(ch5_HScrollBar.Value);
                    }
                    ch5_HScrollBar.Minimum = Convert.ToInt32(strCH_5_Min.Value);
                    myPCAL.SetChannelLowerLimit(FormatChannelValue(strCH_5_Min.Value), 5);
                }
            }
        }

        private void strCH_6_Min_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    if (ch6_HScrollBar.Value < strCH_6_Min.Value)
                    {
                        ch6_HScrollBar.Value = Convert.ToInt32(strCH_6_Min.Value);
                        strCH_6_Current.Text = Convert.ToString(ch6_HScrollBar.Value);
                    }
                    ch6_HScrollBar.Minimum = Convert.ToInt32(strCH_6_Min.Value);
                    myPCAL.SetChannelLowerLimit(FormatChannelValue(strCH_6_Min.Value), 6);
                }
            }
        }

        private void strCH_7_Min_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    if (ch7_HScrollBar.Value < strCH_7_Min.Value)
                    {
                        ch7_HScrollBar.Value = Convert.ToInt32(strCH_7_Min.Value);
                        strCH_7_Current.Text = Convert.ToString(ch7_HScrollBar.Value);
                    }
                    ch7_HScrollBar.Minimum = Convert.ToInt32(strCH_7_Min.Value);
                    myPCAL.SetChannelLowerLimit(FormatChannelValue(strCH_7_Min.Value), 7);
                }
            }
        }

        private void strCH_8_Min_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    if (ch8_HScrollBar.Value < strCH_8_Min.Value)
                    {
                        ch8_HScrollBar.Value = Convert.ToInt32(strCH_8_Min.Value);
                        strCH_8_Current.Text = Convert.ToString(ch8_HScrollBar.Value);
                    }
                    ch8_HScrollBar.Minimum = Convert.ToInt32(strCH_8_Min.Value);
                    myPCAL.SetChannelLowerLimit(FormatChannelValue(strCH_8_Min.Value), 8);
                }
            }
        }

        private void strCH_1_Max_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    if (ch1_HScrollBar.Value > strCH_1_Max.Value)
                    {
                        ch1_HScrollBar.Value = Convert.ToInt32(strCH_1_Max.Value);
                        strCH_1_Current.Text = Convert.ToString(ch1_HScrollBar.Value);
                    }
                    ch1_HScrollBar.Maximum = Convert.ToInt32(strCH_1_Max.Value);
                    myPCAL.SetChannelUpperLimit(FormatChannelValue(strCH_1_Max.Value), 1);
                }
            }
        }

        private void strCH_2_Max_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    if (ch2_HScrollBar.Value > strCH_2_Max.Value)
                    {
                        ch2_HScrollBar.Value = Convert.ToInt32(strCH_2_Max.Value);
                        strCH_2_Current.Text = Convert.ToString(ch2_HScrollBar.Value);
                    }
                    ch2_HScrollBar.Maximum = Convert.ToInt32(strCH_2_Max.Value);
                    myPCAL.SetChannelUpperLimit(FormatChannelValue(strCH_2_Max.Value), 2);
                }
            }
        }

        private void strCH_3_Max_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    if (ch3_HScrollBar.Value > strCH_3_Max.Value)
                    {
                        ch3_HScrollBar.Value = Convert.ToInt32(strCH_3_Max.Value);
                        strCH_3_Current.Text = Convert.ToString(ch3_HScrollBar.Value);
                    }
                    ch3_HScrollBar.Maximum = Convert.ToInt32(strCH_3_Max.Value);
                    myPCAL.SetChannelUpperLimit(FormatChannelValue(strCH_3_Max.Value), 3);
                }
            }
        }

        private void strCH_4_Max_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    if (ch4_HScrollBar.Value > strCH_4_Max.Value)
                    {
                        ch4_HScrollBar.Value = Convert.ToInt32(strCH_4_Max.Value);
                        strCH_4_Current.Text = Convert.ToString(ch4_HScrollBar.Value);
                    }
                    ch4_HScrollBar.Maximum = Convert.ToInt32(strCH_4_Max.Value);
                    myPCAL.SetChannelUpperLimit(FormatChannelValue(strCH_4_Max.Value), 4);
                }
            }
        }

        private void strCH_5_Max_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    if (ch5_HScrollBar.Value > strCH_5_Max.Value)
                    {
                        ch5_HScrollBar.Value = Convert.ToInt32(strCH_5_Max.Value);
                        strCH_5_Current.Text = Convert.ToString(ch5_HScrollBar.Value);
                    }
                    ch5_HScrollBar.Maximum = Convert.ToInt32(strCH_5_Max.Value);
                    myPCAL.SetChannelUpperLimit(FormatChannelValue(strCH_5_Max.Value), 5);
                }
            }
        }

        private void strCH_6_MAx_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    if (ch6_HScrollBar.Value > strCH_6_Max.Value)
                    {
                        ch6_HScrollBar.Value = Convert.ToInt32(strCH_6_Max.Value);
                        strCH_6_Current.Text = Convert.ToString(ch6_HScrollBar.Value);
                    }
                    ch6_HScrollBar.Maximum = Convert.ToInt32(strCH_6_Max.Value);
                    myPCAL.SetChannelUpperLimit(FormatChannelValue(strCH_6_Max.Value), 6);
                }
            }
        }

        private void strCH_7_Max_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    if (ch7_HScrollBar.Value > strCH_7_Max.Value)
                    {
                        ch7_HScrollBar.Value = Convert.ToInt32(strCH_7_Max.Value);
                        strCH_7_Current.Text = Convert.ToString(ch7_HScrollBar.Value);
                    }
                    ch7_HScrollBar.Maximum = Convert.ToInt32(strCH_7_Max.Value);
                    myPCAL.SetChannelUpperLimit(FormatChannelValue(strCH_7_Max.Value), 7);
                }
            }
        }

        private void strCH_8_Max_ValueChanged(object Sender, EventArgs e)
        {
            if (myPCAL.LinkConnected())
            {
                if (bDataLoaded)
                {
                    if (ch8_HScrollBar.Value > strCH_8_Max.Value)
                    {
                        ch8_HScrollBar.Value = Convert.ToInt32(strCH_8_Max.Value);
                        strCH_8_Current.Text = Convert.ToString(ch8_HScrollBar.Value);
                    }
                    ch8_HScrollBar.Maximum = Convert.ToInt32(strCH_8_Max.Value);
                    myPCAL.SetChannelUpperLimit(FormatChannelValue(strCH_8_Max.Value), 8);
                }
            }
        }

        private void ch1_HScrollBar_Scroll(object sender, ScrollEventArgs e)
        {
            strCH_1_Current.Text = Convert.ToString(ch1_HScrollBar.Value);
            if (HPMath)
            {
                myPCAL.SetHPChannelPulseLength(1, Convert.ToString(ch1_HScrollBar.Value * 5));
            }
            else
            {
                myPCAL.SetChannelPulseLength(1, Convert.ToString(ch1_HScrollBar.Value));
            }
        }

        private void ch2_HScrollBar_Scroll(object sender, ScrollEventArgs e)
        {
            strCH_2_Current.Text = Convert.ToString(ch2_HScrollBar.Value);
            if (HPMath)
            {
                myPCAL.SetHPChannelPulseLength(2, Convert.ToString(ch2_HScrollBar.Value * 5));
            }
            else
            {
                myPCAL.SetChannelPulseLength(2, Convert.ToString(ch2_HScrollBar.Value));
            }
        }

        private void ch3_HScrollBar_Scroll(object sender, ScrollEventArgs e)
        {
            strCH_3_Current.Text = Convert.ToString(ch3_HScrollBar.Value);
            if (HPMath)
            {
                myPCAL.SetHPChannelPulseLength(3, Convert.ToString(ch3_HScrollBar.Value * 5));
            }
            else
            {
                myPCAL.SetChannelPulseLength(3, Convert.ToString(ch3_HScrollBar.Value));
            }
        }

        private void ch4_HScrollBar_Scroll(object sender, ScrollEventArgs e)
        {
            strCH_4_Current.Text = Convert.ToString(ch4_HScrollBar.Value);
            if (HPMath)
            {
                myPCAL.SetHPChannelPulseLength(4, Convert.ToString(ch4_HScrollBar.Value * 5));
            }
            else
            {
                myPCAL.SetChannelPulseLength(4, Convert.ToString(ch4_HScrollBar.Value));
            }
        }

        private void ch5_HScrollBar_Scroll(object sender, ScrollEventArgs e)
        {
            strCH_5_Current.Text = Convert.ToString(ch5_HScrollBar.Value);
            if (HPMath)
            {
                myPCAL.SetHPChannelPulseLength(5, Convert.ToString(ch5_HScrollBar.Value * 5));
            }
            else
            {
                myPCAL.SetChannelPulseLength(5, Convert.ToString(ch5_HScrollBar.Value));
            }
        }

        private void ch6_HScrollBar_Scroll(object sender, ScrollEventArgs e)
        {
            strCH_6_Current.Text = Convert.ToString(ch6_HScrollBar.Value);
            if (HPMath)
            {
                myPCAL.SetHPChannelPulseLength(6, Convert.ToString(ch6_HScrollBar.Value * 5));
            }
            else
            {
                myPCAL.SetChannelPulseLength(6, Convert.ToString(ch6_HScrollBar.Value));
            }
        }

        private void ch7_HScrollBar_Scroll(object sender, ScrollEventArgs e)
        {
            strCH_7_Current.Text = Convert.ToString(ch7_HScrollBar.Value);
            if (HPMath)
            {
                myPCAL.SetHPChannelPulseLength(7, Convert.ToString(ch7_HScrollBar.Value * 5));
            }
            else
            {
                myPCAL.SetChannelPulseLength(7, Convert.ToString(ch7_HScrollBar.Value));
            }
        }

        private void ch8_HScrollBar_Scroll(object sender, ScrollEventArgs e)
        {
            strCH_8_Current.Text = Convert.ToString(ch8_HScrollBar.Value);
            if (HPMath)
            {
                myPCAL.SetHPChannelPulseLength(8, Convert.ToString(ch8_HScrollBar.Value * 5));
            }
            else
            {
                myPCAL.SetChannelPulseLength(8, Convert.ToString(ch8_HScrollBar.Value));
            }
        }

        private void TimeOut_ValueChanged(object sender, EventArgs e)
        {
            string myStringBuffer = "";
            if (myPCAL.LinkConnected())
            {
                if (TimeOut.Value < 10)
                {
                    myStringBuffer = "0";
                }
                if (TimeOut.Value < 100)
                {
                    myStringBuffer += "0";
                }
                myPCAL.SetPiKoderTimeOut(myStringBuffer + Convert.ToString(TimeOut.Value));
                if (TimeOut.Value != 0)
                {
                    startHeartBeat(Convert.ToInt32(TimeOut.Value) * 100 / 2); // set to shorter interval to make sure to account for admin
                }
                else
                {
                    stopHeartBeat();
                }
            }
        }

        private void miniSSCOffset_ValueChanged(object sender, EventArgs e)
        {
            if (bDataLoaded)
            {
                string myStringBuffer = "";
                if (myPCAL.LinkConnected())
                {
                    if (miniSSCOffset.Value < 100)
                    {
                        myStringBuffer = "0";
                    }
                    if (miniSSCOffset.Value < 10)
                    {
                        myStringBuffer += "0";
                    }
                    myPCAL.SetPiKoderMiniSSCOffset(myStringBuffer + Convert.ToString(miniSSCOffset.Value));
                }
            }
        }

        private void PPM_Channels_ValueChanged(object sender, EventArgs e)
        {
            if (bDataLoaded)
            {
                if (TypeId.Text != "SSC PRO")
                {
                    if (Convert.ToInt32(PPM_Channels.Value) < 1)
                    {
                        PPM_Channels.Value = 1;
                    }
                    else if (Convert.ToInt32(PPM_Channels.Value) > 8)
                    {
                        PPM_Channels.Value = 8;
                    }
                    else
                    {
                        myPCAL.SetPiKoderPPMSettings(Convert.ToInt32(PPM_Channels.Value), PPMmode);
                    }
                }
                else
                {
                    myPCAL.SetPiKoderI2CBaseAdress(Convert.ToInt32(PPM_Channels.Value));
                }
            }
        }

        private void PPM_Mode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (bDataLoaded)
            {
                if (PPM_Mode.SelectedIndex >= 0)
                {
                    myPCAL.SetPiKoderPPMSettings(Convert.ToInt32(PPM_Channels.Value), PPM_Mode.SelectedIndex);
                    PPMmode = PPM_Mode.SelectedIndex;
                    PPM_Mode.ClearSelected();
                }
            }
        }

        private int ScalePulseWidth(string strChannelBuffer, bool HPMath)
        {
            if (HPMath)
            {
                return int.Parse(strChannelBuffer) / 5;
            }
            else
            {
                return int.Parse(strChannelBuffer);
            }
        }

        private int EvaluateIOType(string strChannelBuffer)
        {
            if (String.Compare(strChannelBuffer, "P") == 0)
            {
                return 0;
            }
            else
            {
                return 1;
            }
        }

        private void RetrieveChannel1Information(bool HPMath)
        {
            string strChannelBuffer = "";
            if (!boolErrorFlag)
            {   // retrieve channel value
                myPCAL.GetPulseLength(ref strChannelBuffer, 1, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    ch1_HScrollBar.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_1_Current.Text = Convert.ToString(ch1_HScrollBar.Value);
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {   // retrieve neutral value
                myPCAL.GetNeutralPosition(ref strChannelBuffer, 1, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    strCH_1_Neutral.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_1_Neutral.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {   // retrieve lower limit 
                myPCAL.GetLowerLimit(ref strChannelBuffer, 1, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    int min = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_1_Min.Value = min;
                    ch1_HScrollBar.Minimum = min;
                    strCH_1_Min.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {   // retrieve upper limit 
                myPCAL.GetUpperLimit(ref strChannelBuffer, 1, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    int max = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_1_Max.Value = max;
                    ch1_HScrollBar.Maximum = max;
                    strCH_1_Max.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }

            if (IOSwitching)
            {
                if (!boolErrorFlag)
                {
                    myPCAL.GetIOType(ref strChannelBuffer, 1);
                    if (strChannelBuffer == "TimeOut")
                    {
                        boolErrorFlag = true;
                    }
                    else
                    {
                        ListBox1.Enabled = true;
                        ListBox1.SelectedIndex = EvaluateIOType(strChannelBuffer);
                        ListBox1.ClearSelected();
                        ListBox1.ForeColor = Color.Black;
                    }
                }
            }
        }

        private void RetrieveChannel2Information(bool HPMath)
        {
            string strChannelBuffer = "";
            if (!boolErrorFlag)
            {
                myPCAL.GetPulseLength(ref strChannelBuffer, 2, HPMath);
                {
                    if (strChannelBuffer != "TimeOut")
                    {
                        ch2_HScrollBar.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                        strCH_2_Current.Text = Convert.ToString(ch2_HScrollBar.Value);
                    }
                    else
                    {
                        boolErrorFlag = true;
                    }
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetNeutralPosition(ref strChannelBuffer, 2, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    strCH_2_Neutral.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_2_Neutral.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetLowerLimit(ref strChannelBuffer, 2, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    int min = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_2_Min.Value = min;
                    ch2_HScrollBar.Minimum = min;
                    strCH_2_Min.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetUpperLimit(ref strChannelBuffer, 2, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    int max = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_2_Max.Value = max;
                    ch2_HScrollBar.Maximum = max;
                    strCH_2_Max.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (IOSwitching)
            {
                if (!boolErrorFlag)
                {
                    myPCAL.GetIOType(ref strChannelBuffer, 2);
                    if (strChannelBuffer == "TimeOut")
                    {
                        boolErrorFlag = true;
                    }
                    else
                    {
                        ListBox2.Enabled = true;
                        ListBox2.SelectedIndex = EvaluateIOType(strChannelBuffer);
                        ListBox2.ClearSelected();
                        ListBox2.ForeColor = Color.Black;
                    }
                }
            }
        }

        private void RetrieveChannel3Information(bool HPMath)
        {
            string strChannelBuffer = "";
            if (!boolErrorFlag)
            {
                myPCAL.GetPulseLength(ref strChannelBuffer, 3, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    ch3_HScrollBar.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_3_Current.Text = Convert.ToString(ch3_HScrollBar.Value);
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetNeutralPosition(ref strChannelBuffer, 3, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    strCH_3_Neutral.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_3_Neutral.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetLowerLimit(ref strChannelBuffer, 3, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    int min = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_3_Min.Value = min;
                    ch3_HScrollBar.Minimum = min;
                    strCH_3_Min.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetUpperLimit(ref strChannelBuffer, 3, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    int max = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_3_Max.Value = max;
                    ch3_HScrollBar.Maximum = max;
                    strCH_3_Max.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (IOSwitching)
            {
                if (!boolErrorFlag)
                {
                    myPCAL.GetIOType(ref strChannelBuffer, 3);
                    if (strChannelBuffer == "TimeOut")
                    {
                        boolErrorFlag = true;
                    }
                    else
                    {
                        ListBox3.Enabled = true;
                        ListBox3.SelectedIndex = EvaluateIOType(strChannelBuffer);
                        ListBox3.ClearSelected();
                        ListBox3.ForeColor = Color.Black;
                    }
                }
            }
        }
        private void RetrieveChannel4Information(bool HPMath)
        {
            string strChannelBuffer = "";
            if (!boolErrorFlag)
            {
                myPCAL.GetPulseLength(ref strChannelBuffer, 4, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    ch4_HScrollBar.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_4_Current.Text = Convert.ToString(ch4_HScrollBar.Value);
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetNeutralPosition(ref strChannelBuffer, 4, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    strCH_4_Neutral.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_4_Neutral.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetLowerLimit(ref strChannelBuffer, 4, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    int min = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_4_Min.Value = min;
                    ch4_HScrollBar.Minimum = min;
                    strCH_4_Min.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetUpperLimit(ref strChannelBuffer, 4, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    int max = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_4_Max.Value = max;
                    ch4_HScrollBar.Maximum = max;
                    strCH_4_Max.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (IOSwitching)
            {
                if (!boolErrorFlag)
                {
                    myPCAL.GetIOType(ref strChannelBuffer, 4);
                    if (strChannelBuffer == "TimeOut")
                    {
                        boolErrorFlag = true;
                    }
                    else
                    {
                        ListBox4.Enabled = true;
                        ListBox4.SelectedIndex = EvaluateIOType(strChannelBuffer);
                        ListBox4.ClearSelected();
                        ListBox4.ForeColor = Color.Black;
                    }
                }
            }
        }

        private void RetrieveChannel5Information(bool HPMath)
        {
            string strChannelBuffer = "";
            if (!boolErrorFlag)
            {
                myPCAL.GetPulseLength(ref strChannelBuffer, 5, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    ch5_HScrollBar.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_5_Current.Text = Convert.ToString(ch5_HScrollBar.Value);
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetNeutralPosition(ref strChannelBuffer, 5, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    strCH_5_Neutral.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_5_Neutral.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetLowerLimit(ref strChannelBuffer, 5, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    int min = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_5_Min.Value = min;
                    ch5_HScrollBar.Minimum = min;
                    strCH_5_Min.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetUpperLimit(ref strChannelBuffer, 5, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    int max = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_5_Max.Value = max;
                    ch5_HScrollBar.Maximum = max;
                    strCH_5_Max.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (IOSwitching)
            {
                if (!boolErrorFlag)
                {
                    myPCAL.GetIOType(ref strChannelBuffer, 5);
                    if (strChannelBuffer == "TimeOut")
                    {
                        boolErrorFlag = true;
                    }
                    else
                    {
                        ListBox5.Enabled = true;
                        ListBox5.SelectedIndex = EvaluateIOType(strChannelBuffer);
                        ListBox5.ClearSelected();
                        ListBox5.ForeColor = Color.Black;
                    }
                }
            }
        }

        private void RetrieveChannel6Information(bool HPMath)
        {
            string strChannelBuffer = "";
            if (!boolErrorFlag)
            {
                myPCAL.GetPulseLength(ref strChannelBuffer, 6, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    ch6_HScrollBar.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_6_Current.Text = Convert.ToString(ch6_HScrollBar.Value);
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetNeutralPosition(ref strChannelBuffer, 6, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    strCH_6_Neutral.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_6_Neutral.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetLowerLimit(ref strChannelBuffer, 6, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    int min = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_6_Min.Value = min;
                    ch6_HScrollBar.Minimum = min;
                    strCH_6_Min.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetUpperLimit(ref strChannelBuffer, 6, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    int max = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_6_Max.Value = max;
                    ch6_HScrollBar.Maximum = max;
                    strCH_6_Max.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (IOSwitching)
            {
                if (!boolErrorFlag)
                {
                    myPCAL.GetIOType(ref strChannelBuffer, 6);
                    if (strChannelBuffer == "TimeOut")
                    {
                        boolErrorFlag = true;
                    }
                    else
                    {
                        ListBox6.Enabled = true;
                        ListBox6.SelectedIndex = EvaluateIOType(strChannelBuffer);
                        ListBox6.ClearSelected();
                        ListBox6.ForeColor = Color.Black;
                    }
                }
            }
        }

        private void RetrieveChannel7Information(bool HPMath)
        {
            string strChannelBuffer = "";
            if (!boolErrorFlag)
            {
                myPCAL.GetPulseLength(ref strChannelBuffer, 7, HPMath);
                if (strChannelBuffer != "TimeOut")
                {

                    ch7_HScrollBar.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_7_Current.Text = Convert.ToString(ch7_HScrollBar.Value);
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetNeutralPosition(ref strChannelBuffer, 7, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    strCH_7_Neutral.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_7_Neutral.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetLowerLimit(ref strChannelBuffer, 7, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    int min = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_7_Min.Value = min;
                    ch7_HScrollBar.Minimum = min;
                    strCH_7_Min.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetUpperLimit(ref strChannelBuffer, 7, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    int max = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_7_Max.Value = max;
                    ch7_HScrollBar.Maximum = max;
                    strCH_7_Max.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (IOSwitching)
            {
                if (!boolErrorFlag)
                {
                    myPCAL.GetIOType(ref strChannelBuffer, 7);
                    if (strChannelBuffer == "TimeOut")
                    {
                        boolErrorFlag = true;
                    }
                    else
                    {
                        ListBox7.Enabled = true;
                        ListBox7.SelectedIndex = EvaluateIOType(strChannelBuffer);
                        ListBox7.ClearSelected();
                        ListBox7.ForeColor = Color.Black;
                    }
                }
            }
        }

        private void RetrieveChannel8Information(bool HPMath)
        {
            string strChannelBuffer = "";
            if (!boolErrorFlag)
            {
                myPCAL.GetPulseLength(ref strChannelBuffer, 8, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    ch8_HScrollBar.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_8_Current.Text = Convert.ToString(ch8_HScrollBar.Value);
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetNeutralPosition(ref strChannelBuffer, 8, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    strCH_8_Neutral.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_8_Neutral.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetLowerLimit(ref strChannelBuffer, 8, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    int min = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_8_Min.Value = min;
                    ch8_HScrollBar.Minimum = min;
                    strCH_8_Min.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (!boolErrorFlag)
            {
                myPCAL.GetUpperLimit(ref strChannelBuffer, 8, HPMath);
                if (strChannelBuffer != "TimeOut")
                {
                    int max = ScalePulseWidth(strChannelBuffer, HPMath);
                    strCH_8_Max.Value = max;
                    ch8_HScrollBar.Maximum = max;
                    strCH_8_Max.ForeColor = Color.Black;
                }
                else
                {
                    boolErrorFlag = true;
                }
            }
            if (IOSwitching)
            {
                if (!boolErrorFlag)
                {
                    myPCAL.GetIOType(ref strChannelBuffer, 8);
                    if (strChannelBuffer == "TimeOut")
                    {
                        boolErrorFlag = true;
                    }
                    else
                    {
                        ListBox8.Enabled = true;
                        ListBox8.SelectedIndex = EvaluateIOType(strChannelBuffer);
                        ListBox8.ClearSelected();
                        ListBox8.ForeColor = Color.Black;
                    }
                }
            }
        }

        // The following block of routines handles the I/O type changes for the PiKoder/SSC
        // The form loads a specific listbox eventserver which then calls the actual generic sevrice
        private void ListBox1_SelectedIndexChanged(object Sender, EventArgs e)
        {
            if (ListBox1.SelectedIndex < 0)
            {
                return;
            }
            OutputTypeChange(1, ListBox1.SelectedIndex);
            ListBox1.SelectedItem = -1;
        }

        private void ListBox2_SelectedIndexChanged(object Sender, EventArgs e)
        {
            if (ListBox2.SelectedIndex < 0)
            {
                return;
            }
            OutputTypeChange(2, ListBox2.SelectedIndex);
            ListBox2.SelectedItem = -1;
        }
        private void ListBox3_SelectedIndexChanged(object Sender, EventArgs e)
        {
            if (ListBox3.SelectedIndex < 0)
            {
                return;
            }
            OutputTypeChange(3, ListBox3.SelectedIndex);
            ListBox3.SelectedItem = -1;
        }

        private void ListBox4_SelectedIndexChanged(object Sender, EventArgs e)
        {
            if (ListBox4.SelectedIndex < 0)
            {
                return;
            }
            OutputTypeChange(4, ListBox4.SelectedIndex);
            ListBox4.SelectedItem = -1;
        }

        private void ListBox5_SelectedIndexChanged(object Sender, EventArgs e)
        {
            if (ListBox5.SelectedIndex < 0)
            {
                return;
            }
            OutputTypeChange(5, ListBox5.SelectedIndex);
            ListBox5.SelectedItem = -1;
        }

        private void ListBox6_SelectedIndexChanged(object Sender, EventArgs e)
        {
            if (ListBox6.SelectedIndex < 0)
            {
                return;
            }
            OutputTypeChange(6, ListBox6.SelectedIndex);
            ListBox6.SelectedItem = -1;
        }

        private void ListBox7_SelectedIndexChanged(object Sender, EventArgs e)
        {
            if (ListBox7.SelectedIndex < 0)
            {
                return;
            }
            OutputTypeChange(7, ListBox7.SelectedIndex);
            ListBox7.SelectedItem = -1;
        }

        private void ListBox8_SelectedIndexChanged(object Sender, EventArgs e)
        {
            if (ListBox8.SelectedIndex < 0)
            {
                return;
            }
            OutputTypeChange(8, ListBox8.SelectedIndex);
            ListBox8.SelectedItem = -1;
        }

        // <summary>
        // The following Sub handles the communication to the PiKoder regarding I/O type changes
        // </summary>
        // <param name="iChannelNo"> channel number to be served </param>
        // <param name="iSettingIndex"> selected index representing output type </param>
        private void OutputTypeChange(int iChannelNo, int iSettingIndex)
        {
            if (!bDataLoaded)
            {
                return;
            }
            if (strPiKoderType.Contains("SSC"))
            {
                return;
            }
            if (!boolErrorFlag)
            {
                if (iSettingIndex == 0)
                {
                    iChannelSetting[iChannelNo] = 0;
                    boolErrorFlag = !myPCAL.SetChannelOutputType(iChannelNo, "P");
                }
                else
                {
                    iChannelSetting[iChannelNo] = 1;
                    boolErrorFlag = !myPCAL.SetChannelOutputType(iChannelNo, "S");
                }
            }
        }

        private void Connect2PiKoder(PiKoderCommunicationAbstractionLayer.iPhysicalLink iLinkType)
        {
            ledBulb1.On = true;     // indicate connection testing
            GroupBox2.Invalidate();
            Refresh();
            boolErrorFlag = !(myPCAL.EstablishLink(AvailableCOMPorts.Items[AvailableCOMPorts.TopIndex].ToString(), iLinkType));
            if (!boolErrorFlag)
            {
                RetrievePiKoderType(ref strPiKoderType);
                if (strPiKoderType.Contains("UART2PPM"))
                {
                    TypeId.Text = "UART2PPM";
                    DisplayFoundMessage(TypeId.Text);
                    RetrieveUART2PPMParameters();
                }
                else if (strPiKoderType.Contains("USB2PPM"))
                {
                    TypeId.Text = "USB2PPM";
                    DisplayFoundMessage(TypeId.Text);
                    RetrieveUSB2PPMParameters();
                }
                else if (strPiKoderType.Contains("SSC-HP"))
                {
                    TypeId.Text = "SSC-HP";
                    DisplayFoundMessage(TypeId.Text);
                    RetrieveSSC_HPParameters();
                }
                else if (strPiKoderType.Contains("SSC PRO"))
                {
                    TypeId.Text = "SSC PRO";
                    DisplayFoundMessage(TypeId.Text);
                    RetrieveSSC_PROParameters();
                }
                else if (strPiKoderType.Contains("SSCe (free)"))
                {
                    TypeId.Text = "SSCe (free)";
                    DisplayFoundMessage(TypeId.Text);
                    RetrieveSSCeDEMOParameters();
                }
                else if (strPiKoderType.Contains("SSCe"))
                {
                    TypeId.Text = "SSCe";
                    DisplayFoundMessage(TypeId.Text);
                    RetrieveSSCeParameters();
                }
                else if (strPiKoderType.Contains("SSC"))
                {
                    TypeId.Text = "SSC";
                    DisplayFoundMessage(TypeId.Text);
                    RetrieveSSCParameters();
                }
                else
                {
                    ledBulb1.On = false;
                    boolErrorFlag = true;
                    if (iLinkType == PiKoderCommunicationAbstractionLayer.iPhysicalLink.iSerialLink)
                    {
                        myMessage = "Device on " + AvailableCOMPorts.Items[AvailableCOMPorts.TopIndex].ToString() + " not supported";
                    }
                    else
                    {
                        myMessage = "Device on " + ConnectedAP.Text + " not supported";
                    }
                }
                ch1_HScrollBar.Enabled = true;     // show and enable sliders
                ch1_HScrollBar.Visible = true;
                ch2_HScrollBar.Enabled = true;
                ch2_HScrollBar.Visible = true;
                ch3_HScrollBar.Enabled = true;
                ch3_HScrollBar.Visible = true;
                ch4_HScrollBar.Enabled = true;
                ch4_HScrollBar.Visible = true;
                ch5_HScrollBar.Enabled = true;
                ch5_HScrollBar.Visible = true;
                ch6_HScrollBar.Enabled = true;
                ch6_HScrollBar.Visible = true;
                ch7_HScrollBar.Enabled = true;
                ch7_HScrollBar.Visible = true;
                ch8_HScrollBar.Enabled = true;
                ch8_HScrollBar.Visible = true;
            }
            else
            {
                Timer1.Enabled = true;
                ledBulb1.Blink(0);
                if (ConnectCOM.Checked)
                {
                    myMessage = "Could not open " + AvailableCOMPorts.Items[AvailableCOMPorts.TopIndex].ToString();
                }
                else
                {
                    myMessage = "Could not connect to " + ConnectedAP.Text;
                }
            }
        }

        private void DisplayFoundMessage(string sPiKoderType)
        {
            string myMessage = "Found ";
            myMessage = myMessage + sPiKoderType + " @ ";
            if (ConnectCOM.Checked)
            {
                myMessage += AvailableCOMPorts.Items[AvailableCOMPorts.TopIndex].ToString();
            }
            else
            {
                myMessage += ConnectedAP.Text;
            }
            TextBox1.Text = myMessage;
        }

        private void ConnectCOM_CheckedChanged(object Sender, EventArgs e)
        {
            TextBox1.Text = "";  // delete any prior error message
            if (ConnectCOM.Checked)
            {
                if (ConnectWLAN.Checked)
                {
                    ConnectWLAN.Checked = false;
                }
                Connect2PiKoder(PiKoderCommunicationAbstractionLayer.iPhysicalLink.iSerialLink);
                if (boolErrorFlag)
                {
                    ConnectCOM.Checked = false;
                }
            }
            else
            {
                myPCAL.DisconnectLink(PiKoderCommunicationAbstractionLayer.iPhysicalLink.iSerialLink);
                if (!boolErrorFlag) // just a simple disconnect, no error...
                {
                    myMessage = TypeId.Text + "@" + AvailableCOMPorts.Items[AvailableCOMPorts.TopIndex].ToString() + " disconnected";
                }
                CleanUpUI();
                TextBox1.Text = myMessage;
            }
        }

        private void ConnectWLAN_CheckedChanged(object Sender, EventArgs e)
        {
            string myMessage = "";
            TextBox1.Text = "";  // delete any prior error message
            if (ConnectWLAN.Checked)
            {
                if (ConnectedAP.Text == "")
                {   // make sure we are connected to an AP
                    myMessage = "Not connected to a WLAN";
                    ConnectWLAN.Checked = false;
                }
                else
                {
                    if (ConnectCOM.Checked)
                    {
                        ConnectCOM.Checked = false;
                    }
                    Connect2PiKoder(PiKoderCommunicationAbstractionLayer.iPhysicalLink.iWLANlink);
                    if (boolErrorFlag)
                    {
                        ConnectWLAN.Checked = false;
                    } else
                    {
                        string vString = "";
                        myPCAL.GetESPFirmwareVersion(ref vString);
                        espFW.Text = vString;
                    }
                }
            }
            else
            {
                myPCAL.DisconnectLink(PiKoderCommunicationAbstractionLayer.iPhysicalLink.iWLANlink);
                if (boolErrorFlag)
                {
                    myMessage = "Access Point " + ConnectedAP.Text + " not supported";
                }
                CleanUpUI();
                TextBox1.Text = myMessage;
            }
        }

        private void AvailableCOMPorts_SelectedIndexChanged(object Sender, EventArgs e)
        {
            AvailableCOMPorts.ClearSelected();
        }

        private void RetrieveUSB2PPMParameters()
        {
            string strChannelBuffer = "";
            if (myPCAL.LinkConnected())
            {
                if (!boolErrorFlag)
                {   // request firm ware version from PiKoder    
                    myPCAL.GetFirmwareVersion(ref strChannelBuffer);
                    if (strChannelBuffer != "TimeOut")
                    {
                        strSSC_Firmware.Text = strChannelBuffer;
                        if (Double.Parse(strChannelBuffer, CultureInfo.InvariantCulture) > 2.04)
                        {
                            MessageBox.Show("The PiKoder firmware version found is not supported! Please goto www.pikoder.com and upgrade PCC Control Center to the latest version.", "Error Message", MessageBoxButtons.OK);
                            Application.Exit();
                        }
                        else if (Double.Parse(strChannelBuffer, CultureInfo.InvariantCulture) > 2.02)
                        {
                            PPMModeLegacy = false;
                            bUART2PPM_StartUpValues = true;
                            ProtectedSaveMode = true;
                        }
                        else if (Double.Parse(strChannelBuffer, CultureInfo.InvariantCulture) > 2.00)
                        {
                            bUART2PPM_StartUpValues = true;
                            ProtectedSaveMode = true;
                        }
                        else if (Double.Parse(strChannelBuffer, CultureInfo.InvariantCulture) < 1.02)
                        {
                            MessageBox.Show("The PiKoder firmware version found is not supported! Please goto www.pikoder.com and upgrade PCC Control Center to the latest version.", "Error Message", MessageBoxButtons.OK);
                            Application.Exit();
                        }
                    }
                    else
                    {
                        boolErrorFlag = true;
                    }
                }

                GroupBox13.Enabled = true;  // setup form
                GroupBox13.Visible = true;
                GroupBox17.Enabled = true;
                GroupBox17.Visible = true;

                bDataLoaded = false;

                // RetrievePiKoderParameters()   ' left the retrieve part here locally because USB2PPM is not supported anymore

                if (!boolErrorFlag)
                {   // retrieve channel value
                    myPCAL.GetPulseLength(ref strChannelBuffer, 1, HPMath);
                    if (strChannelBuffer != "TimeOut")
                    {
                        ch1_HScrollBar.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                        strCH_1_Current.Text = Convert.ToString(ch1_HScrollBar.Value);
                    }
                    else
                    {
                        DialogResult dr = MessageBox.Show(
                                "Unable to retrieve valid channel 1 pulse length. Load factory default value?", "USB2PPM Parameter Error",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (dr == DialogResult.Yes)
                        {
                            if (myPCAL.SetChannelPulseLength(1, "1500"))
                            {
                                ch1_HScrollBar.Value = 1500;
                                strCH_1_Current.Text = "1500";
                            }
                            else
                            {
                                dr = MessageBox.Show(
                                        "Unable to correct channel 1 pulse length. For further assistance please contact support@pikoder.com.", "USB2PPM Parameter Error",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                                boolErrorFlag = true;
                                myMessage = "Unable to retrieve valid channel 1 pulse length";
                            }
                        } else
                        {
                            boolErrorFlag = true;
                            myMessage = "Unable to retrieve valid channel 1 pulse length";
                        }
                    }
                }

                if (!boolErrorFlag)
                {
                    myPCAL.GetPulseLength(ref strChannelBuffer, 2, HPMath);
                    if (strChannelBuffer != "TimeOut")
                    {
                        ch2_HScrollBar.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                        strCH_2_Current.Text = Convert.ToString(ch2_HScrollBar.Value);
                    }
                    else
                    {
                        DialogResult dr = MessageBox.Show(
                                "Unable to retrieve valid channel 2 pulse length. Load factory default value?", "USB2PPM Parameter Error",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (dr == DialogResult.Yes)
                        {
                            if (myPCAL.SetChannelPulseLength(2, "1500"))
                            {
                                ch2_HScrollBar.Value = 1500;
                                strCH_2_Current.Text = "1500";
                            }
                            else
                            {
                                dr = MessageBox.Show(
                                        "Unable to correct channel 2 pulse length. For further assistance please contact support@pikoder.com.", "USB2PPM Parameter Error",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                                boolErrorFlag = true;
                                myMessage = "Unable to retrieve valid channel 2 pulse length";
                            }
                        }
                        else
                        {
                            boolErrorFlag = true;
                            myMessage = "Unable to retrieve valid channel 2 pulse length";
                        }
                    }
                }

                if (!boolErrorFlag)
                {
                    myPCAL.GetPulseLength(ref strChannelBuffer, 3, HPMath);
                    if (strChannelBuffer != "TimeOut")
                    {
                        ch3_HScrollBar.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                        strCH_3_Current.Text = Convert.ToString(ch3_HScrollBar.Value);
                    }
                    else
                    {
                        DialogResult dr = MessageBox.Show(
                                "Unable to retrieve valid channel 3 pulse length. Load factory default value?", "USB2PPM Parameter Error",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (dr == DialogResult.Yes)
                        {
                            if (myPCAL.SetChannelPulseLength(3, "1500"))
                            {
                                ch3_HScrollBar.Value = 1500;
                                strCH_3_Current.Text = "1500";
                            }
                            else
                            {
                                dr = MessageBox.Show(
                                        "Unable to correct channel 3 pulse length. For further assistance please contact support@pikoder.com.", "USB2PPM Parameter Error",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                                boolErrorFlag = true;
                                myMessage = "Unable to retrieve valid channel 3 pulse length";
                            }
                        }
                        else
                        {
                            boolErrorFlag = true;
                            myMessage = "Unable to retrieve valid channel 3 pulse length";
                        }
                    }
                }

                if (!boolErrorFlag)
                {
                    myPCAL.GetPulseLength(ref strChannelBuffer, 4, HPMath);
                    if (strChannelBuffer != "TimeOut")
                    {
                        ch4_HScrollBar.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                        strCH_4_Current.Text = Convert.ToString(ch4_HScrollBar.Value);
                    }
                    else
                    {
                        DialogResult dr = MessageBox.Show(
                                "Unable to retrieve valid channel 4 pulse length. Load factory default value?", "USB2PPM Parameter Error",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (dr == DialogResult.Yes)
                        {
                            if (myPCAL.SetChannelPulseLength(4, "1500"))
                            {
                                ch4_HScrollBar.Value = 1500;
                                strCH_4_Current.Text = "1500";
                            }
                            else
                            {
                                dr = MessageBox.Show(
                                        "Unable to correct channel 4 pulse length. For further assistance please contact support@pikoder.com.", "USB2PPM Parameter Error",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                                boolErrorFlag = true;
                                myMessage = "Unable to retrieve valid channel 4 pulse length";
                            }
                        }
                        else
                        {
                            boolErrorFlag = true;
                            myMessage = "Unable to retrieve valid channel 4 pulse length";
                        }
                    }
                }

                if (!boolErrorFlag)
                {
                    myPCAL.GetPulseLength(ref strChannelBuffer, 5, HPMath);
                    if (strChannelBuffer != "TimeOut")
                    {
                        ch5_HScrollBar.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                        strCH_5_Current.Text = Convert.ToString(ch5_HScrollBar.Value);
                    }
                    else
                    {
                        DialogResult dr = MessageBox.Show(
                                "Unable to retrieve valid channel 5 pulse length. Load factory default value?", "USB2PPM Parameter Error",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (dr == DialogResult.Yes)
                        {
                            if (myPCAL.SetChannelPulseLength(5, "1500"))
                            {
                                ch5_HScrollBar.Value = 1500;
                                strCH_5_Current.Text = "1500";
                            }
                            else
                            {
                                dr = MessageBox.Show(
                                        "Unable to correct channel 5 pulse length. For further assistance please contact support@pikoder.com.", "USB2PPM Parameter Error",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                                boolErrorFlag = true;
                                myMessage = "Unable to retrieve valid channel 5 pulse length";
                            }
                        }
                        else
                        {
                            boolErrorFlag = true;
                            myMessage = "Unable to retrieve valid channel 5 pulse length";
                        }
                    }
                }

                if (!boolErrorFlag)
                {
                    myPCAL.GetPulseLength(ref strChannelBuffer, 6, HPMath);
                    if (strChannelBuffer != "TimeOut")
                    {
                        ch6_HScrollBar.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                        strCH_6_Current.Text = Convert.ToString(ch6_HScrollBar.Value);
                    }
                    else
                    {
                        DialogResult dr = MessageBox.Show(
                                "Unable to retrieve valid channel 6 pulse length. Load factory default value?", "USB2PPM Parameter Error",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (dr == DialogResult.Yes)
                        {
                            if (myPCAL.SetChannelPulseLength(6, "1500"))
                            {
                                ch6_HScrollBar.Value = 1500;
                                strCH_6_Current.Text = "1500";
                            }
                            else
                            {
                                dr = MessageBox.Show(
                                        "Unable to correct channel 6 pulse length. For further assistance please contact support@pikoder.com.", "USB2PPM Parameter Error",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                                boolErrorFlag = true;
                                myMessage = "Unable to retrieve valid channel 6 pulse length";
                            }
                        }
                        else
                        {
                            boolErrorFlag = true;
                            myMessage = "Unable to retrieve valid channel 6 pulse length";
                        }
                    }
                }

                if (!boolErrorFlag)
                {
                    myPCAL.GetPulseLength(ref strChannelBuffer, 7, HPMath);
                    if (strChannelBuffer != "TimeOut")
                    {
                        ch7_HScrollBar.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                        strCH_7_Current.Text = Convert.ToString(ch7_HScrollBar.Value);
                    }
                    else
                    {
                        DialogResult dr = MessageBox.Show(
                                "Unable to retrieve valid channel 7 pulse length. Load factory default value?", "USB2PPM Parameter Error",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (dr == DialogResult.Yes)
                        {
                            if (myPCAL.SetChannelPulseLength(7, "1500"))
                            {
                                ch7_HScrollBar.Value = 1500;
                                strCH_7_Current.Text = "1500";
                            }
                            else
                            {
                                dr = MessageBox.Show(
                                        "Unable to correct channel 7 pulse length. For further assistance please contact support@pikoder.com.", "USB2PPM Parameter Error",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                                boolErrorFlag = true;
                                myMessage = "Unable to retrieve valid channel 7 pulse length";
                            }
                        }
                        else
                        {
                            boolErrorFlag = true;
                            myMessage = "Unable to retrieve valid channel 7 pulse length";
                        }
                    }
                }

                if (!boolErrorFlag)
                {
                    myPCAL.GetPulseLength(ref strChannelBuffer, 8, HPMath);
                    if (strChannelBuffer != "TimeOut")
                    {
                        ch8_HScrollBar.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                        strCH_8_Current.Text = Convert.ToString(ch8_HScrollBar.Value);
                    }
                    else
                    {
                        DialogResult dr = MessageBox.Show(
                                "Unable to retrieve valid channel 8 pulse length. Load factory default value?", "USB2PPM Parameter Error",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (dr == DialogResult.Yes)
                        {
                            if (myPCAL.SetChannelPulseLength(8, "1500"))
                            {
                                ch8_HScrollBar.Value = 1500;
                                strCH_8_Current.Text = "1500";
                            }
                            else
                            {
                                dr = MessageBox.Show(
                                        "Unable to correct channel 8 pulse length. For further assistance please contact support@pikoder.com.", "USB2PPM Parameter Error",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                                boolErrorFlag = true;
                                myMessage = "Unable to retrieve valid channel 8 pulse length";
                            }
                        }
                        else
                        {
                            boolErrorFlag = true;
                            myMessage = "Unable to retrieve valid channel 8 pulse length";
                        }
                    }
                }

                // set min & max information for all channels
                strCH_1_Min.Value = iDefaultMinValue;
                ch1_HScrollBar.Minimum = iDefaultMinValue;
                strCH_1_Min.ForeColor = Color.LightGray;
                strCH_1_Min.Enabled = false;
                strCH_1_Max.Value = iDefaultMaxValue;
                ch1_HScrollBar.Maximum = iDefaultMaxValue;
                strCH_1_Max.ForeColor = Color.LightGray;
                strCH_1_Max.Enabled = false;

                strCH_2_Min.Value = iDefaultMinValue;
                ch2_HScrollBar.Minimum = iDefaultMinValue;
                strCH_2_Min.ForeColor = Color.LightGray;
                strCH_2_Min.Enabled = false;
                strCH_2_Max.Value = iDefaultMaxValue;
                ch2_HScrollBar.Maximum = iDefaultMaxValue;
                strCH_2_Max.ForeColor = Color.LightGray;
                strCH_2_Max.Enabled = false;

                strCH_3_Min.Value = iDefaultMinValue;
                ch3_HScrollBar.Minimum = iDefaultMinValue;
                strCH_3_Min.ForeColor = Color.LightGray;
                strCH_3_Min.Enabled = false;
                strCH_3_Max.Value = iDefaultMaxValue;
                ch3_HScrollBar.Maximum = iDefaultMaxValue;
                strCH_3_Max.ForeColor = Color.LightGray;
                strCH_3_Max.Enabled = false;

                strCH_4_Min.Value = iDefaultMinValue;
                ch4_HScrollBar.Minimum = iDefaultMinValue;
                strCH_4_Min.ForeColor = Color.LightGray;
                strCH_4_Min.Enabled = false;
                strCH_4_Max.Value = iDefaultMaxValue;
                ch4_HScrollBar.Maximum = iDefaultMaxValue;
                strCH_4_Max.ForeColor = Color.LightGray;
                strCH_4_Max.Enabled = false;

                strCH_5_Min.Value = iDefaultMinValue;
                ch5_HScrollBar.Minimum = iDefaultMinValue;
                strCH_5_Min.ForeColor = Color.LightGray;
                strCH_5_Min.Enabled = false;
                strCH_5_Max.Value = iDefaultMaxValue;
                ch5_HScrollBar.Maximum = iDefaultMaxValue;
                strCH_5_Max.ForeColor = Color.LightGray;
                strCH_5_Max.Enabled = false;

                strCH_6_Min.Value = iDefaultMinValue;
                ch6_HScrollBar.Minimum = iDefaultMinValue;
                strCH_6_Min.ForeColor = Color.LightGray;
                strCH_6_Min.Enabled = false;
                strCH_6_Max.Value = iDefaultMaxValue;
                ch6_HScrollBar.Maximum = iDefaultMaxValue;
                strCH_6_Max.ForeColor = Color.LightGray;
                strCH_6_Max.Enabled = false;

                strCH_7_Min.Value = iDefaultMinValue;
                ch7_HScrollBar.Minimum = iDefaultMinValue;
                strCH_7_Min.ForeColor = Color.LightGray;
                strCH_7_Min.Enabled = false;
                strCH_7_Max.Value = iDefaultMaxValue;
                ch7_HScrollBar.Maximum = iDefaultMaxValue;
                strCH_7_Max.ForeColor = Color.LightGray;
                strCH_7_Max.Enabled = false;

                strCH_8_Min.Value = iDefaultMinValue;
                ch8_HScrollBar.Minimum = iDefaultMinValue;
                strCH_8_Min.ForeColor = Color.LightGray;
                strCH_8_Min.Enabled = false;
                strCH_8_Max.Value = iDefaultMaxValue;
                ch8_HScrollBar.Maximum = iDefaultMaxValue;
                strCH_8_Max.ForeColor = Color.LightGray;
                strCH_8_Max.Enabled = false;


                if (!bUART2PPM_StartUpValues)
                {
                    GroupBox11.Enabled = false;     // neutral positions
                    GroupBox11.Visible = false;
                    GroupBox7.Enabled = false;      // Save Parameters
                    GroupBox7.Visible = false;
                    GroupBox8.Enabled = false;      // miniSSC Offset
                }
                else
                {
                    GroupBox11.Text = "startup";     // neutral positions
                    if (!boolErrorFlag)
                    {   // retrieve neutral value
                        myPCAL.GetNeutralPosition(ref strChannelBuffer, 1, HPMath);
                        if (strChannelBuffer != "TimeOut")
                        {
                            strCH_1_Neutral.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                            strCH_1_Neutral.ForeColor = Color.Black;
                        }
                        else
                        {
                            DialogResult dr = MessageBox.Show(
                                    "Unable to retrieve valid channel 1 startup value. Load factory default value?", "USB2PPM Parameter Error",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                            if (dr == DialogResult.Yes)
                            {
                                if (myPCAL.SetChannelNeutral("1500", 1))
                                {
                                    strCH_1_Neutral.Value = 1500;
                                    strCH_1_Neutral.ForeColor = Color.Black;
                                }
                                else
                                {
                                    dr = MessageBox.Show(
                                            "Unable to correct channel 1 startup value. For further assistance please contact support@pikoder.com.", "USB2PPM Parameter Error",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    boolErrorFlag = true;
                                    myMessage = "Unable to retrieve valid channel 1 startup value";
                                }
                            }
                            else
                            {
                                boolErrorFlag = true;
                                myMessage = "Unable to retrieve valid channel 1 startup value";
                            }
                        }
                    }

                    if (!boolErrorFlag)
                    {   // retrieve neutral value
                        myPCAL.GetNeutralPosition(ref strChannelBuffer, 2, HPMath);
                        if (strChannelBuffer != "TimeOut")
                        {
                            strCH_2_Neutral.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                            strCH_2_Neutral.ForeColor = Color.Black;
                        }
                        else
                        {
                            DialogResult dr = MessageBox.Show(
                                    "Unable to retrieve valid channel 2 startup value. Load factory default value?", "USB2PPM Parameter Error",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                            if (dr == DialogResult.Yes)
                            {
                                if (myPCAL.SetChannelNeutral("1500", 2))
                                {
                                    strCH_2_Neutral.Value = 1500;
                                    strCH_2_Neutral.ForeColor = Color.Black;
                                }
                                else
                                {
                                    dr = MessageBox.Show(
                                            "Unable to correct channel 2 startup value. For further assistance please contact support@pikoder.com.", "USB2PPM Parameter Error",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    boolErrorFlag = true;
                                    myMessage = "Unable to retrieve valid channel 2 startup value";
                                }
                            }
                            else
                            {
                                boolErrorFlag = true;
                                myMessage = "Unable to retrieve valid channel 2 startup value";
                            }
                        }
                    }

                    if (!boolErrorFlag)
                    {   // retrieve neutral value
                        myPCAL.GetNeutralPosition(ref strChannelBuffer, 3, HPMath);
                        if (strChannelBuffer != "TimeOut")
                        {
                            strCH_3_Neutral.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                            strCH_3_Neutral.ForeColor = Color.Black;
                        }
                        else
                        {
                            DialogResult dr = MessageBox.Show(
                                    "Unable to retrieve valid channel 3 startup value. Load factory default value?", "USB2PPM Parameter Error",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                            if (dr == DialogResult.Yes)
                            {
                                if (myPCAL.SetChannelNeutral("1500", 3))
                                {
                                    strCH_3_Neutral.Value = 1500;
                                    strCH_3_Neutral.ForeColor = Color.Black;
                                }
                                else
                                {
                                    dr = MessageBox.Show(
                                            "Unable to correct channel 3 startup value. For further assistance please contact support@pikoder.com.", "USB2PPM Parameter Error",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    boolErrorFlag = true;
                                    myMessage = "Unable to retrieve valid channel 3 startup value";
                                }
                            }
                            else
                            {
                                boolErrorFlag = true;
                                myMessage = "Unable to retrieve valid channel 3 startup value";
                            }
                        }
                    }

                    if (!boolErrorFlag)
                    {   // retrieve neutral value
                        myPCAL.GetNeutralPosition(ref strChannelBuffer, 4, HPMath);
                        if (strChannelBuffer != "TimeOut")
                        {
                            strCH_4_Neutral.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                            strCH_4_Neutral.ForeColor = Color.Black;
                        }
                        else
                        {
                            DialogResult dr = MessageBox.Show(
                                    "Unable to retrieve valid channel 4 startup value. Load factory default value?", "USB2PPM Parameter Error",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                            if (dr == DialogResult.Yes)
                            {
                                if (myPCAL.SetChannelNeutral("1500", 4))
                                {
                                    strCH_4_Neutral.Value = 1500;
                                    strCH_4_Neutral.ForeColor = Color.Black;
                                }
                                else
                                {
                                    dr = MessageBox.Show(
                                            "Unable to correct channel 4 startup value. For further assistance please contact support@pikoder.com.", "USB2PPM Parameter Error",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    boolErrorFlag = true;
                                    myMessage = "Unable to retrieve valid channel 4 startup value";
                                }
                            }
                            else
                            {
                                boolErrorFlag = true;
                                myMessage = "Unable to retrieve valid channel 4 startup value";
                            }
                        }
                    }

                    if (!boolErrorFlag)
                    {   // retrieve neutral value
                        myPCAL.GetNeutralPosition(ref strChannelBuffer, 5, HPMath);
                        if (strChannelBuffer != "TimeOut")
                        {
                            strCH_5_Neutral.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                            strCH_5_Neutral.ForeColor = Color.Black;
                        }
                        else
                        {
                            DialogResult dr = MessageBox.Show(
                                    "Unable to retrieve valid channel 5 startup value. Load factory default value?", "USB2PPM Parameter Error",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                            if (dr == DialogResult.Yes)
                            {
                                if (myPCAL.SetChannelNeutral("1500", 5))
                                {
                                    strCH_5_Neutral.Value = 1500;
                                    strCH_5_Neutral.ForeColor = Color.Black;
                                }
                                else
                                {
                                    dr = MessageBox.Show(
                                            "Unable to correct channel 5 startup value. For further assistance please contact support@pikoder.com.", "USB2PPM Parameter Error",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    boolErrorFlag = true;
                                    myMessage = "Unable to retrieve valid channel 5 startup value";
                                }
                            }
                            else
                            {
                                boolErrorFlag = true;
                                myMessage = "Unable to retrieve valid channel 5 startup value";
                            }
                        }
                    }

                    if (!boolErrorFlag)
                    {   // retrieve neutral value
                        myPCAL.GetNeutralPosition(ref strChannelBuffer, 6, HPMath);
                        if (strChannelBuffer != "TimeOut")
                        {
                            strCH_6_Neutral.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                            strCH_6_Neutral.ForeColor = Color.Black;
                        }
                        else
                        {
                            DialogResult dr = MessageBox.Show(
                                    "Unable to retrieve valid channel 6 startup value. Load factory default value?", "USB2PPM Parameter Error",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                            if (dr == DialogResult.Yes)
                            {
                                if (myPCAL.SetChannelNeutral("1500", 6))
                                {
                                    strCH_6_Neutral.Value = 1500;
                                    strCH_6_Neutral.ForeColor = Color.Black;
                                }
                                else
                                {
                                    dr = MessageBox.Show(
                                            "Unable to correct channel 6 startup value. For further assistance please contact support@pikoder.com.", "USB2PPM Parameter Error",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    boolErrorFlag = true;
                                    myMessage = "Unable to retrieve valid channel 6 startup value";
                                }
                            }
                            else
                            {
                                boolErrorFlag = true;
                                myMessage = "Unable to retrieve valid channel 6 startup value";
                            }
                        }
                    }

                    if (!boolErrorFlag)
                    {   // retrieve neutral value
                        myPCAL.GetNeutralPosition(ref strChannelBuffer, 7, HPMath);
                        if (strChannelBuffer != "TimeOut")
                        {
                            strCH_7_Neutral.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                            strCH_7_Neutral.ForeColor = Color.Black;
                        }
                        else
                        {
                            DialogResult dr = MessageBox.Show(
                                    "Unable to retrieve valid channel 7 startup value. Load factory default value?", "USB2PPM Parameter Error",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                            if (dr == DialogResult.Yes)
                            {
                                if (myPCAL.SetChannelNeutral("1500", 7))
                                {
                                    strCH_7_Neutral.Value = 1500;
                                    strCH_7_Neutral.ForeColor = Color.Black;
                                }
                                else
                                {
                                    dr = MessageBox.Show(
                                            "Unable to correct channel 7 startup value. For further assistance please contact support@pikoder.com.", "USB2PPM Parameter Error",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    boolErrorFlag = true;
                                    myMessage = "Unable to retrieve valid channel 7 startup value";
                                }
                            }
                            else
                            {
                                boolErrorFlag = true;
                                myMessage = "Unable to retrieve valid channel 7 startup value";
                            }
                        }
                    }

                    if (!boolErrorFlag)
                    {   // retrieve neutral value
                        myPCAL.GetNeutralPosition(ref strChannelBuffer, 8, HPMath);
                        if (strChannelBuffer != "TimeOut")
                        {
                            strCH_8_Neutral.Value = ScalePulseWidth(strChannelBuffer, HPMath);
                            strCH_8_Neutral.ForeColor = Color.Black;
                        }
                        else
                        {
                            DialogResult dr = MessageBox.Show(
                                    "Unable to retrieve valid channel 8 startup value. Load factory default value?", "USB2PPM Parameter Error",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                            if (dr == DialogResult.Yes)
                            {
                                if (myPCAL.SetChannelNeutral("1500", 8))
                                {
                                    strCH_8_Neutral.Value = 1500;
                                    strCH_8_Neutral.ForeColor = Color.Black;
                                }
                                else
                                {
                                    dr = MessageBox.Show(
                                            "Unable to correct channel 8 startup value. For further assistance please contact support@pikoder.com.", "USB2PPM Parameter Error",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    boolErrorFlag = true;
                                    myMessage = "Unable to retrieve valid channel 8 startup value";
                                }
                            }
                            else
                            {
                                boolErrorFlag = true;
                                myMessage = "Unable to retrieve valid channel 8 startup value";
                            }
                        }
                    }
                
                }

                GroupBox8.Visible = false;
                GroupBox4.Enabled = false;       // zero offset
                GroupBox4.Visible = false;

                PPM_Mode.Items.Add("positive");
                PPM_Mode.Items.Add("negative (Futaba)");

                if (!boolErrorFlag)
                {   // retrieve PPM settings
                    if (PPMModeLegacy)  // retrieve not possible -> push default
                    {
                        if (myPCAL.SetPiKoderPPMChannels(8))
                        {
                            if (myPCAL.SetPiKoderPPMMode(1))
                            {
                                if (myPCAL.SetPiKoderPPMMode(1))
                                {
                                    PPM_Mode.SelectedIndex = 1;
                                    PPMmode = 1;
                                }
                                else
                                {
                                    boolErrorFlag = true;
                                    myMessage = "Unable to set PPM values";
                                }
                            }
                            else
                            {
                                boolErrorFlag = true;
                                myMessage = "Unable to set PPM values";
                            }
                        }
                        else
                        {
                            boolErrorFlag = true;
                            myMessage = "Unable to set PPM values";
                        }
                    }
                    else
                    {
                        myPCAL.GetPPMSettings(ref strChannelBuffer);
                        if (strChannelBuffer != "TimeOut")
                        {
                            PPM_Channels.Value = strChannelBuffer[0] - '0';
                            if (strChannelBuffer[1] == 'P')
                            {
                                PPM_Mode.SelectedIndex = 0;
                                PPMmode = 0;
                            }
                            else
                            {
                                PPM_Mode.SelectedIndex = 1;
                                PPMmode = 1;
                            }

                        }
                        else
                        {
                            DialogResult dr = MessageBox.Show(
                                    "Unable to retrieve valid PPM settings. Load factory default value?", "USB2PPM Parameter Error",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                            if (dr == DialogResult.Yes)
                            {
                                if (myPCAL.SetPiKoderPPMSettings(8,1))
                                {
                                    PPM_Channels.Value = 8;
                                    PPM_Mode.SelectedIndex = 1;
                                    PPMmode = 1;
                                }
                                else
                                {
                                    dr = MessageBox.Show(
                                            "Unable to correct PPM settings. For further assistance please contact support@pikoder.com.", "USB2PPM Parameter Error",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    boolErrorFlag = true;
                                    myMessage = "Unable to retrieve valid PPM settings";
                                }
                            }
                            else
                            {
                                boolErrorFlag = true;
                                myMessage = "Unable to retrieve valid PPM settings";
                            }
                        }
                    }

                }
            }

            PPM_Channels.ForeColor = Color.Black;
            PPM_Mode.ForeColor = Color.Black;
            PPM_Mode.ClearSelected();
            bDataLoaded = true;
            IndicateConnectionOk();
        }

        private void PCCpro_FormClosing(Object sender, FormClosingEventArgs e)
        {
        }

    }
}

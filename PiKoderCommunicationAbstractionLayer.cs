﻿// This class is designed as an abstraction layer to serve the different hardware communication links to the PiKoder 
// such as COM or WLAN. All PiKoder parameters or commands are executed transparently using respective methods. 
//
// The class maintains information regarding the protocol when involved by EstablishLink(). From this point onwards an 
// existing connection would be monitored and an event would be generated once the connection is lost (LinkLost).
// 
// Copyright 2020 Gregor Schlechtriem
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
//

using System;
using System.Windows.Forms;

public class PiKoderCommunicationAbstractionLayer
{
    public enum iPhysicalLink

    {
        iSerialLink,
        iWLANlink
    }

    private SerialLink mySerialLink = new SerialLink();
    private WLANLink myWLANLink = new WLANLink();
    private bool Connected = false;     // connection status
    private iPhysicalLink iConnectedTo;

    public bool LinkConnected()
    {
        return Connected;
    }

    public bool EstablishLink(string SelectedPort, iPhysicalLink ConnectionType)
    {
        Connected = false;
        if (ConnectionType == iPhysicalLink.iSerialLink)
        {
            Connected = mySerialLink.EstablishSerialLink(SelectedPort);
            if (Connected)
            {
                iConnectedTo = iPhysicalLink.iSerialLink;
            }
        }
        else if (ConnectionType == iPhysicalLink.iWLANlink)
        {
            Connected = myWLANLink.EstablishWLANLink();
            if (Connected)
            {
                iConnectedTo = iPhysicalLink.iWLANlink;
            }
        }
        return Connected;
    }

    public void DisconnectLink(iPhysicalLink ConnectionType)
    {
        if (Connected & (ConnectionType == iPhysicalLink.iSerialLink))
        {
            mySerialLink.MyForm_Dispose();
        }
        else if (Connected & (ConnectionType == iPhysicalLink.iWLANlink))
        {
            myWLANLink.MyForm_Dispose();
        }
        MyForm_Dispose();
    }

    public void GetPulseLengthFT(ref string SerialInputString, int iChannelNo, bool HPMath, bool ProtectedSaveMode)
    {
        int timeOut = 0;
        do
        {
            if (iConnectedTo == iPhysicalLink.iSerialLink)
            {
                mySerialLink.SendDataToSerial(iChannelNo.ToString() + "?");
                SerialInputString = mySerialLink.SerialReceiver();
            }
            else if (iConnectedTo == iPhysicalLink.iWLANlink)
            {
                myWLANLink.SendDataToWLAN(iChannelNo.ToString() + "?");
                SerialInputString = myWLANLink.Receiver();
            }
            timeOut += 1;
        } while (!ValidatePulseValue(ref SerialInputString, HPMath) & timeOut < 5);
        if (timeOut == 5)
        {
            if (SerialInputString.Length == 4)
            {
                // read numbers, but value not within boundries
                DialogResult dr = MessageBox.Show("Pulse width for channel " + iChannelNo.ToString() +
                    " outside of range (" + SerialInputString + "). Reset to default?", "Inconsistent channel value",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                if (dr == DialogResult.Cancel)
                {
                    Environment.Exit(1);
                }
                SerialInputString = "1500";
                SetChannelPulseLength(iChannelNo, "1500");
                SetPiKoderPreferences(ProtectedSaveMode);
                return;
            }
            SerialInputString = "TimeOut";
        }
    }

    public void GetNeutralPositionFT(ref string SerialInputString, int iChannelNo, bool HPMath, bool ProtectedSaveMode)
    {
        int timeOut = 0;
        do
        {
            if (iConnectedTo == iPhysicalLink.iSerialLink)
            {
                mySerialLink.SendDataToSerial("N" + iChannelNo.ToString() + "?");
                SerialInputString = mySerialLink.SerialReceiver();
            }
            else if (iConnectedTo == iPhysicalLink.iWLANlink)
            {
                myWLANLink.SendDataToWLAN("N" + iChannelNo.ToString() + "?");
                SerialInputString = myWLANLink.Receiver();
            }
            timeOut += 1;
        } while (!ValidatePulseValue(ref SerialInputString, HPMath) & timeOut < 5);
        if (timeOut == 5)
        {
            if (SerialInputString.Length == 4)
            {
                // read numbers, but value not within boundries
                DialogResult dr = MessageBox.Show("Neutral setting for channel " + iChannelNo.ToString() +
                    " outside of range (" + SerialInputString + "). Reset to default?", "Inconsistent channel value",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                if (dr == DialogResult.Cancel)
                {
                    Environment.Exit(1);
                }
                SerialInputString = "1500";
                SetChannelNeutral("1500", iChannelNo);
                SetPiKoderPreferences(ProtectedSaveMode);
                return;
            }
            SerialInputString = "TimeOut";
        }
    }

    public void GetLowerLimitFT(ref string SerialInputString, int iChannelNo, bool HPMath, bool ProtectedSaveMode)
    {
        int timeOut = 1;
        do
        {
            if (iConnectedTo == iPhysicalLink.iSerialLink)
            {
                mySerialLink.SendDataToSerial("L" + iChannelNo.ToString() + "?");
                SerialInputString = mySerialLink.SerialReceiver();
            }
            else if (iConnectedTo == iPhysicalLink.iWLANlink)
            {
                myWLANLink.SendDataToWLAN("L" + iChannelNo.ToString() + "?");
                SerialInputString = myWLANLink.Receiver();
            }
            timeOut += 1;
        } while (!ValidatePulseValue(ref SerialInputString, HPMath) & timeOut < 5);
        if (timeOut == 5)
        {
            if (SerialInputString.Length == 4)
            {
                // read numbers, but value not within boundries
                DialogResult dr = MessageBox.Show("Lower limit for channel " + iChannelNo.ToString() +
                    " outside of range (" + SerialInputString + "). Reset to default?", "Inconsistent channel value",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                if (dr == DialogResult.Cancel)
                {
                    Environment.Exit(1);
                }
                SerialInputString = "1000";
                SetChannelLowerLimit("1000", iChannelNo);
                SetPiKoderPreferences(ProtectedSaveMode);
                return;
            }
            SerialInputString = "TimeOut";
        }
    }

    public void GetUpperLimitFT(ref string SerialInputString, int iChannelNo, bool HPMath, bool ProtectedSaveMode)
    {
        int timeOut = 0;
        do
        {
            if (iConnectedTo == iPhysicalLink.iSerialLink)
            {
                mySerialLink.SendDataToSerial("U" + iChannelNo.ToString() + "?");
                SerialInputString = mySerialLink.SerialReceiver();
            }
            else if (iConnectedTo == iPhysicalLink.iWLANlink)
            {
                myWLANLink.SendDataToWLAN("U" + iChannelNo.ToString() + "?");
                SerialInputString = myWLANLink.Receiver();
            }
            timeOut += 1;
        } while (!ValidatePulseValue(ref SerialInputString, HPMath) & timeOut < 5);
        if (timeOut == 5)
        {
            if (SerialInputString.Length == 4)
            {
                // read numbers, but value not within boundries
                DialogResult dr = MessageBox.Show("Upper limit for channel " + iChannelNo.ToString() +
                    " outside of range (" + SerialInputString + "). Reset to default?", "Inconsistent channel value",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                if (dr == DialogResult.Cancel)
                {
                    Environment.Exit(1);
                }
                SerialInputString = "2000";
                SetChannelUpperLimit("2000", iChannelNo);
                SetPiKoderPreferences(ProtectedSaveMode);
                return;
            }
            SerialInputString = "TimeOut";
        }
    }

    public void GetIOType(ref string SerialInputString, int iChannelNo)
    {
        if (iConnectedTo == iPhysicalLink.iSerialLink)
        {
            mySerialLink.SendDataToSerial("O" + iChannelNo.ToString() + "?");
            SerialInputString = mySerialLink.SerialReceiver();
        }
        else if (iConnectedTo == iPhysicalLink.iWLANlink)
        {
            myWLANLink.SendDataToWLAN("O" + iChannelNo.ToString() + "?");
            SerialInputString = myWLANLink.Receiver();
        }
    }

    public void GetFirmwareVersion(ref string SerialInputString)
    {
        int iTimeOut = 0;
        do
        {
            if (iConnectedTo == iPhysicalLink.iSerialLink)
            {
                mySerialLink.SendDataToSerial("0");
                SerialInputString = mySerialLink.SerialReceiver();
            }
            else if (iConnectedTo == iPhysicalLink.iWLANlink)
            {
                myWLANLink.SendDataToWLAN("0");
                SerialInputString = myWLANLink.Receiver();
            }
            iTimeOut += 1;
        } while (SerialInputString == "TimeOut" & iTimeOut < 5);
    }

    public void GetESPFirmwareVersion(ref string SerialInputString)
    {
        int iTimeOut = 0;
        do
        {
            if (iConnectedTo == iPhysicalLink.iSerialLink)
            {
            }
            else if (iConnectedTo == iPhysicalLink.iWLANlink)
            {
                myWLANLink.SendDataToWLAN("$");
                SerialInputString = myWLANLink.Receiver();
            }
            iTimeOut += 1;
        } while (SerialInputString == "TimeOut" & iTimeOut < 5);
    }

    public void GetStatusRecord(ref string SerialInputString)
    {
        int iTimeOut = 0;
        do
        {
            if (iConnectedTo == iPhysicalLink.iSerialLink)
            {
                while (mySerialLink.SerialReceiver(1) != "TimeOut")
                {
                }
                mySerialLink.SendDataToSerial("?");
                if (mySerialLink.SerialLinkConnected())
                {
                    SerialInputString = mySerialLink.SerialReceiver(3);
                }
                else
                {
                    SerialInputString = "TimeOut";
                    return;
                }
            }
            else if (iConnectedTo == iPhysicalLink.iWLANlink)
            {
                myWLANLink.SendDataToWLAN("?");
                SerialInputString = myWLANLink.Receiver();
            }
            iTimeOut++;
        } while (!ValidateTypeRecord(SerialInputString) & (iTimeOut < 5));
        if (iTimeOut == 5)
        {
            SerialInputString = "TimeOut";
        }
    }


    public void GetTimeOut(ref string SerialInputString)
    {
        int iTimeOut = 0;
        SerialInputString = "";
        do
        {
            if (iConnectedTo == iPhysicalLink.iSerialLink)
            {
                mySerialLink.SendDataToSerial("T?");
                SerialInputString = mySerialLink.SerialReceiver();
            }
            else if (iConnectedTo == iPhysicalLink.iWLANlink)
            {
                myWLANLink.SendDataToWLAN("T?");
                SerialInputString = myWLANLink.Receiver();
            }
            iTimeOut += 1;
        } while (!ValidateTimeOut(SerialInputString) & iTimeOut < 5);
        if (iTimeOut == 5)
        {
            SerialInputString = "TimeOut";
        }
    }

    public void GetMiniSSCOffset(ref string SerialInputString)
    {
        int iTimeOut = 0;
        SerialInputString = "";
        do
        {
            if (iConnectedTo == iPhysicalLink.iSerialLink)
            {
                mySerialLink.SendDataToSerial("M?");
                SerialInputString = mySerialLink.SerialReceiver();
            }
            else if (iConnectedTo == iPhysicalLink.iWLANlink)
            {
                myWLANLink.SendDataToWLAN("M?");
                SerialInputString = myWLANLink.Receiver();
            }
            iTimeOut += 1;
        } while (!ValidateZeroOffset(SerialInputString) & iTimeOut < 5);
        if (iTimeOut == 5)
        {
            SerialInputString = "TimeOut";
        }
    }

    public void GetISCBaseAddress(ref string SerialInputString)
    {
        int iTimeOut = 0;
        SerialInputString = "";
        do
        {
            if (iConnectedTo == iPhysicalLink.iSerialLink)
            {
                mySerialLink.SendDataToSerial("I?");
                SerialInputString = mySerialLink.SerialReceiver();
            }
            else if (iConnectedTo == iPhysicalLink.iWLANlink)
            {
                myWLANLink.SendDataToWLAN("I?");
                SerialInputString = myWLANLink.Receiver();
            }
            iTimeOut += 1;
        } while (!ValidateI2CBaseAddress(SerialInputString) & iTimeOut < 5);
        if (iTimeOut == 5)
        {
            SerialInputString = "TimeOut";
        }
    }

    public void GetPPMSettings(ref string SerialInputString)
    {
        int iTimeOut = 0;
        SerialInputString = "";
        do
        {
            if (iConnectedTo == iPhysicalLink.iSerialLink)
            {
                mySerialLink.SendDataToSerial("P?");
                SerialInputString = mySerialLink.SerialReceiver();
            }
            else if (iConnectedTo == iPhysicalLink.iWLANlink)
            {
                myWLANLink.SendDataToWLAN("P?");
                SerialInputString = myWLANLink.Receiver();
            }
            iTimeOut += 1;
        } while (!ValidatePPMSettings(SerialInputString) & iTimeOut < 5);
        if (iTimeOut == 5)
        {
            SerialInputString = "TimeOut";
        }
    }

    public bool SetChannelNeutral(string strNeutralVal, int iChannelNo)
    {
        string ReturnCode = "";
        if (iConnectedTo == iPhysicalLink.iSerialLink)
        {
            mySerialLink.SendDataToSerial("N" + iChannelNo.ToString() + "=" + strNeutralVal);
            ReturnCode = mySerialLink.SerialReceiver();
        }
        else if (iConnectedTo == iPhysicalLink.iWLANlink)
        {
            myWLANLink.SendDataToWLAN("N" + iChannelNo.ToString() + "=" + strNeutralVal);
            ReturnCode = myWLANLink.Receiver();
        }
        return InterpretReturnCode(ReturnCode);
    }

    public bool SetChannelLowerLimit(string strNeutralVal, int iChannelNo)
    {
        string ReturnCode = "";
        if (iConnectedTo == iPhysicalLink.iSerialLink)
        {
            mySerialLink.SendDataToSerial("L" + iChannelNo.ToString() + "=" + strNeutralVal);
            ReturnCode = mySerialLink.SerialReceiver();
        }
        else if (iConnectedTo == iPhysicalLink.iWLANlink)
        {
            myWLANLink.SendDataToWLAN("L" + iChannelNo.ToString() + "=" + strNeutralVal);
            ReturnCode = myWLANLink.Receiver();
        }
        return InterpretReturnCode(ReturnCode);
    }

    public bool SetChannelUpperLimit(string strNeutralVal, int iChannelNo)
    {
        string ReturnCode = "";
        if (iConnectedTo == iPhysicalLink.iSerialLink)
        {
            ReturnCode = mySerialLink.SendDataToSerialwithAck("U" + iChannelNo.ToString() + "=" + strNeutralVal);
        }
        else if (iConnectedTo == iPhysicalLink.iWLANlink)
        {
            myWLANLink.SendDataToWLAN("U" + iChannelNo.ToString() + "=" + strNeutralVal);
            ReturnCode = myWLANLink.Receiver();
        }
        return InterpretReturnCode(ReturnCode);
    }

    public bool SetChannelOutputType(int iChannelNo, string strOutputType)
    {
        string ReturnCode = "";
        if (iConnectedTo == iPhysicalLink.iSerialLink)
        {
            mySerialLink.SendDataToSerial("O" + iChannelNo.ToString() + "=" + strOutputType);
            ReturnCode = mySerialLink.SerialReceiver();
        }
        else if (iConnectedTo == iPhysicalLink.iWLANlink)
        {
            myWLANLink.SendDataToWLAN("O" + iChannelNo.ToString() + "=" + strOutputType);
            ReturnCode = myWLANLink.Receiver();
        }
        return InterpretReturnCode(ReturnCode);
    }

    public bool SetChannelPulseLength(int iChannelNo, string strPulseLength)
    {
        string strSendString;
        string ReturnCode = "";
        strSendString = iChannelNo.ToString() + "=";
        if (strPulseLength.Length == 3)
        {
            strSendString += "0";
        }
        strSendString += strPulseLength;
        if (iConnectedTo == iPhysicalLink.iSerialLink)
        {
            mySerialLink.SendDataToSerial(strSendString);
            ReturnCode = mySerialLink.SerialReceiver();
        }
        else if (iConnectedTo == iPhysicalLink.iWLANlink)
        {
            myWLANLink.SendDataToWLAN(strSendString);
            ReturnCode = myWLANLink.Receiver();
        }
        return InterpretReturnCode(ReturnCode);
    }

    public bool SetHPChannelPulseLength(int iChannelNo, string strPulseLength)
    {
        string strSendString;
        string ReturnCode = "";
        strSendString = iChannelNo.ToString() + "=";
        if (strPulseLength.Length == 4)
        {
            strSendString += "0";
        }
        strSendString += strPulseLength;
        if (iConnectedTo == iPhysicalLink.iSerialLink)
        {
            mySerialLink.SendDataToSerial(strSendString);
            ReturnCode = mySerialLink.SerialReceiver();
        }
        else if (iConnectedTo == iPhysicalLink.iWLANlink)
        {
            myWLANLink.SendDataToWLAN(strSendString);
            ReturnCode = myWLANLink.Receiver();
        }
        return InterpretReturnCode(ReturnCode);
    }

    public bool SetPiKoderTimeOut(string strTimeOut)
    {
        string myString = "";
        if (strTimeOut.Length == 1)
        {
            myString = "00" + myString;
        }
        if (strTimeOut.Length == 2)
        {
            myString = "0" + myString;
        }
        if (iConnectedTo == iPhysicalLink.iSerialLink)
        {
            return InterpretReturnCode(mySerialLink.SendDataToSerialwithAck("T=" + myString + strTimeOut));
        }
        else if (iConnectedTo == iPhysicalLink.iWLANlink)
        {
            myWLANLink.SendDataToWLAN("T=" + myString + strTimeOut);
        }
        return InterpretReturnCode(myWLANLink.Receiver());
    }

    public bool SetPiKoderMiniSSCOffset(string strMiniSSCOffset)
    {
        string strSendString = "";
        if (strMiniSSCOffset.Length == 1)
        {
            strSendString = "00" + strSendString;
        }
        if (strMiniSSCOffset.Length == 2)
        {
            strSendString = "0" + strSendString;
        }
        if (iConnectedTo == iPhysicalLink.iSerialLink)
        {
            return InterpretReturnCode(mySerialLink.SendDataToSerialwithAck("M=" + strSendString + strMiniSSCOffset));
        }
        else 
        {
            myWLANLink.SendDataToWLAN("M=" + strSendString + strMiniSSCOffset);
            return InterpretReturnCode(myWLANLink.Receiver());
        }
    }

    public bool SetPiKoderPreferences(bool ProtectedSaveMode)
    {
        string CommandStr = "S";
        if (ProtectedSaveMode)
        {
            CommandStr += "U]U]";
        }
        if (iConnectedTo == iPhysicalLink.iSerialLink)
        {
            mySerialLink.SendDataToSerial(CommandStr);
            return InterpretReturnCode(mySerialLink.SerialReceiver());
        }
        else 
        {
            myWLANLink.SendDataToWLAN(CommandStr);
            return InterpretReturnCode(myWLANLink.Receiver());
        }
    }

    public bool SetPiKoderPPMSettings(int iNumberChannels, int myPPMmode)
    {
        string strSendString = Convert.ToString(iNumberChannels);
        if (myPPMmode == 0)
        {
            strSendString += "P";
        }
        else
        {
            strSendString += "N";
        }
        if (iConnectedTo == iPhysicalLink.iSerialLink)
        {
            return InterpretReturnCode(mySerialLink.SendDataToSerialwithAck("P=" + strSendString));
        }
        else
        {
            myWLANLink.SendDataToWLAN("P=" + strSendString);
            return InterpretReturnCode(myWLANLink.Receiver());
        }
    }


    public bool SetPiKoderPPMChannels(int iNumberChannels)
    {
        byte[] myByteArray = { 83, 21, 0, 0 };
        myByteArray[3] = Convert.ToByte(iNumberChannels);
        if (iConnectedTo == iPhysicalLink.iSerialLink)
        {
            if (Connected)
            {
                mySerialLink.SendBinaryDataToSerial(myByteArray, 4);
            }
            return true;
        }
        else
        {
            return false;
        }
    }

    public bool SetPiKoderPPMMode(int iPPMMode)
    {
        byte[] myByteArray = { 83, 22, 0, 0 };
        myByteArray[3] = Convert.ToByte(iPPMMode);
        if (iConnectedTo == iPhysicalLink.iSerialLink)
        {
            if (Connected)
            {
                mySerialLink.SendBinaryDataToSerial(myByteArray, 4);
            }
        }
        return true;
    }

    public bool SetPiKoderI2CBaseAdress(int iI2CBaseAdress)
    {
        string strSendString = "";
        if (iI2CBaseAdress < 10)
        {
            strSendString += "0";
        }
        strSendString += iI2CBaseAdress.ToString();
        if (iConnectedTo == iPhysicalLink.iSerialLink)
        {
            return InterpretReturnCode(mySerialLink.SendDataToSerialwithAck("I=" + strSendString));
        }
        else
        {
            myWLANLink.SendDataToWLAN("I=" + strSendString);
            return InterpretReturnCode(myWLANLink.Receiver());
        }
    }


    public bool PiKoderConnected()
    {
        if (iConnectedTo == iPhysicalLink.iSerialLink)
        {
            return mySerialLink.PiKoderConnected();
        }
        else
        {
            return false;
        }
    }


    private bool ValidateTypeRecord(string strVal)
    {
        if (strVal == "TimeOut")
        {
            return false;
        }
        if ((strVal[0] == 'T') & (strVal[1] == '='))
        {
            if (strVal.Length >= 5) 
            {
                return true;
            }
        }
        return false;
    }


    private bool ValidatePulseValue(ref string strVal, bool HPMath)
    {
        if (strVal == "TimeOut")
        {
            return false;
        }
        double intChannelPulseLength = 0;
        try
        {
            intChannelPulseLength = Double.Parse(strVal);  // no check on chars this time
        }
        catch
        {
            // format error, falsche message
            strVal = "FormatError";
            return false;
        }
        if (HPMath)
        {
            if ((intChannelPulseLength < 3750) | (intChannelPulseLength > 11250))
            {
                return false;
            }
            // format string
            if ((intChannelPulseLength < 10000) & (strVal.Length == 5))
            {
                strVal = strVal.Substring(1,4);                
            }
            return true;
        }
        else
        {
            if ((intChannelPulseLength < 750) | (intChannelPulseLength > 2250))
            {
                return false;
            }
            // format string
            if ((intChannelPulseLength < 1000) & (strVal.Length == 4))
            {
                strVal = strVal.Substring(1,3);
            }
            return true;
        }
    }

    private bool ValidateZeroOffset(string strVal)
    {
        if (strVal == "TimeOut")
        {
            return false;
        }
        int intZeroOffset = int.Parse(strVal);  // no check on chars this time
        if ((intZeroOffset < 0) | (intZeroOffset > 248))
        {
            return false;
        }
        return true;
    }

    private bool ValidateTimeOut(string strVal)
    {
        if (strVal == "TimeOut")
        {
            return false;
        }
        int intZeroOffset = int.Parse(strVal);  // no check on chars this time
        if ((intZeroOffset < 0) | (intZeroOffset > 999))
        {
            return false;
        }
        return true;
    }

    private bool ValidatePPMSettings(string strVal)
    {
        if (strVal[0] < '1' | strVal[0] > '8')
        {
            return false;
        }
        if (strVal[1] != 'P' & strVal[1] != 'N')
        {
            return false;
        }
        return true;
    }


    private bool ValidateI2CBaseAddress(string strVal)
    {
        if (strVal == "TimeOut")
        {
            return false;
        }
        int intZeroOffset = int.Parse(strVal);  // no check on chars this time
        if ((intZeroOffset < 0) | (intZeroOffset > 80))
        {
            return false;
        }
        return true;
    }

    private bool InterpretReturnCode(string RC)
    {
        if (RC.Contains("!"))
        {
            return true;
        }
        return false;
    }

    public void MyForm_Dispose()
    {
        Connected = false;  // make sure to force new connect
    }
}
// This class is designed to interface to a PiKoder through an available virtual serial
// port (COMn). 
//
// The class would be involved by EstablishSerialLink(). From this point onwards an 
// existing connection would be monitored and an event would be generated once the
// connection is lost (SerialLinkLost).
//
// The class provides for specialized methods to read and write PiKoder/COM information
// such as GetPulseLength() and SentPulseLength(). Please refer to the definitons for more
// details.
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
using System.IO.Ports;
using System.Threading;
using System.Text;
using System.Windows.Forms;

public class SerialLink
{
    static bool _connected = false;
    static SerialPort _serialPort = new SerialPort();
   
    public bool SerialLinkConnected()
    {
        return _connected;
    }

    public bool EstablishSerialLink(string SelectedPort) 
    { 
        if ((_serialPort.PortName == SelectedPort) & _connected) return true;  // port already connected
        if (_connected) _serialPort.Close();    // another port has been selected   
        try
        {
            _serialPort.PortName = SelectedPort;
            _serialPort.BaudRate = 9600;
            _serialPort.Parity = Parity.None;
            _serialPort.DataBits = 8;
            _serialPort.StopBits = StopBits.One;
            _serialPort.Handshake = Handshake.None;
            _serialPort.ReadTimeout = 100;
            _serialPort.WriteTimeout = 100;

            _serialPort.Open();
            _connected = true;
            return true;
        }
        catch (Exception ex)
        {
#if DEBUG
            MessageBox.Show(ex.ToString());
#endif
            _connected = false;
            return false;
        }
    }

    public string SerialReceiver ()
    {
        string _message = "";
        bool _receiving = true;
        bool _messageStarted = false;
        int _eomDetect = 2;
        int j = 0;
        byte _byte;
    
        while (_receiving)
        {
            if (_connected & (_serialPort.BytesToRead > 0))
            {
                _byte = (byte) _serialPort.ReadByte();
                if ((_byte != 0xD) & (_byte != 0xA)) 
                {
                    _message += Convert.ToChar(_byte);
                    _messageStarted = true;
                }
                else if (_messageStarted)
                {
                    _eomDetect -= 1;
                    if (_eomDetect == 0) 
                    {
                        _receiving = false;
                    }
                }
            }
            else
            {
                j += 1;
                if (j == 20)
                {
                    _receiving = false;
                    return "TimeOut";
                }
                System.Threading.Thread.Sleep(10);
            }
        }
        return _message; 
    }

    public string SendDataToSerialwithAck(string strWriteBuffer)
    {
        try
        {
            _serialPort.Write(strWriteBuffer);
            return SerialReceiver();
        }
        catch (Exception ex)
        {
#if DEBUG
            MessageBox.Show(ex.ToString());
#endif
            _connected = false;
            return "?";
        }
    }

    public void SendDataToSerial(string strWriteBuffer)
    {
        try
        {
            _serialPort.Write(strWriteBuffer);
        }
        catch (Exception ex)
        {
#if DEBUG
            MessageBox.Show(ex.ToString());
#endif
            _connected = false;
            _serialPort.Close();
        }
    }

    public void SendBinaryDataToSerial(byte[] myByteArray, int numBytes)
    {
        try
        {
            _serialPort.Write(myByteArray, 0, numBytes);
        }
        catch (Exception ex)
        {
#if DEBUG
            MessageBox.Show(ex.ToString());
#endif
            _connected = false;
        }
    }

    public void MyForm_Dispose()
    {
        try
        {
            _serialPort.Close();
            _connected = false;     //make sure to force new connect
        }
        catch (Exception ex)
        {
#if DEBUG
            MessageBox.Show(ex.ToString());
#endif
            _connected = false;
        }
    }

    // Sent a single char and retreive error message
    public bool PiKoderConnected()
    {
        if (_connected)
        {
            try
            {
                _serialPort.Write("*");
                if (SerialReceiver() == "?")
                {
                    return true;
                }
                _connected = false;
                return false;
            }
            catch (Exception ex)
            {
#if DEBUG
                MessageBox.Show(ex.ToString());
#endif
            }
        }
        _connected = false;
        return false;
    }
}
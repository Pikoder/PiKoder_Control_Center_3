// This is a program for evaluating the PiKoder platform - please refer to http://pikoder.com for more details.
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


using Suave.Utils;
using System;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Windows.Forms;

public class WLANLink
{
    // Set up variables
    private const int txPort = 12001;   // Port number to send data on
    private const int rxPort = 12000;   // Port number to recieve data on
    private const string AP_Address = "192.168.4.1";  // Standard IP address defined in radio
    private UdpClient receivingClient;  // Client for handling incxoming data
    private UdpClient sendingClient;    // Client for sending data
    private Thread receivingThread;     // Create a separate thread to listen for incoming data, helps to prevent the form from freezing up
    private bool closing = false;       // Used to close clients if form is closing
    private bool Connected = false;     // connection status
    private string MessageBuffer = "";
    private bool MessageFullyReceived = false;
   
    // Initialize listening & sending subs
    //
    public bool EstablishWLANLink()
    {
        InitializeSender();     // Initializes startup of sender client
        InitializeReceiver();   // Starts listening for incoming data                                             
        Connected = true;
        MessageFullyReceived = false;
        MessageBuffer = "";
        return Connected;
    }
 
    // Setup sender client
    //
    private void InitializeSender()
    {
        sendingClient = new UdpClient(AP_Address, txPort);
    }

    // Setup receiving client
    //
    private void InitializeReceiver()
    {
        receivingClient = new UdpClient(rxPort);
    }

    public string Receiver()
    {
        int iTimeOutCounter = 0;
        while ((!MessageFullyReceived) & (iTimeOutCounter < 5))
        {
            iTimeOutCounter++;
            Thread.Sleep(100);
        }
        if (iTimeOutCounter == 5)
        {
            try
            {
                receivingThread.Abort();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            MessageFullyReceived = false;   // free up buffer
            return "TimeOut";
        }
        MessageFullyReceived = false;   // free up buffer
        return MessageBuffer;
    }

    private void ReceiverThread()
    {
        IPEndPoint endPoint = new IPEndPoint(System.Net.IPAddress.Parse(AP_Address), rxPort);    // Listen for incoming data from any IP address on the specified port
        string myMessage = "";
        bool Receiving = true;
        bool messageStarted = false;
        int eomDetect = 2;
        MessageBuffer = ""; // make sure to start fresh
        while (Receiving)
        {
            try
            {
                byte[] rcvbytes = receivingClient.Receive(ref endPoint);    // Receive incoming bytes
                if (Buffer.ByteLength(rcvbytes) > 1)    // receive strings from ESP
                {
                    myMessage = System.Text.Encoding.ASCII.GetString(rcvbytes);
                    Receiving = false;
                }
                else
                {
                    if ((rcvbytes[0] != 0xD) & (rcvbytes[0] != 0xA))
                    {
                        myMessage += System.Text.Encoding.ASCII.GetString(rcvbytes); // Convert bytes back to string
                        messageStarted = true;
                    }
                    else if (messageStarted)
                    {
                        eomDetect--;
                        if (eomDetect == 0)
                        {
                            Receiving = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
#if DEBUG
                MessageBox.Show(ex.ToString());
#endif
            }
        }
        MessageBuffer = String.Copy(myMessage);
        MessageFullyReceived = true;
    }

    public bool WLANLinkConnected()
    {
        return Connected;
    }

    public void SendDataToWLAN(string strWriteBuffer)
    {
        byte[] sendbytes = Encoding.ASCII.GetBytes(strWriteBuffer);
        ThreadStart start = new ThreadStart(ReceiverThread);
        receivingThread = new Thread(start)
        {
            IsBackground = true
        };
        receivingThread.Start();
        try
        {
            sendingClient.Send(sendbytes, sendbytes.Length);
        }
        catch (Exception ex)
        {
            // MessageBox.Show(ex.ToString());
        }
    }

    public void MyForm_Dispose()
    {
        try
        {
            receivingThread.Abort();
        }
        catch (Exception ex)
        {
            // MessageBox.Show(ex.ToString());
        }
        MessageFullyReceived = false;
        receivingClient.Close();
        sendingClient.Close();
    }
}

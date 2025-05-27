using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.IO.Ports;
using System.Runtime.InteropServices;

namespace LBHH_Red {
  public partial class LBHH_HWI {

    // internal Byte[] moMsg;                        // ICD Section 5 LO message

    //Recieve Lightning Originated Messages
    private void ReceiveLoMessages(BackgroundWorker bw) {
      /*
       * Background worker thread to receive Lightning Orginated LBHH Hardwire
       * Interface messages.  It should be started when a serial port (COMx) is
       * opened to attempt connecting to a device.  It should be stopped when a
       * device is disconnected (or application is closed).
       * Stop this thread by setting runMsgThread = false
       * First it checks to make sure serial port configured with neccessary buffer size.
       * When running, it sleeps a short time and then checks for COM port data to read
       * (it will also awaken if the moMsgEvent is set [do when runMsgThread set 'false']).
       * It performs the first level of parsing - it finds the hardwire interface message
       * wrapper, extracts the specified length message, and compares computed checksum
       * agains checksum in message.  If it all passes the sniff test, a call is made to
       * sort the message by LO MFT and extract its data.
       */
      string diagStr = DateTime.UtcNow.ToString("HH':'mm':'ss.fff");
      diagStr += " STARTING THREAD 'Receive LO Messages'\n";
      bw.ReportProgress((int)BackgroundUpdate.WriteLoToRtb, diagStr);

      byte[] comData;         // data read from the serial port COMx
      UInt16 dataIdx = 0;     // index into comData
      byte comByte;           // single byte from comData[dataIdx]
      int msgReadState = 0;
      string loResponse = ""; // LBHH originated response
      int maxPortRead = 0;    // will be set to value returned from COMx serial port

      Monitor.Enter(comPort);               // wait for and lock down access to the COM port
      try {
        maxPortRead = comPort.ReadBufferSize; // get port's Read Buffer size so can makeVisible it is sufficient
        if (MaxHwMsgLen > maxPortRead) {
          throw new Exception(string.Format("{0} read buffer size too small\n(have/need {1:G}/{2:G} bytes)",
                                            comPort.PortName, maxPortRead, MaxHwMsgLen));
        }
      } catch (Exception ex) {                                       // If we've hit a problem with the COM port
        runMsgThreads = false;                // Do NOT let the messaging threads continue to run
        string msg = "Error starting LO Message thread.\nException: " + ex.ToString();
        string cap = " LO Message Thread ";
        MessageBoxButtons btn = MessageBoxButtons.OK;
        MessageBoxIcon icon = MessageBoxIcon.Hand;
        bw.ReportProgress((int)BackgroundUpdate.WriteLoToRtb,
                            (DateTime.UtcNow.ToString("HH':'mm':'ss.fff") + cap + msg + "\n"));
        MessageBox.Show(msg, cap, btn, icon);
        bw.ReportProgress((int)BackgroundUpdate.MsgThreadFatality);
      } finally {
        comData = new byte[0];                // and intialize buffer for data read from COM port
      }
      Monitor.Exit(comPort);                // Ensure that we release the lock on accessing COM port
      // Thread initialization now complete and ready to
      while (runMsgThreads && comPort.IsOpen) // run main loop of the receive LO message thread
      {
        loMsgEvent.WaitOne(1, false);        // wait for event or check every 1/50 second for received message
                                             // 1/50 second would allow about 19 characters max to be buffered
        if (runMsgThreads && comPort.IsOpen)  // Ensure runMsgThreads didn't change while thread was blocking/waiting
        {
          try {
            // Acquire exclusive lock on serial port, read available data, and release lock
            Monitor.Enter(comPort);
            try {
              int bytesToRead = ((comPort != null) && comPort.IsOpen) ? comPort.BytesToRead : 0;
              if (0 < bytesToRead) {
                comData = new byte[bytesToRead];
                comPort.Read(comData, 0, bytesToRead);
              }
            } catch (Exception ex) {
              runMsgThreads = false;            // Do NOT let the messaging threads continue to run
              comData = new byte[0];
              string msg = "Error reading LO data from serial port.\nException: " + ex.ToString();
              string cap = " COM Port ";
              MessageBoxButtons btn = MessageBoxButtons.OK;
              MessageBoxIcon icon = MessageBoxIcon.Exclamation;
              bw.ReportProgress((int)BackgroundUpdate.WriteLoToRtb,
                                  (DateTime.UtcNow.ToString("HH':'mm':'ss.fff") + cap + msg + "\n"));
              MessageBox.Show(msg, cap, btn, icon);
              bw.ReportProgress((int)BackgroundUpdate.MsgThreadFatality);
            }
            Monitor.Exit(comPort);

            dataIdx = 0;
            while (comData.Length > dataIdx) {// extract LO data from COM port data by pulling
              comByte = comData[dataIdx];     // individual byte out of buffered COM port data
                                              //bw.ReportProgress((int)BackgroundUpdate.WriteLoToRtb, ("RECV: " + (char)comByte + "\n")); //for debug only
              switch (msgReadState) {
                case 0:
                  if (comByte == '¡') {       // Key on the delimiter for starting the stream
                    msgReadState++;   }       // when it is found start buffering data
                  break;
                case 1:
                  if (comByte == '¶') {       // Key on the delimiter for stopping the stream
                    msgReadState = 0;         // when it is found restart search for start value, but first process what we have
                    bw.ReportProgress((int)BackgroundUpdate.WriteLoToRtb, (DateTime.UtcNow.ToString("HH':'mm':'ss.fff") + "-> RCV: ¡" + loResponse + "¶\n"));
                    parseLoResponse(bw, loResponse, ref cfgParams);
                    loResponse = "";//reset string
                  } else if (loResponse.Length < MaxHwMsgLen) {
                    //Add character to string
                    loResponse = loResponse + ((char)comByte);
                  } else {
                    msgReadState = 0;
                    loResponse = "";//reset string
                  }
                  break;
              }
              dataIdx++;// index the next comData byte
            } //end while (comData.Length > dataIdx)
          } catch (Exception ex) {
            string msg = "Error processing LO data.\nException: " + ex.ToString();
            bw.ReportProgress((int)BackgroundUpdate.WriteLoToRtb,
                                (DateTime.UtcNow.ToString("HH':'mm':'ss.fff") + msg + "\n"));
            bw.ReportProgress((int)BackgroundUpdate.ShowModelessNotification, msg);
          } // end try-catch reading & processing serial port data

          if (0 < comData.Length)   {               // now done with data, so if there was any
            comData = new byte[0];  }               // clear out anything read from COM port
          loMsgEvent.Reset();                       // then await event at top of while loop
        }//end if runMsgThreads
      } //end while runMsgThreads

      bw.ReportProgress((int)BackgroundUpdate.DisableLoRespTimer);
      diagStr = DateTime.UtcNow.ToString("HH':'mm':'ss.fff");
      diagStr += " TERMINATING THREAD 'Receive LO Messages'\n";
      bw.ReportProgress((int)BackgroundUpdate.WriteLoToRtb, diagStr);
    } // end method ReceiveMoMessages

    //Parse message type and do relevant actions
    enum msg_t {LogOnly, LogInsight, LogNotice };
    private void parseLoResponse(BackgroundWorker bw, string response, ref ProgParams cfgparams)
    {
      string    msgId;
      bool      isNak;
      string[]  msgFields;
      string    msgToUser = "";
      msg_t     msgType = msg_t.LogOnly;

      isNak     = ('-' == response[0]);
      msgFields = response.Split(' ');
      msgId     = msgFields[0];

      //Remove any ack/nak character from message ID so code is cleaner
      if ('+' == msgId[0] || '-' == msgId[0])
      {
        msgId = msgId.Remove(0, 1);
      }

      if (msgId == "C.GOK") {
        if (loRespStatus == LoRespStatus.GOKreq) {
          //If Device Nack'd key change to user make sure device wasn't switched
          if (isNak) {
            msgToUser = string.Format("Configuration Failed: LBHH NAK'd key activation:" + cfgparams.SelectedinfilkeyString + " at memory index: " + cfgparams.Selectedinfilkey + ".\nEnsure key is in device and that device was not switched.");
            cfgParams.connected = false;
          } else {
            bw.ReportProgress((int)BackgroundUpdate.DisableLoRespTimer);
            cfgParams.connected = true;
          }
        }
      } else if (msgId == "C.OLT") {
        if (loRespStatus == LoRespStatus.OLTreq) {
          if (isNak) {	//Device Nack'd loading Obfuscation Look-up Table
            msgToUser = "Configuration Failed: LBHH NAK'd loading of\nobfuscation key.";
            cfgParams.connected = false;
          } else {
            cfgParams.connected = true;
            bw.ReportProgress((int)BackgroundUpdate.DisableLoRespTimer);
          }
        }
      } else if (msgId == "C.SWF") {
        if (loRespStatus == LoRespStatus.SWFreq) {
          //Device Nack'd waveform change
          if (isNak) {
            msgToUser = string.Format("Configuration Failed: LBHH NAK'd changing to waveform slot " + gpiParams.actOnWaveform);
            cfgParams.connected = false;
          } else {
            cfgParams.connected = true;
            bw.ReportProgress((int)BackgroundUpdate.DisableLoRespTimer);
          }
        }
      } else if (msgId == "C.WFI") {
        if (loRespStatus == LoRespStatus.CWFIreq) {
          // Device Nack'd waveform info change
          if (isNak) {
            msgToUser = string.Format("Configuration Failed: LBHH NAK'd loading Waveform Information for slot " + gpiParams.actOnWaveform);
            cfgParams.connected = false;
          } else {
            //Set flag illustrating that device successfully set Waveform
            cfgParams.connected = true;
            bw.ReportProgress((int)BackgroundUpdate.DisableLoRespTimer);
          }
        }
      } else if (msgId == "D.WFI") {
        if (loRespStatus == LoRespStatus.SWFIreq) {
          if (msgFields.Length == 6) {
            cfgParams.connected = true;
            // Selected Waveform Slot Number
            cfgparams.waveformInfo[Convert.ToInt32(msgFields[1]), 0] = msgFields[1];
            // Selected Waveform Family
            cfgparams.waveformInfo[Convert.ToInt32(msgFields[1]), 1] = msgFields[2];
            // Selected Waveform Device ID
            cfgparams.waveformInfo[Convert.ToInt32(msgFields[1]), 2] = msgFields[3];
            // Selected Waveform Transmit ID
            cfgparams.waveformInfo[Convert.ToInt32(msgFields[1]), 3] = msgFields[4];
            // Selected Waveform Name
            cfgparams.waveformInfo[Convert.ToInt32(msgFields[1]), 4] = msgFields[5];
            // cfgparams.SelWvfrmName = msgFields[5];
            bw.ReportProgress((int)BackgroundUpdate.DisableLoRespTimer);
          } else {
            cfgParams.connected = false;
            msgToUser = string.Format("Query Failed: Waveform Information corrupt. Expected six fields, got " + msgFields.Length.ToString() + ".\nPlease try programming the device again.  If problem persists contact Application issuer.");
          }
        }
      } else if (msgId == "D.GAK") {
        if (loRespStatus == LoRespStatus.GAKreq) {
          if (msgFields.Length == 3) {
            cfgparams.deviceId = msgFields[1];
            //Load Device ID
            cfgparams.SelectedinfilkeyString = msgFields[2];
            //Load Infil String
            cfgParams.infilActiveack = true;
            cfgParams.connected = true;
            bw.ReportProgress((int)BackgroundUpdate.DisableLoRespTimer);
          } else {
            cfgParams.connected = false;
            msgToUser = string.Format("Query Failed: Active Infil Key Names corrupt. Expected three fields, got " + msgFields.Length.ToString() + ".\nPlease try programming the device again. If problem persists contact Application issuer.");
          }
        }
      } else if (msgId == "D.GKN") {
        if (loRespStatus == LoRespStatus.GKNreq) {
          if (msgFields.Length == 7) {
            cfgparams.deviceId  = msgFields[1];   //Unique InfilKey name
            cfgparams.InfilKey1 = msgFields[2];   //Group1 InfilKey name
            cfgparams.InfilKey2 = msgFields[3];   //Group2 InfilKey name
            cfgparams.InfilKey3 = msgFields[4];   //Group3 InfilKey name
            cfgparams.InfilKey4 = msgFields[5];   //Group4 InfilKey name
            cfgparams.InfilKey5 = msgFields[6];   //Group5 InfilKey name
            cfgParams.infilNameack = true;
            cfgParams.connected = true;           //Available Device Keys also used to 'find' device
            bw.ReportProgress((int)BackgroundUpdate.DisableLoRespTimer);
          } else {
            cfgParams.connected = false;
            msgToUser = string.Format("Query Failed: Infil Key Names corrupt. Expected seven fields, got " + msgFields.Length.ToString() + ".\nPlease try programming the device again if problem persists contact Application issuer.");
          }
        }
      //} else if (msgId == "D.WFS") {// (as of 2022/05/01 this no longer used except in dead code)
      //	if (loRespStatus == LoRespStatus.WFSreq) {
      //		if (msgFields.Length == 6) {
      //			cfgParams.connected = true;
      //			// Selected Waveform Slot Number
      //			cfgparams.SelWvfrmSlotNum = msgFields[1];
      //			// Selected Wafeform Family
      //			cfgparams.SelWvfrmFamily = msgFields[2];
      //			// Selected Waveform Device ID
      //			cfgparams.SelWvfrmDeviceId = msgFields[3];
      //			// Selected Waveform Transmit ID
      //			cfgparams.SelWvfrmTransmitId = msgFields[4];
      //			// Selected Waveform Name
      //			cfgparams.SelectedwaveformString = msgFields[5];
      //			// cfgparams.SelWvfrmName = msgFields[5];
      //			bw.ReportProgress((int)BackgroundUpdate.DisableLoRespTimer);
      //		} else {
      //			cfgParams.connected = false;
      //			msgToUser = string.Format("Query Failed: Selected Waveform Information corrupt. Expected six fields, got " + msgFields.Length.ToString() + ".\nPlease try programming the device again if problem persists contact Application issuer.");
      //		}
      //	}
      //} else if (msgId == "D.WFN") {	// Return loaded waveform names  // (as of 2022/05/01 this no longer used except in dead code)
      //	if (loRespStatus == LoRespStatus.WFNreq) {
      //		if (msgFields.Length == 5) {
      //			cfgParams.connected = true;
      //			//Load InfilKey 1
      //			cfgparams.Waveform1 = msgFields[1];
      //			//Load InfilKey 2
      //			cfgparams.Waveform2 = msgFields[2];
      //			//Load InfilKey 3
      //			cfgparams.Waveform3 = msgFields[3];
      //			//Load InfilKey 4
      //			cfgparams.Waveform4 = msgFields[4];
      //			bw.ReportProgress((int)BackgroundUpdate.DisableLoRespTimer);
      //		} else {
      //			cfgParams.connected = false;
      //			msgToUser = string.Format("Query Failed: Names of Loaded Waveforms corrupt. Expected five fields, got " + msgFields.Length.ToString() + ".\nPlease try programming the device again if problem persists contact Application issuer.");
      //		}
      //	}
      } else if (msgId == "O.ECP") {	//This guy doesn't seem to work - LBHH switching COMx before ACK/NAK emitted
        if (loRespStatus == LoRespStatus.GKNreq) {
          //Opened Bolt Debug Port Successfully
          if (isNak) { //Device Nack'd
            msgToUser = "Configuration Failed: Device NAK'd opening Bolt Debug.\nCycle Power.  If this happens again contact Hardware provider.";
            cfgParams.connected = false;
          } else {
            cfgParams.connected = true;
            bw.ReportProgress((int)BackgroundUpdate.DisableLoRespTimer);
          }
        }
      } else if (msgId == "C.GLK") {
        if (loRespStatus == LoRespStatus.GLKreq) {
          //Received Response for Load Key Command
          if (isNak) {
            msgToUser = "Configuration Failed: Device NAK'd loading infil key.\nCycle Power.  If this happens again contact Hardware provider.";
            cfgParams.connected = false;
          } else {
            cfgParams.connected = true;
            bw.ReportProgress((int)BackgroundUpdate.DisableLoRespTimer);
          }
        }
      } else if (msgId == "C.GZK") {
        if (LoRespStatus.GZKreq == loRespStatus) {
          if (isNak) {
            msgToUser = "Configuration Failed: Device NAK'd wiping keys.\nCycle Power.  If this happens again contact Hardware provider.";
            cfgParams.connected = false;
          } else {
            cfgParams.connected = true;
            bw.ReportProgress((int)BackgroundUpdate.DisableLoRespTimer);
          }
        }
      } else if (msgId == "H.ATE") {
        if (LoRespStatus.ATE0req == loRespStatus) {
          if (isNak) {
            msgToUser = "Configuration Failed: Device NAK'd Anti-Tamper Disable \nCycle Power.  If this happens again contact Hardware provider.";
            cfgParams.connected = false;
            msgType = msg_t.LogNotice;
         } else {
            cfgParams.connected = true;
            bw.ReportProgress((int)BackgroundUpdate.DisableLoRespTimer);
            msgToUser = "Anti-Tamper Disabled\n\n"
                + "Before opening unit you must\n"
                + "   1) power  off  LBHH\n"
                + "   2) power  ON  LBHH\n"
                + "and keep power -ON- when open.\n";
            msgType = msg_t.LogNotice;
          }
        } else if (LoRespStatus.ATE1req == loRespStatus) {
          if (isNak) {
            msgToUser = "Configuration Failed: Device NAK'd Anti-Tamper Enable \nCycle Power.  If this happens again contact Hardware provider.";
            cfgParams.connected = false;
            msgType = msg_t.LogNotice;
          } else {
            cfgParams.connected = true;
            bw.ReportProgress((int)BackgroundUpdate.DisableLoRespTimer);
            msgToUser = "Anti-Tamper Enabled";
            msgType = msg_t.LogInsight;
          }
        }
      }
      else if (msgId == "D.RST")
      {
      }
      else if (isNak && (msgId.Contains("S.")))
      { // Generally ignore data subscription ACK responses -- but any NAK indicates firmware incompatibility
        msgToUser = "Device appears incompatible with this application.\nContact issuer of LBHH and Application";
        msgType = msg_t.LogNotice;
      }

      if (0 < msgToUser.Length)
      {
        bw.ReportProgress((int)BackgroundUpdate.WriteLoToRtb, (DateTime.UtcNow.ToString("HH':'mm':'ss.fff") + ": " + msgToUser + "\n"));
        if (msg_t.LogNotice == msgType)                                                   {
          bw.ReportProgress((int)BackgroundUpdate.ShowModelessNotification, msgToUser);   }
        else if (msg_t.LogInsight == msgType)                                             {
          bw.ReportProgress((int)BackgroundUpdate.ShowModelessInsight, msgToUser);        }
      }

    }
  }
}

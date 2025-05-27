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

    // These Programmed Params will be used to ensure the device parameters were properly set

    const int MaxHwMsgLen = 270;   // Max msg is 270-bytes per ICD
    const int HwMtMftOffset = 4;   // LT MFT is at [4]
//  const int HwMtMcOffset = 9;    // Message Count is at [9]
    private ArrayList ltMsgQue = new ArrayList();
    const string NO_ENC_KEY = "0000000000000000000000000000000000000000000000000000000000000000";

      private struct LtMsgReq {
      public string msg;                // array to hold raw hardwire interface message bytes
      public Int32 timeout;             // milliseconds allowed for LO response
    //public Byte[] expectedAck;        // The anticipated Ack Value String (Anything else is considered a NAK)
      public LightningMsgType ltMsg;    // Message type
      public LoRespStatus reqLoResp;
    } // end struct MtMsgReq


    //Type of Message to be sent
    public enum LightningMsgType {
    //qSelWvfrmInfo,  // Query Selected Waveform Information (S.WFS) // (as of 2022/05/01 this no longer used except in dead code)
    //qWfNames,       // Query Loaded Waveform Names (S.WFN) (as of 2022/05/01 this no longer used except in dead code)
      dInfilKeys,     // Delete all Infil Keys (C.GZK)
      lInfilKey,      // Load an Infil Key (C.GLK)
      lObfKey,        // Load an Obfuscation Key (C.OLT)
      lWvfrmInfo,     // Load waveform information (C.WFI)
      qActvKeyNames,  // Query active infil key (S.GAK)
      qKeyNames,      // Query available Infil Keys (S.GKN)
      qWvfrmInfo,     // Query All Loaded Waveform Information (S.WFI)
      sAntiTampOff,   // Set Anti-Tamper OFF (H.ATE 0)
      sAntiTampOn,    // Set Anti-Tamper ON (H.ATE 1)
      sCOMxBoltDbg,   // Set COMx to BOLT Debug Port comms (O.ECP) (NOTE:Baud Rate needs to be changed for proper comms)
      sGrpInfil,      // Selects active Group Infil Key (C.GOK)
      sWaveForm,      // Selects active Waveform on Lightning side (C.SWF)
      sendLtngUsrCmd, // Send Lightning a user entered command from cmdTxtBox in Comms tab
//    sendBoltTek,    // Send BOLT the waveform's TEK key // (as of 2022/05/01 this no longer used except in dead code)
//    sendBoltWfPswd  // Send BOLT the waveform's password // (as of 2022/05/01 this no longer used except in dead code)
    };

    ////Build Query Waveform Names message
    //private void BuildQueryWaveformNames(ref string msg) {// (as of 2022/05/01 this no longer used except in dead code)
    //	//Issue here with serial
    //	msg = "¡S.WFN¶";
    //}

    ////Build Query Selected Waveform Information message
    //private void BuildQuerySelWvfrmInfo(ref string msg) {// (as of 2022/05/01 this no longer used except in dead code)
    //	msg = "¡S.WFS¶";
    //}

    //Build Query Waveform Information message
    private void BuildQueryWaveformInfo(ref string msg) {
      msg = "¡S.WFI¶";
    }

    //Build Configure Waveform Information message
    private void BuildLoadWaveformInfo(ref string msg, string[,] waveformInfo, int currentWaveform) {
      msg = "¡C.WFI " + waveformInfo[currentWaveform, 0] + " "
                             + waveformInfo[currentWaveform, 1] + " "
                             + waveformInfo[currentWaveform, 2] + " "
                             + waveformInfo[currentWaveform, 3] + " "
                             + waveformInfo[currentWaveform, 4] + "¶";
    }

    //Build Infil Poll message
    private void BuildQueryInfilKeyNames(ref string msg) {
      msg = "¡S.GKN¶";
    }

    //Build Delete Key Message
    private void BuildDeleteInfilKeys(ref string msg) {
      msg = "¡C.GZK A¶";
    }

    //Build Select Infil message
    private void BuildSelectGrpInfil(ref string msg, int userSelection) {
      if (userSelection > 0) {
        msg = "¡C.GOK " + userSelection.ToString() + "¶";
      } else {
        msg = "¡C.GOK N¶";
      }
    }
    private void BuildZeroizeOptionalGDB(ref string msg) {
      msg = "¡C.GZK G¶";
    }

    //Build Select Infil message
    private bool BuildLoadInfilKey(ref string msg, int sendKeyselection) {
      if (sendKeyselection == 1) {
        if (gpiParams.cfgInfil1grpindx >= 0) {
          msg = "¡C.GLK " + sendKeyselection.ToString() + " " + infilKeys.GroupKeys[gpiParams.cfgInfil1grpindx].GroupName + " " + infilKeys.GroupKeys[gpiParams.cfgInfil1grpindx].GroupKey + "¶";
        }
      } else if (sendKeyselection == 2) {
        if (gpiParams.cfgInfil2grpindx >= 0) {
          msg = "¡C.GLK " + sendKeyselection.ToString() + " " + infilKeys.GroupKeys[gpiParams.cfgInfil2grpindx].GroupName + " " + infilKeys.GroupKeys[gpiParams.cfgInfil2grpindx].GroupKey + "¶";
        }
      } else if (sendKeyselection == 3) {
        if (gpiParams.cfgInfil3grpindx >= 0) {
          msg = "¡C.GLK " + sendKeyselection.ToString() + " " + infilKeys.GroupKeys[gpiParams.cfgInfil3grpindx].GroupName + " " + infilKeys.GroupKeys[gpiParams.cfgInfil3grpindx].GroupKey + "¶";
        }
      } else if (sendKeyselection == 4) {
        if (gpiParams.cfgInfil4grpindx >= 0) {
          msg = "¡C.GLK " + sendKeyselection.ToString() + " " + infilKeys.GroupKeys[gpiParams.cfgInfil4grpindx].GroupName + " " + infilKeys.GroupKeys[gpiParams.cfgInfil4grpindx].GroupKey + "¶";
        }
      } else if (sendKeyselection == 5) {
        if (gpiParams.cfgInfil5grpindx >= 0) {
          msg = "¡C.GLK " + sendKeyselection.ToString() + " " + infilKeys.GroupKeys[gpiParams.cfgInfil5grpindx].GroupName + " " + infilKeys.GroupKeys[gpiParams.cfgInfil5grpindx].GroupKey + "¶";
        }
      } else if (sendKeyselection == 0) {
        if (infilKeys.deviceID < 10) {
          msg = "¡C.GLK U ID000" + infilKeys.deviceID + " " + infilKeys.uniqueKey + "¶";
        } else if (infilKeys.deviceID < 100) {
          msg = "¡C.GLK U ID00" + infilKeys.deviceID + " " + infilKeys.uniqueKey + "¶";
        } else if (infilKeys.deviceID < 1000) {
          msg = "¡C.GLK U ID0" + infilKeys.deviceID + " " + infilKeys.uniqueKey + "¶";
        } else if (infilKeys.deviceID < 10000) {
          msg = "¡C.GLK U ID" + infilKeys.deviceID + " " + infilKeys.uniqueKey + "¶";
        } else if (infilKeys.deviceID < 100000) {
          msg = "¡C.GLK U I" + infilKeys.deviceID + " " + infilKeys.uniqueKey + "¶";
        } else {
          return false; // Unique ID impossibly large
        }
      } else {
        return false;   // Invalid sendKeyselection parameter
      }
      return true;
    }

    //Build Select Waveform message
    private bool BuildLoadObfKey(ref string msg, int cmdIdx) {
      if (0 > cmdIdx || numObfKeys <= cmdIdx) {
        return false;                         }
      msg = "¡C.OLT " + obfKeys.keys[cmdIdx] + "¶";
      return true;
    }

    //Build Select Waveform message
    private void BuildSelectWvfrm(ref string msg, int userSelection) {
      msg = "¡C.SWF " + userSelection.ToString() + "¶";
    }

    //Build Active Infil message
    private void BuildQueryActvInfil(ref string msg) {
      msg = "¡S.GAK¶";
    }


    //Build Anti Tamper On Command
    private void BuildSetAntiTampOn(ref string msg) {
      msg = "¡H.ATE 1¶";
    }

    //Build Anti Tamper Off Command
    private void BuildSetAntiTampOff(ref string msg) {
      msg = "¡H.ATE 0¶";
    }


    //Build Cmd value
    private void BuildLtngUserCmd(ref string msg, string cmd) {
      msg = "¡" + cmd + "¶";
    }


    //Build Active Infil message
    private void BuildSetCOMxToBoltDebug(ref string msg) {
      msg = "¡O.ECP 2¶";
    }

    ////Build Bolt Password
    //private void BuildBoltWfPassword(ref string msg) {
    //  msg = "??*GPSTimeOut\r\n";
    //}

    ////Build Bolt Waveform Key
    //private void BuildBoltTek(ref string msg, string key) {
    //  msg = key + "\r\n";
    //}


    //Because Messages are relatively simple this will build the message and add the message to the TX queue to the Lightning
    private void QueueLtMessage(LightningMsgType msgType) {
      LtMsgReq ltMsgReq = new LtMsgReq();       // Initialize as JobSeparator
      ltMsgReq.msg = "";                        // which is an empty message
      ltMsgReq.timeout = 0;                     // and won't start timer
      ltMsgReq.ltMsg = msgType;                 // (except initialize w/provided msg UID)
      try {

        switch (msgType) {
          //case LightningMsgType.qWfNames: // (as of 2022/05/01 this no longer used except in dead code)
          //	// mtMsgReq alreay initialized for this empty message (which will never really get sent)
          //	BuildQueryWaveformNames(ref ltMsgReq.msg);
          //	ltMsgReq.reqLoResp = LoRespStatus.WFNreq;
          //	ltMsgReq.timeout = 500;
          //	break;
          //case LightningMsgType.qSelWvfrmInfo:// (as of 2022/05/01 this no longer used except in dead code)
          //	BuildQuerySelWvfrmInfo(ref ltMsgReq.msg);
          //	ltMsgReq.reqLoResp = LoRespStatus.WFSreq;
          //	ltMsgReq.timeout = 500;
          //	break;

          case LightningMsgType.dInfilKeys:
            BuildDeleteInfilKeys(ref ltMsgReq.msg);
            ltMsgReq.reqLoResp = LoRespStatus.GZKreq;
            ltMsgReq.timeout = 500;
            break;

          case LightningMsgType.lInfilKey:
            if ( ! BuildLoadInfilKey(ref ltMsgReq.msg, gpiParams.keyLoad)) {
              ltMsgReq.timeout = 500;
              break;
            }
            ltMsgReq.reqLoResp = LoRespStatus.GLKreq;
            ltMsgReq.timeout = 500;
            break;

          case LightningMsgType.lObfKey:
            if ( ! BuildLoadObfKey(ref ltMsgReq.msg, obfKeys.keyIdx)) {
              ltMsgReq.timeout = 500;
              break;
            }
            ltMsgReq.reqLoResp = LoRespStatus.OLTreq;
            ltMsgReq.timeout = 500;
            break;

          case LightningMsgType.lWvfrmInfo:
            BuildLoadWaveformInfo(ref ltMsgReq.msg, gpiParams.waveformInfo, gpiParams.actOnWaveform);
            ltMsgReq.reqLoResp = LoRespStatus.CWFIreq;
            ltMsgReq.timeout = 500;
            break;

          case LightningMsgType.qActvKeyNames:
            BuildQueryActvInfil(ref ltMsgReq.msg);
            ltMsgReq.reqLoResp = LoRespStatus.GAKreq;
            ltMsgReq.timeout = 500;
            break;

          case LightningMsgType.qKeyNames:
            BuildQueryInfilKeyNames(ref ltMsgReq.msg);
            ltMsgReq.reqLoResp = LoRespStatus.GKNreq;
            ltMsgReq.timeout = 500;
            break;

          case LightningMsgType.qWvfrmInfo:
            BuildQueryWaveformInfo(ref ltMsgReq.msg);
            ltMsgReq.reqLoResp = LoRespStatus.SWFIreq;
            ltMsgReq.timeout = 500;
            break;

          case LightningMsgType.sAntiTampOff:
            BuildSetAntiTampOff(ref ltMsgReq.msg);
            ltMsgReq.reqLoResp = LoRespStatus.ATE0req;
            ltMsgReq.timeout = 500;
            break;

          case LightningMsgType.sAntiTampOn:
            BuildSetAntiTampOn(ref ltMsgReq.msg);
            ltMsgReq.reqLoResp = LoRespStatus.ATE1req;
            ltMsgReq.timeout = 500;
            break;

          case LightningMsgType.sCOMxBoltDbg:
            BuildSetCOMxToBoltDebug(ref ltMsgReq.msg);
            WriteRichTextBox(diagRichTextBox, "------------OPENING BOLT COMMUNICATION------------\n");
            ltMsgReq.reqLoResp = LoRespStatus.ECPreq;
            ltMsgReq.timeout = 1000;
            break;

          case LightningMsgType.sGrpInfil:
            BuildSelectGrpInfil(ref ltMsgReq.msg, gpiParams.Selectedinfilkey);
            ltMsgReq.reqLoResp = LoRespStatus.GOKreq;
            ltMsgReq.timeout = 500;
            break;

          case LightningMsgType.sWaveForm:
            BuildSelectWvfrm(ref ltMsgReq.msg, gpiParams.actOnWaveform);
            ltMsgReq.reqLoResp = LoRespStatus.SWFreq;
            ltMsgReq.timeout = 500;
            break;

          case LightningMsgType.sendLtngUsrCmd:
            BuildLtngUserCmd(ref ltMsgReq.msg, gpiParams.cmdTextbox);
            ltMsgReq.reqLoResp = LoRespStatus.CMDreq;
            ltMsgReq.timeout = 0;
            break;

          //case LightningMsgType.sendBoltTek:
          //  BuildBoltTek(ref ltMsgReq.msg, gpiParams.boltKey);//Global Key String
          //  ltMsgReq.reqLoResp = LoRespStatus.Null;
          //  ltMsgReq.timeout = 5000;
          //  break;

          //case LightningMsgType.sendBoltWfPswd:
          //  BuildBoltWfPassword(ref ltMsgReq.msg);
          //  ltMsgReq.reqLoResp = LoRespStatus.Null;
          //  ltMsgReq.timeout = 5000;
          //  break;

          default:
            throw new Exception(string.Format("Request to queue invalid LT message type: {0}", msgType));
        }

        if (ltMsgReq.msg.ToString() != "") {
          Monitor.Enter(ltMsgQue);
          try {
            ltMsgQue.Add(ltMsgReq);
            ltMsgEvent.Set();
          } catch (Exception ex) {
            string msg = "Exception: " + ex.Message;
            string cap = " Error queuing LT message ";
            //WriteRichTextBox(diagRichTextBox, (System.DateTime.UtcNow.ToString("HH':'mm':'ss.fff") + cap + msg + "\n"));
            ShowModelessMessageBox(msg, cap, InfoType.Insight);
          } finally {
            ltMsgEvent.Set();                     // Set Send Lightning Terminated messages' event and
            Thread.Sleep(0);                      // give it time to see the new message
          }
          Monitor.Exit(ltMsgQue);
        }
      } catch (Exception ex) {
        string msg = "Exception: " + ex.Message;
        string cap = " Error building LT message ";
        //WriteRichTextBox(diagRichTextBox, (System.DateTime.UtcNow.ToString("HH':'mm':'ss.fff") + cap + msg + "\n"));
        ShowModelessMessageBox(msg, cap, InfoType.Insight);
      }
    }

    private void SendLtMessages(BackgroundWorker bw) {
      /*
       * Background worker thread to send Mobile Terminated LBHH Hardwire
       * Interface messages.  It should be started when a serial port (COMx) is
       * opened to attempt connecting to a device.  It should be stopped when a
       * device is disconnected (or application is closed).
       * Stop this thread by setting runMsgThread = false
       */
      string diagStr = DateTime.UtcNow.ToString("HH':'mm':'ss.fff");
      diagStr += " STARTING THREAD 'Send LT Messages'\n";
      bw.ReportProgress((int)BackgroundUpdate.WriteLtToRtb, diagStr);

      string msgStr = "";                     // local copy of command string from ltMsgQue
      byte[] comData = new byte[0];           // local buffer to posting data to COMx
      int maxPortWrite = 0;                   // will be set to value returned from COMx serial port
      Int32 timeOut = 0;                      // used to set the moRespTimer based on value from queued message
      Monitor.Enter(ltMsgQue);
      ltMsgQue.Clear();                       // nothing in the LT message queue (no hot-swap!)
      Monitor.Exit(ltMsgQue);

      Monitor.Enter(comPort);
      try {
        maxPortWrite = comPort.WriteBufferSize; // get port's Write Buffer size so can makeVisible it is sufficient
        if (MaxHwMsgLen > maxPortWrite) {
          throw new Exception(string.Format("{0} write buffer size too small\n(have/need {1:G}/{2:G} bytes)",
                                            comPort.PortName, maxPortWrite, MaxHwMsgLen));
        }
      } catch (Exception ex) {                                       // If we've hit a problem with the COM port
        runMsgThreads = false;                // Do NOT let the messaging threads continue to run
        string msg = "Error starting LT Message thread.\nException: " + ex.ToString();
        string cap = " LT Message Thread ";
        MessageBoxButtons btn = MessageBoxButtons.OK;
        MessageBoxIcon icon = MessageBoxIcon.Hand;
        bw.ReportProgress((int)BackgroundUpdate.WriteLtToRtb,
                          (DateTime.UtcNow.ToString("HH':'mm':'ss.fff") + cap + msg + "\n"));
        MessageBox.Show(msg, cap, btn, icon);
        bw.ReportProgress((int)BackgroundUpdate.MsgThreadFatality);
      }
      Monitor.Exit(comPort);

      // Thread initialization now complete and ready to
      while (runMsgThreads && comPort.IsOpen) // run main loop of the send LT message thread
      {
        ltMsgEvent.WaitOne(200, false);       //wait for an event or check every 1/5 second for message to send

        if (runMsgThreads && comPort.IsOpen)  //Ensure runMsgThreads didn't change while thread was blocking/waiting
        { // Acquire exclusive lock on serial port, write LT data, and release lock
          Monitor.Enter(ltMsgQue);
          if ((LoRespStatus.Null < loRespStatus) && (0 < ltMsgQue.Count)) { // When not waiting for a LO response and a message is queued up, process next message
            LtMsgReq ltMsgReq = (LtMsgReq)ltMsgQue[0];  // Get all data related to next message
            ltMsgQue.RemoveAt(0);                       // then remove message from the queue
            msgStr = ltMsgReq.msg;                      // yeah, if I were to expand its scope I could use ltMstReq.msg where I use msgStr...


            if (HwMtMftOffset < ltMsgReq.msg.Length) { // Only when not encountering special 'Job Separator' can/should we post diag string
              diagStr = DateTime.UtcNow.ToString("HH':'mm':'ss.fff");
              diagStr += " Sending LT " + ltMsgReq.ltMsg;
              diagStr += ", length = " + ltMsgReq.msg.Length;
              diagStr += string.Format(", timeout(ms) = {0}, LO Response Status = {1}, HWI MSG ->",
                                        ltMsgReq.timeout, ltMsgReq.reqLoResp);
              diagStr += ltMsgReq.msg + "\n";
              //bw.ReportProgress((int)BackgroundUpdate.WriteLoToRtb, diagStr);
            }
             
            //comData = ltMsgReq.msg;                     // Set message data into buffer.
            comData = new byte[(msgStr.Length) * sizeof(char)];
            for (int i = 0; i < msgStr.Length; i++) {
              comData[i] = (byte) msgStr[i];        }
            loRespStatus = ltMsgReq.reqLoResp;          // Note the type of response expected so Bo thread knows what to look for
            timeOut = ltMsgReq.timeout;                 // and time allowed for that response
            activeLtMsg = ltMsgReq.ltMsg;               // Remember type of message to send
          }
          Monitor.Exit(ltMsgQue);

          if (0 < comData.Length) {                         // When there is message buffered to send
            Monitor.Enter(comPort);                         // obtain serial port so we can...
            try {
              if ((comPort != null) && comPort.IsOpen)      // ensure port exist and is open
              {                                             // before we try to
                comPort.Write(comData, 0, comData.Length);  // write data out the COM port
                if (timeOut > 0)                            // When message has a timeout value
                {                                           // start the timer with that timeout
                  bw.ReportProgress((int)BackgroundUpdate.StartLoRespTimer, timeOut);
                }
                bw.ReportProgress((int)BackgroundUpdate.WriteLtToRtb, (DateTime.UtcNow.ToString("HH':'mm':'ss.fff") + "-> CMD: " + msgStr + "\n"));
              }
            } //end try
            catch (Exception ex) {
              runMsgThreads = false;          // Do NOT let the messaging threads continue to run
              string msg = "Error writing data to serial port.\n" + comPort.PortName + " Exception:\n" + ex.ToString();
              string cap = " COM Port ";
              MessageBoxButtons btn = MessageBoxButtons.OK;
              MessageBoxIcon icon = MessageBoxIcon.Exclamation;
              bw.ReportProgress((int)BackgroundUpdate.WriteLoToRtb,
                                (DateTime.UtcNow.ToString("HH':'mm':'ss.fff") + cap + msg + "\n"));
              MessageBox.Show(msg, cap, btn, icon);
              bw.ReportProgress((int)BackgroundUpdate.MsgThreadFatality);
            } finally {
              comData = new byte[0];
            }
            Monitor.Exit(comPort);
          } //end if message != ""

          ltMsgEvent.Reset(); //reset the wait for the next send command button press
        } //end if runMsgThreads
      } //end while runMsgThreads

      diagStr = DateTime.UtcNow.ToString("HH':'mm':'ss.fff");
      diagStr += " TERMINATING THREAD 'Send LT Messages'\n";
      bw.ReportProgress((int)BackgroundUpdate.WriteLoToRtb, diagStr);
    } //end method SendMtMessages



    //Stops message send queue
    private int DiscardLtMsgJob() {
      /*
       * Remove LT messages from queue until find one that does not require a LO reply;
       * the 'JobSeparator' messages fit this description (as do ACK/NAK which don't
       * foresee the desire to discard - thus don't discard based on non-zero length).
       * Returns the number of LT messages that were discarded
       * NOTE - WFHWI.msgCnt will be thrown off by the number of messages discarded
       *        but that doesn't really matter
       */
      LtMsgReq ltMsgReq;      // copy of message to potentially be discarded
      int startCount;         // how many messages in queue at start of process
      int deltaCount;         // how many messages were removed from queue

      Monitor.Enter(ltMsgQue);
      startCount = ltMsgQue.Count;                        // Get number of messages at start of this
      while (0 < ltMsgQue.Count) {                                                   // Loop, potentially through all messages.
        ltMsgReq = (LtMsgReq)ltMsgQue[0];                 // Look at first message remaining in queue
        if (LoRespStatus.Null != ltMsgReq.reqLoResp) {   // if message requires response
          ltMsgQue.RemoveAt(0);
        }   // remove that message
        else {   // otherwise
          break;
        }   // stop looking, done with removals
      }
      deltaCount = startCount - ltMsgQue.Count;           // Calculate number of messages removed
      Monitor.Exit(ltMsgQue);

      return deltaCount;
    } // end method DiscardMtMsgJob



  }
}

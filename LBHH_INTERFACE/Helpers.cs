using System;
using System.ComponentModel;
using System.Drawing;
using System.Media;             // for SystemSounds.Beep
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;

namespace LBHH_Red {
  public partial class LBHH_HWI {
    public static byte GetNibble(char hexChar) {
      if (hexChar >= '0' && hexChar <= '9')
        return (byte)(hexChar - '0');
      else if (hexChar >= 'a' && hexChar <= 'f')
        return (byte)(0x0A + (byte)(hexChar - 'a'));
      else
        return (byte)(0x0A + (byte)(hexChar - 'A'));
    }


    private void BackgroundWorkerUpdate(object sender, ProgressChangedEventArgs e) {
      /*
      * Could have made our background workers ProgressChanged event handler just point to this
      * Rather than expecting a integer representing progress percentage, send casted
      * BackgroundUpdate enumeration.  Other required supporting data (e.g. a string),
      * retrieved by properly casting object accessed via e.UserState.
      */
      switch (e.ProgressPercentage) {
        case (int)BackgroundUpdate.Null:
          break;

        case (int)BackgroundUpdate.MsgThreadFatality:
          try {
            comPort.Close();   // this will throw an exception if COM port unavailable
            comPort.Dispose(); // this will throw an exception if COM port unavailable
          } catch (Exception) {
            ; // We could expect an exception - no worries, catch & discard
          } finally {
            comPort = new SerialPort(); // Ensure comPort is left neither connected nor NULL
          }
          break;

        case (int)BackgroundUpdate.DisableLoRespTimer: //For determining how long it took for message to be recieved
          loRespTimer.Enabled = false;
          break;

        case (int)BackgroundUpdate.StartLoRespTimer:   //For determining how long it took for message to be recieved
          loRespTimer.Interval = (int)e.UserState;
          loRespTimer.Enabled = true;
          break;

        case (int)BackgroundUpdate.WriteLoToRtb:
          string loStr = (string)e.UserState;
          ParseLightningText(loStr);                    // TODO 2022/05/01 -- this is just SO wrong to do.  Use LtMsgThread and LoMsgThread !!!
          WriteRichTextBox(diagRichTextBox, loStr);
          break;

        case (int)BackgroundUpdate.WriteLtToRtb:
          string ltStr = (string)e.UserState;
          ParseBoltText(ltStr);
          WriteRichTextBox(diagRichTextBox, ltStr);
          break;

        case (int)BackgroundUpdate.ShowModelessInsight:
          ShowModelessMessageBox((string)e.UserState, "", InfoType.Insight);
          break;

        case (int)BackgroundUpdate.ShowModelessNotification:
          ShowModelessMessageBox((string)e.UserState, "", InfoType.Notification);
          break;

        default:
          string msg = "Unhandled Progress Report: " + e.ProgressPercentage.ToString();
          string cap = " Background Worker ";
          ShowModelessMessageBox(msg, cap, InfoType.Insight);
          break;
      }
    }


    private void DiagBackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e) { // Call the method that handles common progress changed reports for the threads
      BackgroundWorkerUpdate(sender, e);
    } // end method DiagBackgroundWorker_ProgressChanged

    private void DiagBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
    {
      // Do not access the form's BackgroundWorker reference directly.
      // Instead, use the reference provided by the sender parameter.
      BackgroundWorker bw = sender as BackgroundWorker;

      // Start background worker thread to sniff COM port and display all it sees
      ReceiveDiagData(bw);
    } // end method DiagBackgroundWorker_DoWork

    private void LoBackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
    { // Call the method that handles common progress changed reports for the threads
      BackgroundWorkerUpdate(sender, e);
    } // end method MoBackgroundWorker_ProgressChanged

    private void LoBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
    {
      // Do not access the form's BackgroundWorker reference directly.
      // Instead, use the reference provided by the sender parameter.
      BackgroundWorker bw = sender as BackgroundWorker;

      // Start background worker thread to receive Mobile Originated messages.
      ReceiveLoMessages(bw);
    }

    private void LoRespTimer_Tick(object sender, EventArgs e) {
      /*
       * Event handler for expiration of timer used when a LO response is expected.
       * Expiration is interpretted as failure - so remove remnants of a job that
       * may require more than one LT message to be sent.  Also, notify user of the
       * failure to reply to the LT message and how many other LT messages were discarded.
       */
      string msg;
      string cap;
      MessageBoxButtons btn = MessageBoxButtons.OK;
      MessageBoxIcon icon;
      // LO interface has timed out
      loRespTimer.Enabled = false;                  // so disable the timer until
                                                    // needed again.
      int discarded = DiscardBtMsgJob();            // Discard LT msgs that make up a job
      loRespStatus = LoRespStatus.Null;             // and have no expectations of LO msgs

      if (cfgParams.connected)                      // Not programming but need to let user
      {                                             // know that something timed-out
        msg = "Device response timeout\nto LT " + activeLtMsg + "\n"
            + discarded.ToString() + " LT commands discarded.\n\n"
            + "Verify connection.";
        cap = " Device COMMS ";
        icon = MessageBoxIcon.Exclamation;
        MessageBox.Show(msg, cap, btn, icon);
      }
    } // end method LoRespTimer_Tick

   private void LtBackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
    { // Call the method that handles common progress changed reports for the threads
      BackgroundWorkerUpdate(sender, e);
    } // end method MtBackgroundWorker_ProgressChanged

    private void LtBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
    {
      // Do not access the form's BackgroundWorker reference directly.
      // Instead, use the reference provided by the sender parameter.
      BackgroundWorker bw = sender as BackgroundWorker;

      // Start background worker thread to send Mobile Terminated messages.
      SendLtMessages(bw);         //SendMtMessages(mtBackgroundWorker);
    }

    private int DiscardBtMsgJob() {
      /*
       * Remove LT messages from queue until find one that does not require a LO reply;
       * the 'JobSeparator' messages fit this description (as do ACK/NAK which don't
       * foresee the desire to discard - thus don't discard based on non-zero length).
       * Returns the number of LT messages that were discarded
       * NOTE - WFHWI.msgCnt will be thrown off by the number of messages discarded
       *        but that doesn't really matter
       */
      LtMsgReq btMsgReq;      // copy of message to potentially be discarded
      int startCount;         // how many messages in queue at start of process
      int deltaCount;         // how many messages were removed from queue

      Monitor.Enter(ltMsgQue);
      startCount = ltMsgQue.Count;                        // Get number of messages at start of this
      while (0 < ltMsgQue.Count) {                                                   // Loop, potentially through all messages.
        btMsgReq = (LtMsgReq)ltMsgQue[0];                 // Look at first message remaining in queue
        if (LoRespStatus.Null != btMsgReq.reqLoResp) {   // if message requires response
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


    private void StopMessagingThreads() {
      string diagStr;

      runMsgThreads = false;              // Set flag to stop the messaging threads

      int loops = 300;                    // loop up to 300 times (3-sec) while waiting for
      do                                  // back ground workers to fully terminate
      {                                   // (this is required under Windows 7)
        ltMsgEvent.Set();                 // Set 'send LT message' thread's event and
        loMsgEvent.Set();                 // set 'receive LO message' thread's event,
        Thread.Sleep(10);                 // then give them a chance to see flag and stop.
        Application.DoEvents();           // Process events that may have occurred.
        loops--;
      } while ((0 <= loops) && (loBackgroundWorker.IsBusy || ltBackgroundWorker.IsBusy));

      if (ltBackgroundWorker.IsBusy) {
        diagStr = DateTime.UtcNow.ToString("HH':'mm':'ss.fff");
        diagStr += " BT Messaging Thread not terminating in timely manner.\n";
      }
      if (loBackgroundWorker.IsBusy) {
        diagStr = DateTime.UtcNow.ToString("HH':'mm':'ss.fff");
        diagStr += " BO Messaging Thread not terminating in timely manner.\n";
      }
    } // end method StopMessagingThreads

    private void StopDiagThread() {
      runDiagThread = false;              // Set flag to stop DiagThread

      int loops = 300;                    // loop up to 300 times (3-sec) while waiting for
      do                                  // back ground worker to fully terminate
      {                                   // (this is required under Windows 7)
        diagEvent.Set();                  // Set its event so it stops using COMx;
        Thread.Sleep(10);                 // give it time to see the flag set to stop
        Application.DoEvents();           // then process events that may have occurred
        loops--;
      } while ((0 <= loops) && diagBackgroundWorker.IsBusy);

      if (diagBackgroundWorker.IsBusy) {
        string diagStr;
        diagStr = DateTime.UtcNow.ToString("HH':'mm':'ss.fff");
        diagStr += " Diagnostics Thread not terminating in timely manner.\n";
      }
    } // end method StopDiagThread


    private void PopulateComPortComboBox(ComboBox comPortComboBox) {
      /*
      * Clear all entries in the ComboBox used to display available COM ports
      * Rebuild the list of available/unused COM ports displayed in the ComboBox
      * If applicable, add name of COM port this application is currently using
      */
      if (null == comPortComboBox) {             // make sure were provided instantiated ComboBox
        return;
      }             // but if not there is nothing we should do

      SerialPort testPort = new SerialPort();     // Use this to test availability of specific port
      string[] ports = SerialPort.GetPortNames(); // Get a list of all potentially available ports
      ComboBox portsToAdd = new ComboBox();
      portsToAdd.Sorted = true;                   // Sort list of ports as they are tested & verified

      if (null == comPort) {       // Make sure comPort has been constructed
        comPort = new SerialPort();
      }       // before it is accessed further

      foreach (string portName in ports) { // before testing each and every potential port and adding those we can open successfully
        if ( ! (comPort.IsOpen && portName == comPort.PortName)) { // don't test/close the port we are already using - bad juju
          try {
            testPort.PortName = portName;
            testPort.Open();
            if (testPort.IsOpen) {
              portsToAdd.Items.Add(portName);
            }
          } catch (UnauthorizedAccessException) {
            ; // We could expect to get this exception - no worries, catch & discard
          } catch (Exception) {
            ; // Exception: COM Port Search
          } finally {
            testPort.Close();
          }
        }
      } // end for each COM port name

      if (comPort.IsOpen) {                                             // if we currently have a COM port open
        portsToAdd.Items.Add(comPort.PortName);   // we mustn't forget to add it to the ComboBox
        try {
          using (FileStream fs = File.Open(portCfgFile, FileMode.OpenOrCreate, FileAccess.ReadWrite)) {
            lock (fs) {
              fs.SetLength(0);
            }
          }
          using (StreamWriter sw = new StreamWriter(portCfgFile)) {
            sw.WriteLine(comPort.PortName);
          }
        } catch (Exception) {
          ; // Exception: Saving used COM Port to File Issue
        }
      }

      //Console.WriteLine("Ports: ");
      //foreach (string portName in comPortComboBox.Items) {
      //	Console.Write(portName + " ");
      //}

      { // Now put validated COM port options in displayed ComboBox with desired port first in the list
        string portName = comPortComboBox.Text;
        comPortComboBox.Items.Clear();
        foreach (var item in portsToAdd.Items) {
          comPortComboBox.Items.Add(item);
        }
        int index = comPortComboBox.FindStringExact(portName);
        if ((-1 != index) && (0 != index))
        {                                               // The intent is to have the desired COM port
          comPortComboBox.Items.RemoveAt(index);        // not in the middle of the list
          comPortComboBox.Items.Insert(0, portName);    // but at the top of the list so that
        }                                               // it is checked first when making connection
        comPortComboBox.SelectedIndex = comPortComboBox.FindStringExact(portName);
      }

      return;
    } // end method PopulateComPortComboBox

    private void OpenLightningComPortProperty(string portName) {
      /*
       * Opens the application's "SerialPort comPort" property for use
       * either for hardwire interface or diagnostics connection to device
       */
      if (comPort.IsOpen) {
        comPort.Close();
      }

      comPort.PortName = portName;          // Use the port selected by operator
      comPort.BaudRate = 19200;            // and configure it to match setting
      comPort.DataBits = 8;                 
      comPort.Parity = Parity.None;
      comPort.StopBits = StopBits.One;
      comPort.Handshake = Handshake.None;
      comPort.DiscardNull = false;          // Receiving binary data, which may have NULLs
      comPort.Encoding = Encoding.ASCII;    // Use ASCII because UTF-xx likely to mess up data
      //comPort.Encoding = Encoding.GetEncoding(1252);    // Could use Windows Extended ASCII for'¡' and '¶' detection if required
      comPort.ReadBufferSize = 4096;        // The next three are defaults that probably
      comPort.ReceivedBytesThreshold = 1;   // don't do much for us, but set them
      comPort.WriteBufferSize = 2048;       // rather than depend on Windows OS defaults
      comPort.WriteTimeout = 1000;          // Set the write timeout to 1 second
      comPort.ReadTimeout = 5000;           // Set the read timeout to 5 seconds

      comPort.Open();

      comPort.BreakState = false;           // This really shouldn't be doing anything
    } // end method OpenComPortProperty

    private void OpenBoltComPort(string portName) {
      /*
       * Opens the application's "SerialPort comPort" property for use
       * either for hardwire interface or diagnostics connection to device
       */

      //If the comPort is open close it
      if (comPort.IsOpen) {
        comPort.Close();
      }
      comPort.PortName = portName;          // Use the port selected by operator
      comPort.BaudRate = 115200;            // and configure it to match setting
      comPort.DataBits = 8;                 
      comPort.Parity = Parity.None;
      comPort.StopBits = StopBits.One;
      comPort.Handshake = Handshake.None;
      comPort.DiscardNull = false;          // Receiving binary data, which may have NULLs
      comPort.Encoding = Encoding.ASCII;    // Use ASCII because UTF-xx could mess up data
      comPort.ReadBufferSize = 4096;        // The next three are defaults that probably
      comPort.ReceivedBytesThreshold = 1;   // don't do much for us, but set them
      comPort.WriteBufferSize = 2048;       // rather than depend on Windows OS defaults
      comPort.WriteTimeout = 1000;          // Set the write timeout to 1 second
      comPort.ReadTimeout = 5000;          // Set the read timeout to 5 seconds

      comPort.Open();

      comPort.BreakState = false;           // This really shouldn't be doing anything
    } // end method OpenComPortProperty

    //Stores Serial Communication from BOLT
    private void ParseBoltText(string text) {
      //Need to add locks to all globals
      boltDbg += text;
    }

    //Stores Serial Communication from Lightning
    private void ParseLightningText(string text) {
      //Need to add locks to all globals
      LightningDiag += text;
    }

    //Write to diagnostic port
    private void WriteRichTextBox(RichTextBox rtb, string text) {
      /*
       * Check that the provided RichTextBox has space to append/write the provided text string.
       * It is designed with specific RichTextBoxes in mind, but will handle unknown ones.
       * If the RTB does not have enough space, it saves the existing text to a file before
       * clearing it and starting anew (with a note of where the pre-existing data was saved).
       */
      bool okToWriteText = true;

      if (null == rtb) { // don't do anything to a null RichTextBox
        okToWriteText = false;
        ShowModelessMessageBox("Requested to write a null RichTextBox", "NULL", InfoType.Notification);
      } else if (rtb.MaxLength <= (rtb.TextLength + text.Length)) { // if need to clear what's in RichTextBox to make room for requested text
        rtb.Clear();                  // before clearing it to make room for new text.
      }

      if (okToWriteText) {     // It's okay to write text if rtb is not null
        rtb.AppendText(text);
      }     // and text is smaller than rtb max length

    } // end method WriteRichTextBox

    private void ShowModelessMessageBox(string message, string caption, InfoType importance) {
      /*
       * Display an auto-sizing modeless message box (sorry, can't be called from BackgroundWorkers).
       * The message box will display the message string in a box having the caption and the
       * CloseBox X in its border (no minimize or maximze).
       * If it is requested to be of Notification-importance, it will be a TopMost window with
       * red lettering so the user has a hard time ignoring it. Insight-importance is more subdued
       * in color and is not TopMost (i.e. user can cover it up).
       */
      Form msgBox = new Form();                         // Construct very basic modeless box form
      msgBox.Size = new Size(280, 150);  // Start with some minimum size (may autosize larger)
      msgBox.ShowIcon = false;                          // Don't put an icon in message box caption bar
      msgBox.MinimizeBox = false;                       // Also want neither a minimize button
      msgBox.MaximizeBox = false;                       // nor maximize button in the caption bar.
      msgBox.ShowInTaskbar = false;                     // Don't place an item in Windows Task Bar

      Label msg = new Label();
      msg.Font = new Font("Arial", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0);

      switch (importance) {
        case InfoType.Insight:
          msgBox.TopMost = false;
          msgBox.BackColor = Color.Silver;
          msg.ForeColor = Color.Navy;
          break;

        case InfoType.Notification:
          msgBox.TopMost = true;
          msgBox.BackColor = Color.Bisque;
          msg.ForeColor = Color.Red;
          SystemSounds.Beep.Play();
          break;
      }

      msgBox.Controls.Add(msg);
      msg.Location = new Point(10, 10);
      msg.Text = message + "\n";
      msg.AutoSize = true;

      msgBox.Text = "LBHH: " + caption;
      msgBox.AutoSize = true;

      msgBox.Show(this);
      msgBox.Focus();
    } // end method ShowModelessMessageBox

    private bool ValidateWaveformConfigurationInformation() {
      string msg = "INVALID Configuration Information:\n";
      string cap = " Configure Device Error ";
      bool errFlag = false;
      int placeholderInt = 0;

      // WAVEFORM 0
      //if (w0p2.Text.Length != 1 || !int.TryParse(w0p2.Text, out placeholderInt)) {
      //	msg += " • Slot 0/P2: Must be 1-digit int.\n";
      //	errFlag = true;
      //}
      if (int.TryParse(w0p3.Text, out placeholderInt)) {
        for (int l = w0p3.Text.Length; l < 6; l++) {
          w0p3.Text = "0" + w0p3.Text;
        }
      } else {
        msg += " • Slot 0 Keying ID: Must be six digits (left pad with 0 as required).\n";
        errFlag = true;
      }

      if (int.TryParse(w0p4.Text, out placeholderInt)) {
        for (int l = w0p4.Text.Length; l < 6; l++) {
          w0p4.Text = "0" + w0p4.Text;
        }
      } else {
        msg += " • Slot 0 Transmit ID: Must be six digits (left pad with 0 as required).\n";
        errFlag = true;
      }

      if (w0p5.Text.Length > 0 && Char.IsLetter(w0p5.Text[0]) ) {
        for (int l = w0p5.Text.Length; l < 6; l++) {
          w0p5.Text += "_";
        }
      } else {
        msg += " • Slot 0 WF Name: Must be six-characters (start with alpha, right pad with '_' as required).\n";
        errFlag = true;
      }

      // WAVEFORM 1
      if (int.TryParse(w1p3.Text, out placeholderInt)) {
        for (int l = w1p3.Text.Length; l < 6; l++) {
          w1p3.Text = "0" + w1p3.Text;
        }
      } else {
        msg += " • Slot 1 Keying ID: Must be six digits (left pad with 0 as required).\n";
        errFlag = true;
      }

      if (int.TryParse(w1p4.Text, out placeholderInt)) {
        for (int l = w1p4.Text.Length; l < 6; l++) {
          w1p4.Text = "0" + w1p4.Text;
        }
      } else {
        msg += " • Slot 1 Transmit ID: Must be six digits (left pad with 0 as required).\n";
        errFlag = true;
      }

      if (w1p5.Text.Length > 0 && Char.IsLetter(w1p5.Text[0])) {
        for (int l = w1p5.Text.Length; l < 6; l++) {
          w1p5.Text += "_";
        }
      } else {
        msg += " • Slot 1 WF Name: Must be six-characters (start with alpha, right pad with '_' as required).\n";
        errFlag = true;
      }

      // WAVEFORM 2
      if (int.TryParse(w2p3.Text, out placeholderInt)) {
        for (int l = w2p3.Text.Length; l < 6; l++) {
          w2p3.Text = "0" + w2p3.Text;
        }
      } else {
        msg += " • Slot 2 Keying ID: Must be six digits (left pad with 0 as required).\n";
        errFlag = true;
      }

      if (int.TryParse(w2p4.Text, out placeholderInt)) {
        for (int l = w2p4.Text.Length; l < 6; l++) {
          w2p4.Text = "0" + w2p4.Text;
        }
      } else {
        msg += " • Slot 2 Transmit ID: Must be six digits (left pad with 0 as required).\n";
        errFlag = true;
      }

      if (w2p5.Text.Length > 0 && Char.IsLetter(w2p5.Text[0])) {
        for (int l = w2p5.Text.Length; l < 6; l++) {
          w2p5.Text += "_";
        }
      } else {
        msg += " • Slot 2 WF Name: Must be six-characters (start with alpha, right pad with '_' as required).\n";
        errFlag = true;
      }

      // WAVEFORM 3
      if (int.TryParse(w3p3.Text, out placeholderInt)) {
        for (int l = w3p3.Text.Length; l < 6; l++) {
          w3p3.Text = "0" + w3p3.Text;
        }
      }
      else {
        msg += " • Slot 3 Keying ID: Must be six digits (left pad with 0 as required).\n";
        errFlag = true;
      }

      if (int.TryParse(w3p4.Text, out placeholderInt)) {
        for (int l = w3p4.Text.Length; l < 6; l++) {
          w3p4.Text = "0" + w3p4.Text;
        }
      } else {
        msg += " • Slot 3 Transmit ID: Must be six digits (left pad with 0 as required).\n";
        errFlag = true;
      }

      if (w3p5.Text.Length > 0 && Char.IsLetter(w3p5.Text[0])) {
        for (int l = w3p5.Text.Length; l < 6; l++) {
          w3p5.Text += "_";
        }
      } else {
        msg += " • Slot 3 WF Name: Must be six-characters (start with alpha, right pad with '_' as required).\n";
        errFlag = true;
      }

      // WAVEFORM 4
      if (int.TryParse(w4p3.Text, out placeholderInt)) {
        for (int l = w4p3.Text.Length; l < 6; l++) {
          w4p3.Text = "0" + w4p3.Text;
        }
      } else {
        msg += " • Slot 4 Keying ID: Must be six digits (left pad with 0 as required).\n";
        errFlag = true;
      }

      if (int.TryParse(w4p4.Text, out placeholderInt)) {
        for (int l = w4p4.Text.Length; l < 6; l++) {
          w4p4.Text = "0" + w4p4.Text;
        }
      } else {
        msg += " • Slot 4 Transmit ID: Must be six digits (left pad with 0 as required).\n";
        errFlag = true;
      }

      if (w4p5.Text.Length > 0 && Char.IsLetter(w4p5.Text[0])) {
        for (int l = w4p5.Text.Length; l < 6; l++) {
          w4p5.Text += "_";
        }
      } else {
        msg += " • Slot 4 WF Name: Must be six-characters (start with alpha, right pad with '_' as required).\n";
        errFlag = true;
      }

      if ( ! errFlag) {
        return true;
      } else {
        MessageBoxButtons btn = MessageBoxButtons.OK;
        MessageBoxIcon icon = MessageBoxIcon.Hand;
        MessageBox.Show(msg, cap, btn, icon);
        return false;
      }			
    }

    private void resetConfigTabs() {
      resetActivateTab();
      resetKeyingTab();
      resetWfInfoTab();
    }

    private void resetActivateTab() {
      devInfilUniqueIdBox.Text = null;
      infilKey1.Text = "No Data";
      infilKey2.Text = "No Data";
      infilKey3.Text = "No Data";
      infilKey4.Text = "No Data";
      infilKey5.Text = "No Data";
      noKey0.Checked = true;
    }

    private void resetKeyingTab() {
      waveformtxtbox.Text = null;
      bricknumtxtbox.Text = null;
      txIdTxtBox.Text = null;
    }

    private void resetWfInfoTab() {
      w0p2.Items.Clear();
      w1p2.Items.Clear();
      w2p2.Items.Clear();
      w3p2.Items.Clear();
      w4p2.Items.Clear();

      w0p3.Text = null;
      w1p3.Text = null;
      w2p3.Text = null;
      w3p3.Text = null;
      w4p3.Text = null;

      w0p4.Text = null;
      w1p4.Text = null;
      w2p4.Text = null;
      w3p4.Text = null;
      w4p4.Text = null;

      w0p5.Items.Clear();  w0p5.Text = null;
      w1p5.Items.Clear();  w1p5.Text = null;
      w2p5.Items.Clear();  w2p5.Text = null;
      w3p5.Items.Clear();  w3p5.Text = null;
      w4p5.Items.Clear();  w4p5.Text = null;
    }

    private void resetMiscTab() {
      passwordTxtbox.Text = "";
      groupNameKeyPairString.Text = "";
      sponsorNameKeyPairString.Text = "";
      exportPassword.Text = "";
      showCharactersCheckBox.Checked = true;
      groupNameKeyIndexUpDown.Value = 1;
    }

  }
}

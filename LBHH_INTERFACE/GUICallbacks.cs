using System;
using System.Xml;
using System.IO;
using System.Media;                           // Needed for SystemSounds.Beep
using System.Windows;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; //Needed for operation with excel
using Microsoft.VisualBasic;                  //Needed for Interaction.InputBox
using System.Collections.Generic;
using System.Drawing;
using System.IO.Ports;
using System.Threading;
using System.Timers;
using System.Text.RegularExpressions;


namespace LBHH_Red {
  public partial class LBHH_HWI {

    //## REGION ################################################################
    #region Not_A_Tab

    private void ComPortComboBox_DropDown(object sender, EventArgs e)
    {
      PopulateComPortComboBox(comPortComboBox);
    }

    // Name: ConnectButton_Click
    // Arguments: sender
    //						e
    // Description: Connects the GUI to the Lightning and attempts to poll available/selected
    //              Infil keys and Exfil waveforms.
    private void ConnectButton_Click(object sender, EventArgs e)
    {
      connectButton.Enabled = false;
      busyBar.Visible = true;          // Indicate activity to user

      if ( ! cfgParams.connected)
      {
        connectButton.Text = "Disconnect";
        comPortComboBox.Enabled = false;
        gpiParams.InitToHardCodedDefaults();
        cfgParams.InitToHardCodedDefaults();
        // If there is a selected COM Port Check First when Auto-Connecting
        // If it fails try other ComPorts in list
        string diagComPort = (diagBackgroundWorker.IsBusy)  // if collecting device diag data
                           ? comPort.PortName               // remember port being used
                           : "";

        //Populate the Com Port ComboBox
        PopulateComPortComboBox(comPortComboBox);

        if (0 == comPortComboBox.Items.Count)
        {
          connectButton.Text = "Connect";
          comPortComboBox.Enabled = true;
          string msg = "No COM Port detected.";
          string cap = " COM Port Search ";
          MessageBoxButtons btn = MessageBoxButtons.OK;
          MessageBoxIcon icon = MessageBoxIcon.Asterisk;
          MessageBox.Show(msg, cap, btn, icon);
        }
        else
        {
          try
          {
            string portName;

            while (0 != comPortComboBox.Items.Count)
            {
              comPortComboBox.SelectedIndex = 0;
              portName = comPortComboBox.SelectedItem.ToString();
              comPortComboBox.Items.RemoveAt(0);
              //Call function that tells RX thread to listen for device configuration

              if (MakeHwiPrgmConnection(portName))
              { // Configuration programming connection made
                //If sucessful unlock interfaces
                comPortComboBox.Items.Clear();    // break looping through candiate COM ports
                PopulateComPortComboBox(comPortComboBox);
                comPortComboBox.SelectedItem = portName;
                this.Text = ConnectedTitle + portName; // Set window/form title bar to its default
                                                       //Load parameters from cfgParam
                infilKeysGrpBox.Enabled = true; //we connected so enable group boxes
                boltSetupGrpBox.Enabled = true; //we connected so enable group boxes

                ltngMsgStartLbl.Enabled = true;
                cmdTxtBox.Enabled = true;
                ltngMsgStopLbl.Enabled = true;
                sendCmdBtn.Enabled = true;
                connectBoltBtn.Enabled = true;

                lightningWaveformConfigurationGroup.Enabled = true; // we connected so enable group boxes
                antiTamperBox.Enabled = true;

                gpiParams.LoadOptions(ref cfgParams);
                SetGpiFromProgParams(ref gpiParams);       // set GPI to last known programmed values
                break;
              }
              else if (0 == comPortComboBox.Items.Count)
              { // Failed to make connection to any candidate COMx
                DisconnectDevice();
                throw new Exception("No device auto-detected");
              }
              busyBar.Refresh();
            } // end while COM ports to test
          } //end try
          catch (Exception ex)
          {
            connectButton.Text = "Connect";
            comPortComboBox.Enabled = true;
            this.Text = DisconnectedTitle;                // Set window/form title bar to its default
            StopMessagingThreads();                       // Ensure shutdown of messaging threads
            string msg = "Auto-Connection FAILED\n" + ex.Message;
            string cap = " Device Connection Error ";
            MessageBoxButtons btn = MessageBoxButtons.OK;
            MessageBoxIcon icon = MessageBoxIcon.Hand;

            if (0 != diagComPort.Length)
            {                // When device diag port open at time
              if (diagComPort != comPort.PortName)        // user clicked 'Device->Auto Detect'
              {                                           // but diag port not last one checked,
                comPort.Close();                          // close whatever port may be open now
                OpenLightningComPortProperty(diagComPort);// and switch to what had been diag port.
              }                                           // Anytime had been listening to diag
                                                          // SetGpiFromProgParams(ref cfgParams);       // set GPI to last known programmed values
              DisconnectDevice();                         // and act as if just disconnected HWI
            }                                             // (i.e. listen for device diag data)
            else if (comPort.IsOpen)                      // When device diag port not previously open
            {                                             // (we are here because we failed to connect)
              comPort.Close();                            // close COM port in case another app needs it
            }

            MessageBox.Show(msg, cap, btn, icon);         // Tell user of failure to auto-detect
            PopulateComPortComboBox(comPortComboBox);
          } //end catch exception

        } // end if no ports found else

        if (numObfKeys != obfKeys.keys.Count)       // While attempting to connect to LBHH and
        {                                           // only when haven't already read infil keys
          tabControl.Enabled = false;               // disable basically everything since Excel is slow
          if ( ! ReadInfilKeyFile())                // and attempt to do read key file now
          {                                         // If fail reading keys then
            programInfilKeysButton.Enabled = false; // never enable these buttons regardless of
            loadObfKeysBtn.Enabled = false;         // 'Infil Keys' group box status
            string cap = "INFIL KEYS DISABLED";
            string msg = "NOTICE: all 'Infil Keys' operations are disabled until\n"
                       + "a valid key file is installed and program is restarted.\n\n"
                       + "Ignore this if you do not need to load infil keys.\n\n"
                       + "!!! NEVER INSTALL KEY FILE ON A NETWORKED COMPUTER !!!";
            ShowModelessMessageBox(msg,cap,InfoType.Notification);
          }
          tabControl.Enabled = true;                // Re-enable all that should be
        }

      }
      else if (cfgParams.connected)
      {
        infilKeysGrpBox.Enabled = false; // we disconnected so disable group boxes
        boltSetupGrpBox.Enabled = false; // we disconnected so disable group boxes

        ltngMsgStartLbl.Enabled = false;
        cmdTxtBox.Enabled = false;
        ltngMsgStopLbl.Enabled = false;
        sendCmdBtn.Enabled = false;
        connectBoltBtn.Enabled = false;

        lightningWaveformConfigurationGroup.Enabled = false; // we disconnected so disable group boxes
        antiTamperBox.Enabled = false;
        GetProgParamsFromGpi(ref gpiParams);  // Get parameters from display

        resetConfigTabs();
        this.Text = DisconnectedTitle;        // Set window/form title bar to its default
        connectButton.Text = "Connect";
        comPortComboBox.Enabled = true;

        StopMessagingThreads();               // Ensure shutdown of messaging threads
        Monitor.Enter(comPort);               // wait for and lock down access to the COM port
        {
          try
          {
            if ((comPort != null) && comPort.IsOpen)
            {                                 // When the COM port is open
              comPort.DiscardOutBuffer();     // discard anything getting ready to send
              comPort.DiscardInBuffer();      // and anything recently received by COMx
            }
          }
          catch (Exception)
          {
            ;
          }
        }
        Monitor.Exit(comPort);                // Ensure we release the lock on accessing COM port
        comPort.Dispose();
        comPort = new SerialPort();           // Ensure comPort is left neither connected nor NULL
        infilgrpBox.Enabled = false;
        configureInfilButton.Enabled = false;
        lightningStatus.Image = minusImg;
        gpiParams.InitToHardCodedDefaults();
        cfgParams.InitToHardCodedDefaults();
      }

      busyBar.Visible = false;
      connectButton.Enabled = true;
    } // end ConnectButton_Click

    private void HelpGeneralToolStripMenuItem_Click(object sender, EventArgs e)
    {
      string msg = "\nWARNING: do NOT twist/unscrew cable attached to LBHH as this can cause SEVERE damage to LBHH or cable !!!\n\n"
                 + "Step 1: Preparation...\n"
                 + "        Connect special 'LBHH External Comms' cable to USB port of PC running this program\n"
                 + "        Press cable into LBHH with marker on the cable (typically a dot) aligned with marker on LBHH port (also a dot)\n"
                 + "        Ensure the LBHH is powered on and is displaying its 'Home' screen\n\n"
                 + "Step 2: Initiate comms with LBHH by clicking the 'Connect' button; program automatically detects the LBHH\n\n"
                 + "Step 3: KEYING tab...\n"
                 + "        in 'Infil Keys' click 'Load Obfuscation Keys' button\n"
                 + "        in 'Bolt Setup', for _every_ Bolt slot with a waveform, click 'Load XML' and load appropriate XML file\n"
                 + "        in 'Bolt Setup' click 'Load TEK Keys and Configure BOLT' and follow direction given by program\n"
                 + "        in 'Infil Keys', for every desired optional group key (1,2,3,4,5), select desired key within its drop-down\n"
                 + "        in 'Infil Keys' verify correct 'Unique ID' then click 'Program Infil Keys'\n\n"
                 + "Step 4: WF Info tab...\n"
                 + "        click 'Query LBHH' then set Slot 0's WF Name to 'U_TEST'\n"
                 + "        for each slot to be enabled, select proper 'Family' and fill in rest of the slot's waveform info\n"
                 + "        for each slot with a waveform, but to be disabled, select 'NONE' for 'Family' and fill in rest of slot's real info\n"
                 + "        for each slot without a wavefore, click it's 'Clear' button\n"
                 + "        click 'Enable for Admin'\n"
                 + "        * This tab can also be used to discover identifying info about the BOLT and all waveforms in a connected LBHH\n\n"
                 + "Step 5: Activate Infil tab...\n"
                 + "        select the Infil Group Key to be used (or 'No key')\n"
                 + "        click 'Configure Infil' (note this does not affect Unique ID)\n"
                 + "        * This tab can also be used to discover the names of all infil keys loaded in a connected LBHH\n\n"
                 + "Step 6: Use LBHH keypad to access its Admin menu, enter an Exfil Choice, and follow LBHH's prompts\n\n"
                 + "Step 7: Click 'Disconnect' then pull cable from LBHH.  Conduct LBHH operational acceptance testing\n\n"
                 + "Step 8: Misc tab...\n"
                 + "        Before LBHH delivery connect LBHH and click 'Anti-Tamper ON'\n"
                 + "        * Before LBHH maintenance connect LBHH and click 'Anti-Tamper OFF'\n"
                 + "        This tab is used to create new associations between a Group name and Key value\n"
                 + "        This tab is also used to 'Export' a sponsor's group keys to a password protected file\n\n"
                 + "NOTE - This program will connect to the first LBHH/Lightning device it detects.  If multiple devices attached to PC,\n"
                 + "use the 'Port Selection' drop-down to select desired COM port before clicking 'Connect'.\n";
      string cap = "Red Tool Instructions and Information";
      ShowModelessMessageBox(msg, cap, InfoType.Insight);
    }

    private void HelpAboutToolStripMenuItem_Click(object sender, EventArgs e)
    {
      // the menu item 'Help -> About' has just been clicked
      string msg, cap;
      //Attribute appAttribute = Attribute.GetCustomAttribute(clsType.Module, typeof(GuidAttribute));
      //    "#########1#########2#########3#########4#########5#########6#########7#########8"
      msg = DisconnectedTitle
          + "\nVersion 2.1.7\n"
          + "Property of U.S. Government\n"
          + "Developed by NSWCDD-H15 Pharos\n\n"
          + "DISTRIBUTION F. Further dissemination only as directed by\n\n"
          + "Naval Surface Warfare Center, Dahlgren Division\n"
          + "H15 Pharos, Evan A. Aanerud\n"
          + "DAHLGREN VA 22448\n"
          + "or higher DoD authority;\n\n"
          + "Direct Military Support and Critical Technology.\n\n"
          + "Any misuse or unauthorized distribution is strictly prohibited\n"
          + "and can result in both civil and criminal penalties.\n";
      cap = "About LBHH Tool";
      MessageBoxButtons btn = MessageBoxButtons.OK;
      MessageBoxIcon icon = MessageBoxIcon.Asterisk;
      MessageBox.Show(msg, cap, btn, icon);
    }

    #endregion Not_A_Tab

    //## REGION ################################################################
    #region Keying Tab

    // Name: LoadObfuscationKeysBtn_Click
    // Arguments: sender
    //            e
    // Description: Handles when user clicks the 'Load Obfuscation Keys' button
    //              to send the obfuscation table to Lightning
    private void LoadObfuscationKeysBtn_Click(object sender, EventArgs e)
    {
      loadObfKeysBtn.Enabled = false;
      busyBar.Visible = true;          // Indicate activity to user

      for (int obfIdx = 0; (numObfKeys > obfIdx) && cfgParams.connected; obfIdx++)
      {
        obfKeys.keyIdx = obfIdx;
        QueueLtMessage(LightningMsgType.lObfKey);
        Thread.Sleep(50);
        while (loRespTimer.Enabled)
        {
          Thread.Sleep(10);
          Application.DoEvents();                 
        }
        busyBar.Refresh();
      }
      if (cfgParams.connected)
      { // If still connected then Lightning did not NAK any load command
        string cap = "SUCCESS";
        string msg = "SUCCESS: All Obfuscation Keys Loaded.";
        ShowModelessMessageBox(msg, cap, InfoType.Insight);
      }
      else
      { // LoMsgThread will clear .connected when there is an issue
        string cap = "ERROR";
        string msg = "ERROR: Not all Obfuscation Keys loaded.\n\nPlease Retry Configuration";
        ShowModelessMessageBox(msg, cap, InfoType.Notification);
        DisconnectDevice();
      }

      busyBar.Visible = false;
      loadObfKeysBtn.Enabled = true;
    } // end LoadObfuscationKeysBtn_Click

    //Name: keyCombo1_SelectedIndexChanged
    //Arguments: sender
    //           e
    //Description: Handles when user selects the Group Key for infil slot 1
    private void keyCombo1_SelectedIndexChanged(object sender, EventArgs e)
    {
      if (keyCombo1.Text == "None")
      {
        gpiParams.cfgInfil1Name = "1NoKey";
        gpiParams.cfgInfil1index = null;
        gpiParams.cfgInfil1grpindx = -1;
        key1TextBox.Text = gpiParams.cfgInfil1index;
        key1TextBox.Refresh();
      }
      else
      {
        for (int i = 0; i < infilKeys.GroupKeys.Count; i++)
        {
          if (infilKeys.GroupKeys[i].GroupName == keyCombo1.Text)
          {
            gpiParams.cfgInfil1Name = infilKeys.GroupKeys[i].GroupName;
            gpiParams.cfgInfil1index = infilKeys.GroupKeys[i].index;
            gpiParams.cfgInfil1grpindx = i;
          }
        }
      }
      key1TextBox.Text = gpiParams.cfgInfil1index;
      key1TextBox.Refresh();
    }

    //Name: keyCombo2_SelectedIndexChanged
    //Arguments: sender
    //           e
    //Description: Handles when user selects the Group Key for infil slot 2
    private void keyCombo2_SelectedIndexChanged(object sender, EventArgs e)
    {
      if (keyCombo2.Text == "None")
      {
        gpiParams.cfgInfil2Name = "2NoKey";
        gpiParams.cfgInfil2index = null;
        gpiParams.cfgInfil2grpindx = -1;
        key2TextBox.Text = gpiParams.cfgInfil2index;
        key2TextBox.Refresh();
      }
      else
      {
        for (int i = 0; i < infilKeys.GroupKeys.Count; i++)
        {
          if (infilKeys.GroupKeys[i].GroupName == keyCombo2.Text)
          {
            gpiParams.cfgInfil2Name = infilKeys.GroupKeys[i].GroupName;
            gpiParams.cfgInfil2index = infilKeys.GroupKeys[i].index;
            gpiParams.cfgInfil2grpindx = i;
          }
        }
      }
      key2TextBox.Text = gpiParams.cfgInfil2index;
      key2TextBox.Refresh();
    }

    //Name: keyCombo3_SelectedIndexChanged
    //Arguments: sender
    //           e
    //Description: Handles when user selects the Group Key for infil slot 3 
    private void keyCombo3_SelectedIndexChanged(object sender, EventArgs e)
    {
      if (keyCombo3.Text == "None")
      {
        gpiParams.cfgInfil3Name = "3NoKey";
        gpiParams.cfgInfil3index = null;
        gpiParams.cfgInfil3grpindx = -1;
        key3TextBox.Text = gpiParams.cfgInfil3index;
        key3TextBox.Refresh();
      }
      else
      {
        for (int i = 0; i < infilKeys.GroupKeys.Count; i++)
        {
          if (infilKeys.GroupKeys[i].GroupName == keyCombo3.Text)
          {
            gpiParams.cfgInfil3Name = infilKeys.GroupKeys[i].GroupName;
            gpiParams.cfgInfil3index = infilKeys.GroupKeys[i].index;
            gpiParams.cfgInfil3grpindx = i;
          }
        }
      }
      key3TextBox.Text = gpiParams.cfgInfil3index;
      key3TextBox.Refresh();
    }

    //Name: keyCombo4_SelectedIndexChanged
    //Arguments: sender
    //           e
    //Description: Handles when user selects the Group Key for infil slot 4
    private void keyCombo4_SelectedIndexChanged(object sender, EventArgs e)
    {
      if (keyCombo4.Text == "None")
      {
        gpiParams.cfgInfil4Name = "4NoKey";
        gpiParams.cfgInfil4index = null;
        gpiParams.cfgInfil4grpindx = -1;
        key4TextBox.Text = gpiParams.cfgInfil4index;
        key4TextBox.Refresh();
      }
      else
      {
        for (int i = 0; i < infilKeys.GroupKeys.Count; i++)
        {
          if (infilKeys.GroupKeys[i].GroupName == keyCombo4.Text)
          {
            gpiParams.cfgInfil4Name = infilKeys.GroupKeys[i].GroupName;
            gpiParams.cfgInfil4index = infilKeys.GroupKeys[i].index;
            gpiParams.cfgInfil4grpindx = i;
          }
        }
      }
      key4TextBox.Text = gpiParams.cfgInfil4index;
      key4TextBox.Refresh();
    }

    //Name: keyCombo5_SelectedIndexChanged
    //Arguments: sender
    //           e
    //Description: Handles when user selects the Group Key for infil slot 5
    private void keyCombo5_SelectedIndexChanged(object sender, EventArgs e)
    {
      if (keyCombo5.Text == "None")
      {
        gpiParams.cfgInfil5Name = "5NoKey";
        gpiParams.cfgInfil5index = null;
        gpiParams.cfgInfil5grpindx = -1;
        key5TextBox.Text = gpiParams.cfgInfil5index;
        key5TextBox.Refresh();
      }
      else
      {
        for (int i = 0; i < infilKeys.GroupKeys.Count; i++)
        {
          if (infilKeys.GroupKeys[i].GroupName == keyCombo5.Text)
          {
            gpiParams.cfgInfil5Name = infilKeys.GroupKeys[i].GroupName;
            gpiParams.cfgInfil5index = infilKeys.GroupKeys[i].index;
            gpiParams.cfgInfil5grpindx = i;
          }
        }
      }
      key5TextBox.Text = gpiParams.cfgInfil5index;
      key5TextBox.Refresh();
    }

    // Name: ProgramInfilKeysButton_Click
    // Arguments: sender
    //            e
    // Description: Loads infil keys into the lightning
    private void ProgramInfilKeysButton_Click(object sender, EventArgs e)
    {
      programInfilKeysButton.Enabled = false;
      busyBar.Visible = true;          // Indicate activity to user

      bool uniqueKeyPulled = true;
      if (uniqueIDupdown.Value != infilKeys.deviceID)                               {
        uniqueKeyPulled = GetUniqueIdKey((int)uniqueIDupdown.Value, ref infilKeys); }

      //Update string representing the device ID
      if (uniqueIDupdown.Value < 10)                                    {
        gpiParams.deviceId = "ID000" + uniqueIDupdown.Value.ToString(); }
      else if (uniqueIDupdown.Value < 100)                              {
        gpiParams.deviceId = "ID00" + uniqueIDupdown.Value.ToString();  }
      else if (uniqueIDupdown.Value < 1000)                             {
        gpiParams.deviceId = "ID0" + uniqueIDupdown.Value.ToString();   }
      else if (uniqueIDupdown.Value < 10000)                            {
        gpiParams.deviceId = "ID" + uniqueIDupdown.Value.ToString();    }
      else if (uniqueIDupdown.Value < 100000)                           {
        gpiParams.deviceId = "I" + uniqueIDupdown.Value.ToString();     }
      else                                                                                      {
        ShowModelessMessageBox("Unique ID impossibly large!", "ERROR", InfoType.Notification);  }

      //Only run if Keys have been retrieved and ID>0
      if (cfgParams.connected && infilKeys.retrievedKeys && uniqueKeyPulled && uniqueIDupdown.Value > 0)
      { //Reload new ID Key if value in up down is different than the pulled device ID
        //Wipe all keys and start fresh
        QueueLtMessage(LightningMsgType.dInfilKeys);
        Thread.Sleep(50);
        while (loRespTimer.Enabled)  // Timer disabled on error (which sets cfgParams.connectd = false) or cmd completion
        {
          Thread.Sleep(10);
          Application.DoEvents();                 
        }
        Thread.Sleep(50);             // Give time for LoMsgThread to post Ltng ACK/NAK response
        Application.DoEvents();                 

        //Make GPI Values Match programming values
        gpiParams.InfilKey1 = gpiParams.cfgInfil1Name;
        gpiParams.InfilKey2 = gpiParams.cfgInfil2Name;
        gpiParams.InfilKey3 = gpiParams.cfgInfil3Name;
        gpiParams.InfilKey4 = gpiParams.cfgInfil4Name;
        gpiParams.InfilKey5 = gpiParams.cfgInfil5Name;
        //Load 6 keys: 1 Unique infil and 5 Group infil
        for (gpiParams.keyLoad = 0; gpiParams.keyLoad < 6; gpiParams.keyLoad++)
        {
          busyBar.Refresh();
          QueueLtMessage(LightningMsgType.lInfilKey);
          Thread.Sleep(50);
          while (loRespTimer.Enabled) // Timer disabled on error (which sets cfgParams.connectd = false) or cmd completion
          {
            Thread.Sleep(10);
            Application.DoEvents();                 
          }
          Thread.Sleep(50);           // Give LBHH a little breather between commands
          Application.DoEvents();     // and let other threads do whatever since we paused
        }
        //Once done check to ensure that the gpi infil key names match the key names on the Bolt
        QueueLtMessage(LightningMsgType.qKeyNames);
        Thread.Sleep(50);
        while (loRespTimer.Enabled)   // Timer disabled on error (which sets cfgParams.connectd = false) or cmd completion
        {
          Thread.Sleep(10);
          Application.DoEvents();                 
        }
        Thread.Sleep(10);             // Give a very brief time for other threads to do whatever unimagined shit
        Application.DoEvents();       // because randomly cfgParams update seems to lag when check .InfilEquals below

        if (gpiParams.InfilEquals(cfgParams))
        { // When names match configuration was a success
          string cap = "SUCCESS";
          string msg = "SUCCESS: All Infil Keys Loaded.";
          ShowModelessMessageBox(msg, cap, InfoType.Insight);
          SetGpiFromProgParams(ref gpiParams);
        }
        else
        { //Device unsuccessfully Loaded Waveforms - so Disconnect prompting user to try again
          string cap = "ERROR";
          string msg = "ERROR: Not all Infil Keys Loaded. \nPlease Retry Configuration";
          ShowModelessMessageBox(msg, cap, InfoType.Notification);
          DisconnectDevice();
        }
      }
      else
      { //Prompt user to select a relevant device ID
        string cap = "ERROR";
        string msg = "ERROR: The Unique ID should not be zero (0).\nPlease set Unique ID to Unit's AME ID.";
        ShowModelessMessageBox(msg, cap, InfoType.Notification);
        //Report error occurred reconnect device
      }

      busyBar.Visible = false;
      programInfilKeysButton.Enabled = true;
    } // end ProgramInfilKeysButton_Click

    //Name: loadXmlslot1_Click
    //Arguments: sender
    //           e
    //Description: Load command XML into the 1st waveform slot
    private void loadXmlslot1_Click(object sender, EventArgs e)
    {
      if (waveforms.hwiParams[0].appName == "" || waveforms.hwiParams[0].appName == null)
      {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.RestoreDirectory = true;
        openFileDialog.InitialDirectory = cfgDir;
        openFileDialog.Filter = "Waveform Config Files (*.xml)|*.xml";
        // Show the Dialog. If user selected a file and clicked OK, use it.
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
          waveforms.LoadCommandsXml(openFileDialog.FileName, 0);
          if (waveforms.hwiParams[0].appName == "" || waveforms.hwiParams[0].appName == null) {
            waveforms.hwiParams[0].appName = "No Name";                                       }
          slotTextbox1.Text = waveforms.hwiParams[0].appName;
          loadXmlslot1.Text = "Remove";
        }
        openFileDialog.Dispose();
      }
      else
      {
        waveforms.hwiParams[0] = new WaveformHWIParams();
        waveforms.hwiParams[0].appName = "";
        slotTextbox1.Text = "";
        loadXmlslot1.Text = "Load XML";
      }
    }

    //Name: loadXmlslot2_Click
    //Arguments: sender
    //           e
    //Description: Load command XML into the 2nd waveform slot
    private void loadXmlslot2_Click(object sender, EventArgs e)
    {
      if (waveforms.hwiParams[1].appName == "" || waveforms.hwiParams[1].appName == null)
      {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.RestoreDirectory = true;
        openFileDialog.InitialDirectory = cfgDir;
        openFileDialog.Filter = "Waveform Config Files (*.xml)|*.xml";
        // Show the Dialog. If user selected a file and clicked OK, use it.
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
          waveforms.LoadCommandsXml(openFileDialog.FileName, 1);
          if (waveforms.hwiParams[1].appName == "" || waveforms.hwiParams[1].appName == null) {
            waveforms.hwiParams[1].appName = "No Name";                                       }
          slotTextbox2.Text = waveforms.hwiParams[1].appName;
          loadXmlslot2.Text = "Remove";
        }
        openFileDialog.Dispose();
      }
      else
      {
        waveforms.hwiParams[1] = new WaveformHWIParams();
        waveforms.hwiParams[1].appName = "";
        slotTextbox2.Text = "";
        loadXmlslot2.Text = "Load XML";
      }
    }

    //Name: loadXmlslot3_Click
    //Arguments: sender
    //           e
    //Description: Load command XML into the 3rd waveform slot
    private void loadXmlslot3_Click(object sender, EventArgs e)
    {
      if (waveforms.hwiParams[2].appName == "" || waveforms.hwiParams[2].appName == null)
      {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.RestoreDirectory = true;
        openFileDialog.InitialDirectory = cfgDir;
        openFileDialog.Filter = "Waveform Config Files (*.xml)|*.xml";
        // Show the Dialog. If user selected a file and clicked OK, use it.
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
          waveforms.LoadCommandsXml(openFileDialog.FileName, 2);
          if (waveforms.hwiParams[2].appName == "" || waveforms.hwiParams[2].appName == null) {
            waveforms.hwiParams[2].appName = "No Name";                                       }
          slotTextbox3.Text = waveforms.hwiParams[2].appName;
          loadXmlslot3.Text = "Remove";
        }
        openFileDialog.Dispose();
      }
      else
      {
        waveforms.hwiParams[2] = new WaveformHWIParams();
        waveforms.hwiParams[2].appName = "";
        slotTextbox3.Text = "";
        loadXmlslot3.Text = "Load XML";
      }
    }

    //Name: loadXmlslot4_Click
    //Arguments: sender
    //           e
    //Description: Load command XML into the 4th waveform slot
    private void loadXmlslot4_Click(object sender, EventArgs e)
    {
      if (waveforms.hwiParams[3].appName == "" || waveforms.hwiParams[3].appName == null)
      {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.RestoreDirectory = true;
        openFileDialog.InitialDirectory = cfgDir;
        openFileDialog.Filter = "Waveform Config Files (*.xml)|*.xml";
        // Show the Dialog. If user selected a file and clicked OK, use it.
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
          waveforms.LoadCommandsXml(openFileDialog.FileName, 3);
          if (waveforms.hwiParams[3].appName == "" || waveforms.hwiParams[3].appName == null) {
            waveforms.hwiParams[3].appName = "No Name";                                       }
          slotTextbox4.Text = waveforms.hwiParams[3].appName;
          loadXmlslot4.Text = "Remove";
        }
        openFileDialog.Dispose();
      }
      else
      {
        waveforms.hwiParams[3] = new WaveformHWIParams();
        waveforms.hwiParams[3].appName = "";
        slotTextbox4.Text = "";
        loadXmlslot4.Text = "Load XML";
      }
    }

    //Name: KeyAndCfgBoltBtn_Click
    //Arguments: object sender
    //           EventArgs e
    //Description: Used to configure the Bolt Hand Held upon being clicked it processes configuration XML files
    //and loads the data into the Bolt
    private void KeyAndCfgBoltBtn_Click(object sender, EventArgs e)
    {
      keyAndCfgBoltBtn.Enabled = false;
      busyBar.Visible = true;          // Indicate activity to user
      int numConfigwfrms = 0; //Number of waveforms cycled through
                              //There will be a while statement here that will cycle through available option
      ConnectToBoltDebugPort(); //Start Bolt Communication and stop Lightning communication
      string msg;
      string cap;
      MessageBoxButtons btn;
      MessageBoxIcon icon;

      //Part one of the connection process load tek
      GetWfmidFromBolt(); //Loads the current waveform ID            
      bool lightningWaveformChange = false;
      //cyle through all loaded waveforms and load parameters
      while (numConfigwfrms < MAX_SLOT) // < MAX_SLOT works because here RED never accesses SLOT 0, otherwise need <=
      {
        busyBar.Refresh();
        int wfmid = waveforms.wfidx + 1; // we skip WFMID 0 so 0-base array index always 1-less than WFMID
        if (waveforms.hwiParams[waveforms.wfidx].appName != "")
        { // when have an XML loaded for the BOLT's active WFMID
          waveformtxtbox.Text = wfmid.ToString();
          //If this is not the first waveform to configure send the waveform change command
          if (numConfigwfrms != 0)
          { // When not still doing first WF config, need to change WF using either Lightning or BOLT Debug command
            if ( ! lightningWaveformChange)
            { //Load next waveform in queue
              if ( ! SendBoltWfchangeCmd(wfmid)) //BOLT operational WFMIDs are from 1 onwards; array wfidx is 0-based
              {
                msg = "ERROR: Slot " + wfmid.ToString() + " has not been loaded.\nCheck BOLT revision and WF load-out.\nCANCELLING CONFIGURATION PROCESS";
                cap = "BOLT SLOT NOT LOADED (wfchange)";
                btn = MessageBoxButtons.OK;
                icon = MessageBoxIcon.Asterisk;
                MessageBox.Show(msg, cap, btn, icon);
                break;
              }
            }
            else // if (lightningWaveformChange) -- i.e. use Ltng C.SWF, instead of BOLT wfchange, command
            {
              msg = "Please cycle LBHH power off-ON and\nTHEN promptly press OK";
              cap = "RESTART DEVICE";
              btn = MessageBoxButtons.OK;
              icon = MessageBoxIcon.Asterisk;
              MessageBox.Show(msg, cap, btn, icon);
              Application.DoEvents();
              Thread.Sleep(500);      // Sleep 1/2 second while LBHH wakes up from power cycle
              //Reconnect to the Bolt
              if ( ! ConnectToLightning(comPort.PortName))
              { //When there was an error establishing Lightning comms - inform user and stop programming process
                msg = "ERROR: Device restart check failed (ConnectToLightning)\nCANCELLING CONFIGURATION PROCESS";
                cap = "RESTART CHECK FAILURE";
                btn = MessageBoxButtons.OK;
                icon = MessageBoxIcon.Asterisk;
                MessageBox.Show(msg, cap, btn, icon);
                break;
              }
              if ( ! MakeLightningSwitchWaveform(wfmid, comPort.PortName))
              { // When there was an error switching waveforms - inform user and stop programming process
                msg = "ERROR: Failed switching to Slot " + wfmid.ToString() + "\nSlot has not been loaded.\nCheck BOLT revision and WF load-out.\nCANCELLING CONFIGURATION PROCESS";
                cap = "BOLT SLOT NOT LOADED (C.SWF #)";
                btn = MessageBoxButtons.OK;
                icon = MessageBoxIcon.Asterisk;
                MessageBox.Show(msg, cap, btn, icon);
                break;
              }
              const int _50ms = 50;                   // use 50ms short sleeps
// TODO - this timing paradigm is sketchy and may not work for future waveforms as it can be iffy as is
// TODO - do like Yellow tool
// TODO - add code to check BOLT Debug output for "WFEnablesFlags: 0xHH" where 0xHH is hex bitfield and then only allow changing to enabled WFs
              int countDown = (30000/_50ms);          // for a total of 30 seconds which is required
              while (0 < countDown)                   // for BOLT to get, process, and complete a
              {                                       // change of waveform (completely reboots BOLT)
                loMsgEvent.Set();                     // Signal 'send LO message' thread to run
                ltMsgEvent.Set();                     // Signal 'send LT message' thread to run
                Thread.Sleep(_50ms);                  // use short sleeps so progressBar scrolls.
                Application.DoEvents();               // Each time the 'while loop' is executed
                countDown--;                          // decrement the countdown until reach 0
              }
              ConnectToBoltDebugPort();           // Switch COMx to the BOLT Debug Port now that BOLT should be ready
              if ( ! GetWfmidFromBolt())
              { //There was an error connecting to BOLT - inform user and break programming process
                msg = "ERROR: Device restart check failed (GetWfmidFromBolt)\nCANCELLING CONFIGURATION PROCESS";
                cap = "RESTART CHECK FAILURE";
                btn = MessageBoxButtons.OK;
                icon = MessageBoxIcon.Asterisk;
                MessageBox.Show(msg, cap, btn, icon);
                break;
              }
            }
          } // end need to change BOLT's active wf (after doing first config)

          
//        if (waveforms.hwiParams[waveforms.wfidx].password != "")
          {
            { // First must always load BOLT waveform TEK(s)    // TODO make decision based on <NeedsKey>
              if ( ! HandlePassword())
              { // When there was an error loading the password - inform user and stop programming process
                msg = "ERROR: Incorrect PASSWORD PROMPT or PASSWORD\n"
                    + "provided in XML file for BOLT waveform.\n"
                    + "Use XML appropriate for waveform loadout.\n"
                    + "CANCELLING BOLT CONFIGURATION PROCESS";
                cap = "INCORRECT PASSWORD";
                btn = MessageBoxButtons.OK;
                icon = MessageBoxIcon.Asterisk;
                MessageBox.Show(msg, cap, btn, icon);
                break;
              }
              if ( ! GetUnitID())
              { // When there was an error finding the Unit ID - inform user and stop programming process
                msg = "ERROR: Transmit ID not found.\nPerhaps incorrect XML file selected";
                cap = "NO USER ID";
                btn = MessageBoxButtons.OK;
                icon = MessageBoxIcon.Asterisk;
                MessageBox.Show(msg, cap, btn, icon);
                break;
              }
              if ( ! GetAndLoadTek())
              { // When there was an error loading the TEK key - inform user and stop programming process
                msg = "ERROR: Waveform's TEK(s) could not be loaded.\nFile not found or BOLT unhappy.\nCANCELLING CONFIGURATION PROCESS";
                cap = "ISSUE LOADING TEK KEY";
                btn = MessageBoxButtons.OK;
                icon = MessageBoxIcon.Asterisk;
                MessageBox.Show(msg, cap, btn, icon);
                break;
              }
            }

            // Second, if pertinent, send Commands from XML, which requires power cycle...
            if (0 < waveforms.commandSlot[waveforms.wfidx].commands.Count)
            { 
              msg = "Please cycle LBHH power off-ON and\nTHEN promptly press OK";
              cap = "RESTART DEVICE";
              btn = MessageBoxButtons.OK;
              icon = MessageBoxIcon.Asterisk;
              MessageBox.Show(msg, cap, btn, icon);
              Application.DoEvents();
              Thread.Sleep(2000);
              //Reconnect to the Bolt
              if ( ! ConnectToLightning(comPort.PortName))
              { //There was an error reconnecting to Lightning - inform user and break programming process
                msg = "ERROR: Device restart check failed (ConnectToLightning)\nCANCELLING CONFIGURATION PROCESS";
                cap = "RESTART CHECK FAILURE";
                btn = MessageBoxButtons.OK;
                icon = MessageBoxIcon.Asterisk;
                MessageBox.Show(msg, cap, btn, icon);
                break;
              }
              // Once Ltng connected we can safely restart Bolt Communication by switching COMx
              ConnectToBoltDebugPort();           // Switch COMx to the BOLT Debug Port
              if ( ! GetWfmidFromBolt())
              { //There was an error reconnecting to BOLT - inform user and break programming process
                msg = "ERROR: Device restart check failed (GetWfmidFromBolt)\nCANCELLING CONFIGURATION PROCESS";
                cap = "RESTART CHECK FAILURE";
                btn = MessageBoxButtons.OK;
                icon = MessageBoxIcon.Asterisk;
                MessageBox.Show(msg, cap, btn, icon);
                break;
              }
              if ( ! HandlePassword())
              { // Incorrect password info in XML file - warn user and break programming process
                msg = "ERROR: Incorrect PASSWORD PROMPT or PASSWORD\n"
                    + "provided in XML file for BOLT waveform.\n"
                    + "Use XML appropriate for waveform loadout.\n"
                    + "CANCELLING BOLT CONFIGURATION PROCESS";
                cap = "INCORRECT PASSWORD";
                btn = MessageBoxButtons.OK;
                icon = MessageBoxIcon.Asterisk;
                MessageBox.Show(msg, cap, btn, icon);
                break;
              }
              if ( ! SendXmlCommands())
              { //An incorrect command was provided
                msg = "ERROR: Incorrect BOLT commands in XML file\nCANCELLING CONFIGURATION PROCESS";
                cap = "INCORRECT COMMAND";
                btn = MessageBoxButtons.OK;
                icon = MessageBoxIcon.Asterisk;
                MessageBox.Show(msg, cap, btn, icon);
                break;
              }
            } // end if XML specified BOLT configuration commands to be sent
          }
/*  TODO - 2022/05/01 Delete this now obsolete code
          else if (waveforms.hwiParams[waveforms.wfidx].password == "" && waveforms.hwiParams[waveforms.wfidx].passwordPrompt == "")
          { // When no password is required start by getting the Bolt's WFMID
            if ( ! GetUnitID())
            { // When there was an error finding the Unit ID - inform user and stop programming process
              msg = "ERROR: Transmit ID not found.\nPerhaps incorrect XML file selected";
              cap = "NO USER ID";
              btn = MessageBoxButtons.OK;
              icon = MessageBoxIcon.Asterisk;
              MessageBox.Show(msg, cap, btn, icon);
              break;
            }
   
            if ( ! GetAndLoadTek())
            { // When there was an error loading the TEK - inform user and stop programming process
              msg = "ERROR: Waveform's TEK(s) could not be loaded.\nWaveform already keyed, file not found, or BOLT generally unhappy.\nCANCELLING CONFIGURATION PROCESS";
              cap = "ISSUE LOADING TEK KEY";
              btn = MessageBoxButtons.OK;
              icon = MessageBoxIcon.Asterisk;
              MessageBox.Show(msg, cap, btn, icon);
              break;
            }

            if (0 >= waveforms.commandSlot[waveforms.wfidx].commands.Count)
            {
              break;
            }
            else if ( ! SendXmlCommands())
            { // When there was an error sending BOLT commands - inform user and stop programming process
              msg = "ERROR: Incorrect BOLT commands in XML file\nCANCELLING CONFIGURATION PROCESS";
              cap = "INCORRECT COMMANDS";
              btn = MessageBoxButtons.OK;
              icon = MessageBoxIcon.Asterisk;
              MessageBox.Show(msg, cap, btn, icon);
              break;
            }

            if (waveforms.hwiParams[waveforms.wfidx].resetCheckCmd != "") // TODO - use <ResetCheckCmd> in XML (not sure why this is nested here)
            {
              if ( ! CheckResetTEK())
              {
                Boolean tekLoaded = false;
                int attemptCount = 0;
                SystemSounds.Beep.Play();
                while ( ! tekLoaded && attemptCount < 3)
                {
                  attemptCount++;
                  String promptInput = Interaction.InputBox("Error TEK Key Pairing not found insert Actual TEK Key ID", "User Argument", "Default", -1, -1);
                  //Check if <Keying ID> string can be converted to integer, then if proper length, then try to get & load TEK
                  if ( ! Regex.IsMatch(promptInput, @"^\d+$"))
                  {
                    msg = "Improper Argument Format - value must be an integer\n";
                    cap = "WARNING: BAD FORMAT!";
                    btn = MessageBoxButtons.OK;
                    icon = MessageBoxIcon.Asterisk;
                    MessageBox.Show(msg, cap, btn, icon);
                  }
                  else if (promptInput.Length != 6)
                  {
                    msg = "Improper Number of Arguments please try again...\nBOLT is expecting 6 characters....\n";
                    cap = "WARNING: BAD FORMAT!";
                    btn = MessageBoxButtons.OK;
                    icon = MessageBoxIcon.Asterisk;
                    MessageBox.Show(msg, cap, btn, icon);
                  }
                  else
                  {
                    waveforms.hwiParams[waveforms.wfidx].tekOffset = (Int32.Parse(promptInput) - Int32.Parse(waveforms.boltParams[waveforms.wfidx].brickNumber)).ToString();
                    tekLoaded = GetAndLoadTek();
                  }
                }

                if (attemptCount == 3)
                {
                  msg = "Error: TEK Key not found. Ask code issuer for assistance\nCANCELLING CONFIGURATION PROCESS";
                  cap = "WARNING: BAD FORMAT!";
                  btn = MessageBoxButtons.OK;
                  icon = MessageBoxIcon.Asterisk;
                  MessageBox.Show(msg, cap, btn, icon);
                  break;
                }
                else
                {
                  if ( ! GetWfmidFromBolt())
                  { //There was an error loading the TEK key warn user and break programming process
                    msg = "ERROR: Device restart check failed (GetWfmidFromBolt)\nCANCELLING CONFIGURATION PROCESS";
                    cap = "RESTART CHECK FAILURE";
                    btn = MessageBoxButtons.OK;
                    icon = MessageBoxIcon.Asterisk;
                    MessageBox.Show(msg, cap, btn, icon);
                    break;
                  }
                  //Send Commands from XML (need to create 4 slots)
                  if ( ! SendXmlCommands())
                  { //An incorrect command was provided
                    msg = "ERROR: Incorrect BOLT commands in XML file\nCANCELLING CONFIGURATION PROCESS";
                    cap = "INCORRECT COMMANDS";
                    btn = MessageBoxButtons.OK;
                    icon = MessageBoxIcon.Asterisk;
                    MessageBox.Show(msg, cap, btn, icon);
                    break;
                  }
                } // end allowing multiple attempts
              }
            } // end if <ResetCheckCmd> in XML (not sure why this is nested here)
          } // end else no Debug Port password required
*/
          lightningWaveformChange = waveforms.hwiParams[waveforms.wfidx].lightningWFChange;
        }

        //Part three if there are more than one slot loaded
        //Send switch waveform to next non configured slot so we can load TEK etc.
        //Prompt user to restart
        waveforms.wfidx = wfmid % MAX_SLOT;
        numConfigwfrms++;
      } // end while (numConfigwfrms < MAX_SLOT)

      busyBar.Visible = false;

      if (numConfigwfrms == MAX_SLOT)
      {
        msg = "BOLT successfully keyed & configured\n\n"
            + "Now cycle LBHH power off-ON, then load\n"
            + "INFIL keys and WF Info before making an\n"
            + "EXFIL Choice from the LBHH Admin Menu.";
        cap = "GREAT SUCCESS!";
        btn = MessageBoxButtons.OK;
        icon = MessageBoxIcon.Asterisk;
        MessageBox.Show(msg, cap, btn, icon);
      }
      else
      {
        msg = "BOLT HAS NOT BEEN COMPLETELY KEYED & CONFIGURED\n"
            + "Please cycle LBHH power off-ON and retry";
        cap = "BOLT SETUP FAILED :( ! ";
        btn = MessageBoxButtons.OK;
        icon = MessageBoxIcon.Asterisk;
        MessageBox.Show(msg, cap, btn, icon);
      }
      DisconnectDevice();
      keyAndCfgBoltBtn.Enabled = true;
    } // end KeyAndCfgBoltBtn_Click

    #endregion Keying Tab

    //## REGION ################################################################
    #region WFInfo Tab

    private void QueryLbhhWaveformDataButton_Click(object sender, EventArgs e) {
      //First check device ID and configuration
      //if device id and available keys do not match error out and disconnect from device
      busyBar.Visible = true;           // Indicate activity to user
      string msg = "";
      boltDbg = ""; //Clear BOLT Debug Buffer  TODO 2022/05/01 not sure why this is here but doubt it harms anything either
      if (cfgParams.connected)
      {                                 // Reconfirm same LBHH connected
        cfgParams.connected = false;    //assume disconnected until otherwise proven
        QueueLtMessage(LightningMsgType.qKeyNames);
        Thread.Sleep(50);
        Application.DoEvents();
        while (loRespTimer.Enabled)  // Timer disabled on error (which sets cfgParams.connectd = false) or cmd completion
        {
          Thread.Sleep(10);
          Application.DoEvents();                 
        }
          //Check if we received expected message
        if (cfgParams.connected)
        { //Before poll wf info, ensure the ID Matches and the device was not hot swapped
          if (cfgParams.deviceId == gpiParams.deviceId) {
            QueueLtMessage(LightningMsgType.qWvfrmInfo);
            Thread.Sleep(50);
            while (loRespTimer.Enabled)
            {
              Thread.Sleep(10);
              Application.DoEvents();                 
            }
            if ( ! cfgParams.connected)
            {
              msg = "ERROR: LBHH NAKed query for WF Info.\nNote LBHH firmware version and\ncontact application tech support.";
            }

            gpiParams.LoadOptions(ref cfgParams);
            SetGpiFromProgParams(ref gpiParams);       // set GPI to last known programmed values

          } else {
            //Report to user that device was switched before finishing
            msg = "LOAD WAVEFORM CONFIGURATION FAILURE: Device ID: " + gpiParams.deviceId + " found during connection does not match Current ID: " + cfgParams.deviceId + " Do not hot-swap devices without Disconnecting! \nDisconnecting.....";
            lightningStatus.Image = xmarkImg;
            DisconnectDevice();
          }
        } else {
          //Failed to reach device
          msg = "LOAD WAVEFORM CONFIGURATION FAILURE: Device ID: " + gpiParams.deviceId + " could not be reached. \nDisconnecting.....";
          lightningStatus.Image = xmarkImg;
          DisconnectDevice();
        }
      } else {
        //Report that connection droppedout
        msg = "LOAD WAVEFORM CONFIGURATION FAILURE: Device ID: " + gpiParams.deviceId + " there is something happening here and what it is ain't exactly clear. \nDisconnecting.....";
        lightningStatus.Image = xmarkImg;
        DisconnectDevice();
      }

      if (msg != "") {
        string cap = "Configuration Notification";
        ShowModelessMessageBox(msg, cap, InfoType.Insight);
      }
      busyBar.Visible = false;
    } // end QueryLbhhWaveformDataButton_Click

    private void WaveformEnableForAdminButton_Click(object sender, EventArgs e) {
      // Validate fields first
      busyBar.Visible = true;          // Indicate activity to user
      if (ValidateWaveformConfigurationInformation()) {
        // build gpiParams.waveformInfo
        string curItem;
        int index;
        char wc = '0';  // waveform class (a base62 enumeration using characters 0-9,A-Z,a-z)
        // w0
        curItem = w0p2.SelectedItem.ToString();
        index = w0p2.FindString(curItem);
        if      (10 > index)  { wc = (char)('0' + index);       }
        else if (36 > index)  { wc = (char)('A' + index - 10);  }
        else if (62 > index)  { wc = (char)('a' + index - 36);  }
        else                  { busyBar.Visible = false; return;}
        gpiParams.waveformInfo[0, 0] = "0";               // Slot #   
        gpiParams.waveformInfo[0, 1] = wc.ToString();     // WF Class (aka Family)
        gpiParams.waveformInfo[0, 2] = w0p3.Text;         // Keying ID
        gpiParams.waveformInfo[0, 3] = w0p4.Text;         // Transmit ID
        gpiParams.waveformInfo[0, 4] = w0p5.Text;         // WF Name (e.g. WFE-8A)

        // w1
        curItem = w1p2.SelectedItem.ToString();
        index = w1p2.FindString(curItem);
        if      (10 > index)  { wc = (char)('0' + index);       }
        else if (36 > index)  { wc = (char)('A' + index - 10);  }
        else if (62 > index)  { wc = (char)('a' + index - 36);  }
        else                  { busyBar.Visible = false; return;}
        gpiParams.waveformInfo[1, 0] = "1";
        gpiParams.waveformInfo[1, 1] = wc.ToString();
        gpiParams.waveformInfo[1, 2] = w1p3.Text;
        gpiParams.waveformInfo[1, 3] = w1p4.Text;
        gpiParams.waveformInfo[1, 4] = w1p5.Text;

        // w2
        curItem = w2p2.SelectedItem.ToString();
        index = w2p2.FindString(curItem);
        if      (10 > index)  { wc = (char)('0' + index);       }
        else if (36 > index)  { wc = (char)('A' + index - 10);  }
        else if (62 > index)  { wc = (char)('a' + index - 36);  }
        else                  { busyBar.Visible = false; return;}
        gpiParams.waveformInfo[2, 0] = "2";
        gpiParams.waveformInfo[2, 1] = wc.ToString();
        gpiParams.waveformInfo[2, 2] = w2p3.Text;
        gpiParams.waveformInfo[2, 3] = w2p4.Text;
        gpiParams.waveformInfo[2, 4] = w2p5.Text;

        // w3
        curItem = w3p2.SelectedItem.ToString();
        index = w3p2.FindString(curItem);
        if      (10 > index)  { wc = (char)('0' + index);       }
        else if (36 > index)  { wc = (char)('A' + index - 10);  }
        else if (62 > index)  { wc = (char)('a' + index - 36);  }
        else                  { busyBar.Visible = false; return;}
        gpiParams.waveformInfo[3, 0] = "3";
        gpiParams.waveformInfo[3, 1] = wc.ToString();
        gpiParams.waveformInfo[3, 2] = w3p3.Text;
        gpiParams.waveformInfo[3, 3] = w3p4.Text;
        gpiParams.waveformInfo[3, 4] = w3p5.Text;


        // w4
        curItem = w4p2.SelectedItem.ToString();
        index = w4p2.FindString(curItem);
        if      (10 > index)  { wc = (char)('0' + index);       }
        else if (36 > index)  { wc = (char)('A' + index - 10);  }
        else if (62 > index)  { wc = (char)('a' + index - 36);  }
        else                  { busyBar.Visible = false; return;}
        gpiParams.waveformInfo[4, 0] = "4";
        gpiParams.waveformInfo[4, 1] = wc.ToString();
        gpiParams.waveformInfo[4, 2] = w4p3.Text;
        gpiParams.waveformInfo[4, 3] = w4p4.Text;
        gpiParams.waveformInfo[4, 4] = w4p5.Text;

        //First check device ID and configuration
        //if device id and available keys do not match error out
        //and disconnect from device
        string msg = "";
        boltDbg = ""; //Clear BOLT Debug Buffer  TODO 2022/05/01 not sure why this is here but doubt it harms anything either
        if (cfgParams.connected) {
          cfgParams.connected = false; //assume disconnected until otherwise proven
          QueueLtMessage(LightningMsgType.qKeyNames);
          Thread.Sleep(50);
          Application.DoEvents();
          while (loRespTimer.Enabled)  // Timer disabled on error (which sets cfgParams.connectd = false) or cmd completion
          {
            Thread.Sleep(10);
            Application.DoEvents();                 
          }
            //Check if we received expected message
          if (cfgParams.connected) {
            //Ensure the ID Matches and the device was not hot swapped
            if (cfgParams.deviceId == gpiParams.deviceId) {
              for (int i = 0; i <= MAX_SLOT && cfgParams.connected; i++)
              {
                Thread.Sleep(50);             // Give time for LoMsgThread to post Ltng ACK/NAK response to RichTextBox
                Application.DoEvents();                 
                busyBar.Refresh();
                gpiParams.actOnWaveform = i;
                QueueLtMessage(LightningMsgType.lWvfrmInfo);
                Thread.Sleep(50);
                while (loRespTimer.Enabled)  // Timer disabled on error (which sets cfgParams.connectd = false) or cmd completion
                {
                  loMsgEvent.Set();                       // Signal 'Receive LO Messages' thread to run
                  Thread.Sleep(10);
                  Application.DoEvents();                 
                }
              }
              if (cfgParams.connected)
              { // If still connected then Lightning did not NAK any command
                ShowModelessMessageBox("SUCCESS: All waveform info loaded.", "LBHH: SUCCESS", InfoType.Insight);
              }
              else
              { // LoMsgThread will clear .connected when there is an issue
                msg = "ERROR: Not all waveform info loaded. \nPlease Retry Configuration";
                DisconnectDevice();
              }
            } else {
              //Report to user that device was switched before finishing
              msg = "LIGHTNING WAVEFORM CONFIGURATION FAILURE: Device ID: " + gpiParams.deviceId + " found during connection does not match Current ID: " + cfgParams.deviceId + " Do not hot-swap devices without Disconnecting! \nDisconnecting.....";
              lightningStatus.Image = xmarkImg;
              DisconnectDevice();
            }
          } else {
            //Failed to reach device
            msg = "LIGHTNING WAVEFORM CONFIGURATION FAILURE: Device ID: " + gpiParams.deviceId + " could not be reached. \nDisconnecting.....";
            lightningStatus.Image = xmarkImg;
            DisconnectDevice();
          }
        } else {
          //Report that connection droppedout
          msg = "LIGHTNING WAVEFORM CONFIGURATION FAILURE: Device ID: " + gpiParams.deviceId + " there is something happening here and what it is ain't exactly clear. \nDisconnecting.....";
          lightningStatus.Image = xmarkImg;
          DisconnectDevice();
        }
        //
        if (msg != "") {
          string cap = "Configuration Notification";
          ShowModelessMessageBox(msg, cap, InfoType.Notification);
        }
      }
      busyBar.Visible = false;
    } // end WaveformEnableForAdminButton_Click

    private void w0p2_SelectedIndexChanged(object sender, EventArgs e) {
      string curItem = w0p2.SelectedItem.ToString();
      int index = w0p2.FindString(curItem);
      w0p5.Items.Clear();

      if (index == 0) {
        w0p3.Text = "000000";
        w0p4.Text = "000000";
        w0p5.Items.Add("NoWF_0");
        w0p5.SelectedIndex = 0;
        w0p5.DropDownStyle = ComboBoxStyle.DropDown;
      } else {
        w0p5.DropDownStyle = ComboBoxStyle.DropDownList;
        int b = 1;
        while (waveformConfigData[index, b] != null) {
          w0p5.Items.Add(waveformConfigData[index, b]);
          b++;
        }
      }			
    }

    private void w1p2_SelectedIndexChanged(object sender, EventArgs e) {
      string curItem = w1p2.SelectedItem.ToString();
      int index = w1p2.FindString(curItem);
      w1p5.Items.Clear();

      if (index == 0) {
        w1p3.Text = "000000";
        w1p4.Text = "000000";
        w1p5.Items.Add("NoWF_1");
        w1p5.SelectedIndex = 0;
        w1p5.DropDownStyle = ComboBoxStyle.DropDown;
      } else {
        w1p5.DropDownStyle = ComboBoxStyle.DropDownList;
        int b = 1;
        while (waveformConfigData[index, b] != null) {
          w1p5.Items.Add(waveformConfigData[index, b]);
          b++;
        }
      }
    }

    private void w2p2_SelectedIndexChanged(object sender, EventArgs e) {
      string curItem = w2p2.SelectedItem.ToString();
      int index = w2p2.FindString(curItem);
      w2p5.Items.Clear();

      if (index == 0) {
        w2p3.Text = "000000";
        w2p4.Text = "000000";
        w2p5.Items.Add("NoWF_2");
        w2p5.SelectedIndex = 0;
        w2p5.DropDownStyle = ComboBoxStyle.DropDown;
      } else {
        w2p5.DropDownStyle = ComboBoxStyle.DropDownList;
        int b = 1;
        while (waveformConfigData[index, b] != null) {
          w2p5.Items.Add(waveformConfigData[index, b]);
          b++;
        }
      }
    }

    private void w3p2_SelectedIndexChanged(object sender, EventArgs e) {
      string curItem = w3p2.SelectedItem.ToString();
      int index = w3p2.FindString(curItem);
      w3p5.Items.Clear();

      if (index == 0) {
        w3p3.Text = "000000";
        w3p4.Text = "000000";
        w3p5.Items.Add("NoWF_3");
        w3p5.SelectedIndex = 0;
        w3p5.DropDownStyle = ComboBoxStyle.DropDown;
      } else {
        w3p5.DropDownStyle = ComboBoxStyle.DropDownList;
        int b = 1;
        while (waveformConfigData[index, b] != null) {
          w3p5.Items.Add(waveformConfigData[index, b]);
          b++;
        }
      }
    }

    private void w4p2_SelectedIndexChanged(object sender, EventArgs e) {
      string curItem = w4p2.SelectedItem.ToString();
      int index = w4p2.FindString(curItem);
      w4p5.Items.Clear();

      if (index == 0) {
        w4p3.Text = "000000";
        w4p4.Text = "000000";
        w4p5.Items.Add("NoWF_4");
        w4p5.SelectedIndex = 0;
        w4p5.DropDownStyle = ComboBoxStyle.DropDown;
      } else {
        w4p5.DropDownStyle = ComboBoxStyle.DropDownList;
        int b = 1;
        while (waveformConfigData[index, b] != null) {
          w4p5.Items.Add(waveformConfigData[index, b]);
          b++;
        }
      }
    }

    private void clear0_Click(object sender, EventArgs e) {
      w0p2.SelectedIndex = 0;
      w0p3.Text = "000000";
      w0p4.Text = "000000";
      w0p5.DropDownStyle = ComboBoxStyle.DropDown;
      w0p5.Items.Clear();
      w0p5.Items.Add("NoWF_0");
      w0p5.SelectedIndex = 0;
    }

    private void clear1_Click(object sender, EventArgs e) {
      w1p2.SelectedIndex = 0;
      w1p3.Text = "000000";
      w1p4.Text = "000000";
      w1p5.DropDownStyle = ComboBoxStyle.DropDown;
      w1p5.Items.Clear();
      w1p5.Items.Add("NoWF_1");
      w1p5.SelectedIndex = 0;
    }

    private void clear2_Click(object sender, EventArgs e) {
      w2p2.SelectedIndex = 0;
      w2p3.Text = "000000";
      w2p4.Text = "000000";
      w2p5.DropDownStyle = ComboBoxStyle.DropDown;
      w2p5.Items.Clear();
      w2p5.Items.Add("NoWF_2");
      w2p5.SelectedIndex = 0;
    }

    private void clear3_Click(object sender, EventArgs e) {
      w3p2.SelectedIndex = 0;
      w3p3.Text = "000000";
      w3p4.Text = "000000";
      w3p5.DropDownStyle = ComboBoxStyle.DropDown;
      w3p5.Items.Clear();
      w3p5.Items.Add("NoWF_3");
      w3p5.SelectedIndex = 0;
    }

    private void clear4_Click(object sender, EventArgs e) {
      w4p2.SelectedIndex = 0;
      w4p3.Text = "000000";
      w4p4.Text = "000000";
      w4p5.DropDownStyle = ComboBoxStyle.DropDown;
      w4p5.Items.Clear();
      w4p5.Items.Add("NoWF_4");
      w4p5.SelectedIndex = 0;
    }

    #endregion WFInfo Tab

    //## REGION ################################################################
    #region Activate Tab

    private void infilKey1_CheckedChanged(object sender, EventArgs e)
    {
      if (infilKey1.Checked)
      {
        gpiParams.Selectedinfilkey = 1;
        gpiParams.SelectedinfilkeyString = infilKey1.Text;
      }
    }

    private void infilKey2_CheckedChanged(object sender, EventArgs e)
    {
      if (infilKey2.Checked)
      {
        gpiParams.Selectedinfilkey = 2;
        gpiParams.SelectedinfilkeyString = infilKey2.Text;
      }
    }

    private void infilKey3_CheckedChanged(object sender, EventArgs e)
    {
      if (infilKey3.Checked)
      {
        gpiParams.Selectedinfilkey = 3;
        gpiParams.SelectedinfilkeyString = infilKey3.Text;
      }
    }

    private void infilKey4_CheckedChanged(object sender, EventArgs e)
    {
      if (infilKey4.Checked)
      {
        gpiParams.Selectedinfilkey = 4;
        gpiParams.SelectedinfilkeyString = infilKey4.Text;
      }
    }

    private void infilKey5_CheckedChanged(object sender, EventArgs e)
    {
      if (infilKey5.Checked)
      {
        gpiParams.Selectedinfilkey = 5;
        gpiParams.SelectedinfilkeyString = infilKey5.Text;
      }
    }

    private void noKey0_CheckedChanged(object sender, EventArgs e)
    {
      if (noKey0.Checked)
      {
        gpiParams.Selectedinfilkey = 0;
        gpiParams.SelectedinfilkeyString = noKey0.Text;
      }
    }

    // Name: configureInfilButton_Click
    // Arguments: sender
    //			      e      
    // Description: Configures selected Infil keys and Exfil waveforms for use in Lightning.
    private void configureInfilButton_Click(object sender, EventArgs e) {
      busyBar.Visible = true;          // Indicate activity to user
      //First check device ID and configuration
      //if device id and available keys do not match error out
      //and disconnect from device
      InfoType msgInfoType = InfoType.Notification;   // most msg boxes in this method are notifications
      string msg = "";
      boltDbg = ""; //Clear BOLT Debug Buffer  TODO 2022/05/01 not sure why this is here but doubt it harms anything either
      infilgrpBox.Enabled = false;
      configureInfilButton.Enabled = false;
      connectButton.Enabled = false;
      if (cfgParams.connected) {

        cfgParams.connected = false; //assume disconnected until otherwise proven
        QueueLtMessage(LightningMsgType.qKeyNames);
        Thread.Sleep(50);
        Application.DoEvents();
        while (loRespTimer.Enabled)  // Timer disabled on error (which sets cfgParams.connectd = false) or cmd completion
        {
          Thread.Sleep(10);                      
          Application.DoEvents();                 
        }
          //Check if we received expected message
        if (cfgParams.connected) {
          //Ensure the ID Matches and the device was not hot swapped
          if (cfgParams.deviceId == gpiParams.deviceId) {
            // Select the Infil message
            QueueLtMessage(LightningMsgType.sGrpInfil);
            Thread.Sleep(50);
            while (loRespTimer.Enabled)  // Timer disabled on error (which sets cfgParams.connectd = false) or cmd completion
            {
              Thread.Sleep(10);                      
              Application.DoEvents();                 
            }

            //Double check Infil poll
            QueueLtMessage(LightningMsgType.qActvKeyNames);
            Thread.Sleep(50);
            while (loRespTimer.Enabled)  // Timer disabled on error (which sets cfgParams.connectd = false) or cmd completion
            {
              Thread.Sleep(10);
              Application.DoEvents();                 
            }

            //If the infil matches expected values (device ID and Selected Value) then start the Bolt process
            if (cfgParams.deviceId == gpiParams.deviceId && ((cfgParams.SelectedinfilkeyString == gpiParams.SelectedinfilkeyString) || (cfgParams.SelectedinfilkeyString == "0GpKey" && gpiParams.Selectedinfilkey == 0))) {
              //Configuration was a success
              msgInfoType = InfoType.Insight;
              msg = "LIGHTNING CONFIGURATION SUCCESS: Device " + gpiParams.deviceId + " was configured to use the " + gpiParams.SelectedinfilkeyString;
              lightningStatus.Image = CheckMrkimg;
              connectButton.Enabled = true;
              infilgrpBox.Enabled = true;
              configureInfilButton.Enabled = true;
            }
            //Command did not go through below
            else {
              //Error Device
              msg = "LIGHTNING CONFIGURATION FAILURE: Device " + gpiParams.deviceId + " was not configured to use infil key: " + gpiParams.SelectedinfilkeyString + "reported key: " + cfgParams.Selectedinfilkey;
              lightningStatus.Image = xmarkImg;
              connectButton.Enabled = true;
              infilgrpBox.Enabled = true;
              configureInfilButton.Enabled = true;
              DisconnectDevice();
            }
          } else {
            //Report to user that device was switched before finishing
            msg = "LIGHTNING CONFIGURATION FAILURE: Device ID: " + gpiParams.deviceId + " found during connection does not match Current ID: " + cfgParams.deviceId + " Do not hot-swap devices without Disconnecting! \nDisconnecting.....";
            lightningStatus.Image = xmarkImg;
            DisconnectDevice();
          }
        } else {
          //Failed to reach device
          msg = "LIGHTNING CONFIGURATION FAILURE: Device ID: " + gpiParams.deviceId + " could not be reached. \nDisconnecting.....";
          lightningStatus.Image = xmarkImg;
          DisconnectDevice();
        }
      } else {
        //Report that connection droppedout
        msg = "LIGHTNING CONFIGURATION FAILURE: Device ID: " + gpiParams.deviceId + " there is something happening here and what it is ain't exactly clear. \nDisconnecting.....";
        lightningStatus.Image = xmarkImg;
        DisconnectDevice();
      }
      //
      if (msg != "") {
        string cap = "Configuration Notification";
        ShowModelessMessageBox(msg, cap, msgInfoType);
      }
      busyBar.Visible = false;
    }

    #endregion Activate Tab

    //## REGION ################################################################
    #region Misc Tab

    //Name: UnlockBtn_Click
    //Arguments: object sender
    //           EventArgs e
    //Description: Unlock interface for adding group key pairings
    private void UnlockBtn_Click(object sender, EventArgs e) {
      busyBar.Visible = true;          // Indicate activity to user
      //Toggle group box if it is disabled and password does not match anticipated value
      if ( ! groupBox4.Enabled) {
        unlockBtn.Text = "Lock";

        string currentSheet = "Sponsor Names";
        Excel.Worksheet excelWorksheet;
        EncryptionFilePathName = appDir + DefEncryptionFileName;
        if (File.Exists(EncryptionFilePathName)) {
          //---------------------------------------------------------Open Excel Document-------------------------------------------------------------
          // Open Document and set global variable for reading
          EncryptionFilePathName = appDir + DefEncryptionFileName;
          Excel.Application appExcel = new Excel.Application();
          appExcel.Visible = false;
          Excel.Workbook WB = appExcel.Workbooks.Open(EncryptionFilePathName, ReadOnly: false, Password: EncryptionPassword);
          Excel.Sheets excelSheets = WB.Worksheets;
          excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
          //-----------------------------------------------------------------------------------------------------------------------------------------
          try {
            groupNames = new List<string>();
            sponsorNames = new List<string>();
            excelWorksheet.Unprotect(EncryptionPassword); // Unlock for editing
            if ( ! string.IsNullOrEmpty((excelWorksheet.Cells[1, 1] as Excel.Range).Value)) {
              // Pull information from Excel
              excelWorksheet.Cells[1, 1].NumberFormat = "@"; // Ensure format is string
              int i = 1;
              while ( ! string.IsNullOrEmpty(Convert.ToString((excelWorksheet.Cells[i, 1] as Excel.Range).Value))) {
                excelWorksheet.Cells[i, 1].NumberFormat = "@"; // Ensure format is string
                sponsorNames.Add(Convert.ToString((excelWorksheet.Cells[i, 1] as Excel.Range).Value));
                i++;
              }

              sponsorNameKeyPairString.Items.Clear();
              foreach (string sponsor in sponsorNames) {
                sponsorNameKeyPairString.Items.Add(sponsor);
              }
            }

            excelWorksheet.UsedRange.Columns.AutoFit();
            excelWorksheet.Protect(EncryptionPassword); // Lock Editing
            currentSheet = "ActiveGroupKeys";
            excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
            excelWorksheet.Unprotect(EncryptionPassword);
            if ( ! string.IsNullOrEmpty((excelWorksheet.Cells[2, 2] as Excel.Range).Value)) {
              // Pull information from Excel
              excelWorksheet.Cells[2, 3].NumberFormat = "@"; // Ensure format is string
              int i = 2;
              while ( ! string.IsNullOrEmpty(Convert.ToString((excelWorksheet.Cells[i, 2] as Excel.Range).Value))) {
                if ( ! string.IsNullOrEmpty(Convert.ToString((excelWorksheet.Cells[i, 3] as Excel.Range).Value))) {
                  excelWorksheet.Cells[i, 3].NumberFormat = "@"; // Ensure format is string
                  groupNames.Add(Convert.ToString((excelWorksheet.Cells[i, 3] as Excel.Range).Value));
                }
      
                i++;
              }

              groupNameKeyPairString.Items.Clear();
              foreach (string group in groupNames) {
                groupNameKeyPairString.Items.Add(group);
              }
            }

            //---------------------------------------------------------Close Excel Document------------------------------------------------------------
            excelWorksheet.UsedRange.Columns.AutoFit();
            excelWorksheet.Protect(EncryptionPassword); // Lock Editing
            WB.Save();
            appExcel.Workbooks.Close();
            appExcel.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(WB);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(appExcel);
            //-----------------------------------------------------------------------------------------------------------------------------------------
            groupBox4.Enabled = true;
          } catch (Exception) {
            WB.Save();
            appExcel.Workbooks.Close();
            appExcel.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(WB);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(appExcel);
            unlockBtn.Text = "Unlock";
          }
          
        } else {
          // Prompt User that File linked is incorrect format and must be resolved before using this feature use popup window
          // Eventually this will be replaced with a popup window that allows a user to set the path of Keys which are then copied to the default directory
          // This will be done in the second release
          // Prompt User that File linked is incorrect format and must be resolved before using this feature (via popup window)
          string msg = "FAILED: Encryption Key File cannot be found at expected location: " + EncryptionFilePathName;
          string cap = "Key Retrieval";
          MessageBoxButtons btn = MessageBoxButtons.OK;
          MessageBoxIcon icon = MessageBoxIcon.Asterisk;
          MessageBox.Show(msg, cap, btn, icon);
          unlockBtn.Text = "Unlock";
        }		
      // If the group box is enabled
      } else {
        groupBox4.Enabled = false;
        unlockBtn.Text = "Unlock";
        resetMiscTab();
      }
      busyBar.Visible = false;
    }


    private void addGroupButton_Click(object sender, EventArgs e) {
      busyBar.Visible = true;          // Indicate activity to user
      string groupInput = groupNameKeyPairString.Text;
      string sponsorInput = sponsorNameKeyPairString.Text;
      // Make sure that first character is a capital letter and length is correct
      if (groupInput.Length == 6 && (groupInput[0] >= 'A') && (groupInput[0] <= 'Z')) {
        if (sponsorInput.Length >= 3 && (sponsorInput[0] >= 'A') && (sponsorInput[0] <= 'Z') && ( ! sponsorInput.Contains("/"))) {
          // Edit the Active Keys 
          AddRemoveActiveKeys(groupInput, sponsorInput, (int)groupNameKeyIndexUpDown.Value);
        } else {
          // Throw invalid sponsor name formatting error
          string msg = "ERROR: Invalid Sponsor Name Format: Name must be at least 3 characters long, start with a capital letter, and not contain and forward slashes (/)";
          string cap = "ERROR: Invalid Key Format!";
          MessageBoxButtons btn = MessageBoxButtons.OK;
          MessageBoxIcon icon = MessageBoxIcon.Asterisk;
          MessageBox.Show(msg, cap, btn, icon);
        }
          
      } else {
        // Throw invalid group name formatting error
        string msg = "ERROR: Invalid Group Name Format: Name must be 6 characters long and start with a capital letter";
        string cap = "ERROR: Invalid Key Format!";
        MessageBoxButtons btn = MessageBoxButtons.OK;
        MessageBoxIcon icon = MessageBoxIcon.Asterisk;
        MessageBox.Show(msg, cap, btn, icon);
      }
      busyBar.Visible = false;
    }

    private void removeGroupButton_Click(object sender, EventArgs e) {
      busyBar.Visible = true;          // Indicate activity to user
      AddRemoveActiveKeys("", "", (int)groupNameKeyIndexUpDown.Value);
      busyBar.Visible = false;
    }

    private void addSponsorButton_Click(object sender, EventArgs e) {
      busyBar.Visible = true;          // Indicate activity to user
      string groupInput = groupNameKeyPairString.Text;
      string sponsorInput = sponsorNameKeyPairString.Text;

      // Check if Sponsor Name is valid
      if (sponsorInput.Length >= 3 && (sponsorInput[0] >= 'A') && (sponsorInput[0] <= 'Z') && ( ! sponsorInput.Contains("/"))) {
        // Add the Sponsor Name
        AddSponsorName(groupInput, sponsorInput, (int)groupNameKeyIndexUpDown.Value);
      } else {
        // Throw invalid sponsor name formatting error
        string msg = "ERROR: Invalid Sponsor Name Format: Name must be at least 3 characters long, start with a capital letter, and not contain and forward slashes (/)";
        string cap = "ERROR: Invalid Key Format!";
        MessageBoxButtons btn = MessageBoxButtons.OK;
        MessageBoxIcon icon = MessageBoxIcon.Asterisk;
        MessageBox.Show(msg, cap, btn, icon);
      }
      busyBar.Visible = false;
    }

    private void removeSponsorButton_Click(object sender, EventArgs e) {
      busyBar.Visible = true;          // Indicate activity to user
      string groupInput = groupNameKeyPairString.Text;
      string sponsorInput = sponsorNameKeyPairString.Text;

      // Check if Sponsor Name is valid
      if (sponsorInput.Length >= 3 && (sponsorInput[0] >= 'A') && (sponsorInput[0] <= 'Z') && ( ! sponsorInput.Contains("/"))) {
        // Add the Sponsor Name
        RemoveSponsorName(groupInput, sponsorInput, (int)groupNameKeyIndexUpDown.Value);
      } else {
        // Throw invalid sponsor name formatting error
        string msg = "ERROR: Invalid Sponsor Name Format: Name must be at least 3 characters long, start with a capital letter, and not contain and forward slashes (/)";
        string cap = "ERROR: Invalid Key Format!";
        MessageBoxButtons btn = MessageBoxButtons.OK;
        MessageBoxIcon icon = MessageBoxIcon.Asterisk;
        MessageBox.Show(msg, cap, btn, icon);
      }
      busyBar.Visible = false;
    }

    private void exportGroupKeysButton_Click(object sender, EventArgs e) {
      string sponsorInput = sponsorNameKeyPairString.Text;

      if ( ! sponsorNames.Contains(sponsorInput)) {
        // Throw a non-existing sponsor name error
        string msg = "ERROR: Invalid Sponsor Name: Name does not exist in keychain";
        string cap = "ERROR: Invalid Sponsor Name!";
        MessageBoxButtons btn = MessageBoxButtons.OK;
        MessageBoxIcon icon = MessageBoxIcon.Hand;
        MessageBox.Show(msg, cap, btn, icon);
      } else {
        if (exportPassword.Text.Length == 0) {
          string msg = "ERROR: Password field must not be empty";
          string cap = "ERROR";
          MessageBoxButtons btn = MessageBoxButtons.OK;
          MessageBoxIcon icon = MessageBoxIcon.Hand;
          MessageBox.Show(msg, cap, btn, icon);
        } else {
          busyBar.Visible = true;          // Indicate activity to user
          ExportSponsorKeys(sponsorInput);
          busyBar.Visible = false;
        }
      }
    }

    //Name: atOnBtn_Click
    //Arguments: sender
    //           e
    //Description: Enable device anti-tamper
    private void atOnBtn_Click(object sender, EventArgs e) {
      busyBar.Visible = true;          // Indicate activity to user
      QueueLtMessage(LightningMsgType.sAntiTampOn);
      Thread.Sleep(50);
      while (loRespTimer.Enabled)  // Timer disabled on error (which sets cfgParams.connectd = false) or cmd completion
      {
        Thread.Sleep(10);
        Application.DoEvents();                 
      } //end
      busyBar.Visible = false;
    }

    //Name: atOffBtn_Click
    //Arguments: sender
    //           e
    //Description: Disable device anti-tamper
    private void atOffBtn_Click(object sender, EventArgs e) {
      busyBar.Visible = true;          // Indicate activity to user
      QueueLtMessage(LightningMsgType.sAntiTampOff);
      Thread.Sleep(50);
      while (loRespTimer.Enabled)  // Timer disabled on error (which sets cfgParams.connectd = false) or cmd completion
      {
        Thread.Sleep(10);
        Application.DoEvents();                 
      }
      busyBar.Visible = false;
    }

    private void showCharactersCheckBox_CheckedChanged(object sender, EventArgs e)
    {
      if ( ! showCharactersCheckBox.Checked)
      {
        exportPassword.PasswordChar = '*';
      }
      else
      {
        exportPassword.PasswordChar = '\0';
      }
    }

    #endregion Misc Tab

    //## REGION ################################################################
    #region Communication/Command Tab

    //Name: sendCmdBtn_Click
    //Arguments: sender
    //           e
    //Description: Sends command in the Diagnostic Box to the Bolt
    private void sendCmdBtn_Click(object sender, EventArgs e) {
      //If these labels are not hidden use lightning format and communication method
      if (ltngMsgStopLbl.Visible && ltngMsgStartLbl.Visible) {
        //Queue cmdTextbox message
        QueueLtMessage(LightningMsgType.sendLtngUsrCmd);
        loRespStatus = LoRespStatus.SWFreq;
        for (int i = 0; i < 10; i++) {
          Thread.Sleep(100);
          Application.DoEvents();
        }
      } else {
        Application.DoEvents();
        SendBoltCommandSlowly((gpiParams.cmdTextbox), false);
      }
    }

    //Name: ConnectBoltBtn_Click
    //Arguments: object sender
    //           EventArgs e
    //Description: Connects the hardware interface to the Bolt Handheld
    //
    private void ConnectBoltBtn_Click(object sender, EventArgs e)
    {
      busyBar.Visible = true;                 // Indicate activity to user
      QueueLtMessage(LightningMsgType.sCOMxBoltDbg);
      ltngMsgStartLbl.Visible = false;        // Semi-obscure way to indicate that we just
      ltngMsgStopLbl.Visible = false;         // commanded Lightning to expose BOLT Debug Port
      boltDbg = "";                           //Reinitialize captured BOLT Debug Port output
      Thread.Sleep(500);                      //Delay to let Red Tool send Set COMx and Lightning to act on it
      StopMessagingThreads();                 //Kill Lightning messaging threads as now want to talk to BOLT
/*    // TODO see if this if{} can be deleted since OpenBoltComPort will close port 2022/05/01
      if (comPort.IsOpen) {                   // Changing port settings, so disconnect from connected port
        StopDiagThread();                     // No need to run diag while connecting
        StopMessagingThreads();               // Shutdown messaging threads before changing port
        Thread.Sleep(200);                    // Give device more time to see 'Disconnect'
        Application.DoEvents();               // and process any message request/user inputs
      }
*/
      OpenBoltComPort(comPort.PortName);

      if ( ! runDiagThread)
      {                                       // WHen diag thread not running,
        runDiagThread = true;                 // set flag that keeps it running
        diagBackgroundWorker.RunWorkerAsync();// and then kick it off
      }
      const int _50ms = 50;                   // use 50ms short sleeps
      int countDown = (30000/_50ms);          // for up to 30 seconds which is required
      while (0 < countDown)                   // for BOLT to get, process, and complete a
      {                                       // change of waveform (completely reboots BOLT)
        Thread.Sleep(_50ms);                  // use short sleeps so progressBar scrolls.
        Application.DoEvents();               // Each time the 'while loop' is executed
        countDown--;                          // decrement the countdown until reach 0
        //busyBar.Refresh();
        //Thread.Sleep(50);
        //Application.DoEvents();                 
        if (boltDbg.Contains("Initializing:.") && !boltDbg.Contains("MONITOR>"))
        {                                     // When find BOLT prompt indicating it is temporarily waiting for a password
          SendBoltCommandSlowly("*", false);  // write "*" out the COM port to stop BOLT from progressing beyond password entry function
          Application.DoEvents();             // update behind the scenes application stuff
          Thread.Sleep(200);                  // Wait for BOLT Debug Port to output text
          countDown = 0;                      // terminate while (countDown)
        }                                     // If BOLT already emitted "MONITOR>" we missed our chance to send password
      }
      busyBar.Visible = false;
    } // end ConnectBoltBtn_Click

    //Name: savediagBtn_Click
    //Arguments: sender
    //           e
    //Description: Save Diagnostic Data to user designated file
    private void savediagBtn_Click(object sender, EventArgs e) {
      /*
       * Generate a suggested filename for the diag data file and then display a Save File
       * dialog for the user.  Don't let the user overwrite a read-only file
       */
      DateTime dateTime = DateTime.UtcNow;
      string fileName = "BOLT-Diag ";
      //Automatically add timestamp to file name
      fileName += string.Format("{0:D4}-{1:D2}-{2:D2} {3:D2}{4:D2}Z.rtf",
                                dateTime.Year, dateTime.Month, dateTime.Day,
                                dateTime.Hour, dateTime.Minute);

      SaveFileDialog saveFileDialog = new SaveFileDialog();
      saveFileDialog.DefaultExt = "rtf";
      saveFileDialog.AddExtension = true;
      saveFileDialog.OverwritePrompt = true;
      saveFileDialog.RestoreDirectory = true;
      saveFileDialog.InitialDirectory = diagDir; // Change to diagnostic directory
      saveFileDialog.Filter = "BOLT Diag File (*.rtf)|*.rtf";
      saveFileDialog.FileName = fileName;

      // Show the Dialog. If user selected a file and clicked OK, use it.
      if (saveFileDialog.ShowDialog() == DialogResult.OK) {
        fileName = saveFileDialog.FileName;
        //Ensure we can actually write to the desired file
        if (File.Exists(fileName) && FileAttributes.ReadOnly == (File.GetAttributes(fileName) & FileAttributes.ReadOnly)) {
          string msg = "Sorry! You don't have permission\nto overwrite this read-only file.";
          string cap = "File Is Read-Only";
          MessageBoxButtons btn = MessageBoxButtons.OK;
          MessageBoxIcon icon = MessageBoxIcon.Asterisk;
          MessageBox.Show(msg, cap, btn, icon);
        } else {
          diagRichTextBox.SaveFile(fileName);
        }
      }
    }

    //Name: savediagBtn_Click
    //Arguments: sender
    //           e
    //Description: Clear Items inside the diagnostic text box
    private void clrdiagBtn_Click(object sender, EventArgs e) {
      diagRichTextBox.Clear();
    }

    //Name: cmdTxtBox_TextChanged
    //Arguments: object sender
    //           EventArgs e
    //Description: Updates internal variable to be sent to Lightning/Bolt when text box is changed
    private void cmdTxtBox_TextChanged(object sender, EventArgs e)
    {
      //Update command to send when value is changed in the text box
      gpiParams.cmdTextbox = cmdTxtBox.Text;
    }

    #endregion Communication/Command Tab

  }
}

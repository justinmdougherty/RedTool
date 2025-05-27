using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Media;                           // Needed for SystemSounds.Beep
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using System.Timers;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;                  //Needed for Interaction.InputBox

namespace LBHH_Red
{
  public partial class LBHH_HWI : Form
  {
    //## REGION ################################################################
    #region DECLARATIONS

    Version appVer = new Version(1, 0, 22121);
    readonly string ConnectedTitle;
    readonly string DisconnectedTitle;
    const string ConfigFolder = "\\Configuration Files";
    const string DiagnosticsFolder = "\\Diagnostics";
    const string ExportFolder = "\\Export";
    const string TekFolder = "\\Tek Keys";

    const int MAX_SLOT = 4;       // BOLT wf slot, aka WFMID, is 0-based i.e. 0-4

    public class BoltKey
    {
      public bool deviceOpen;     // TODO - useful??? or delete
      public bool idOpen;
      public bool keyOpen;
      public bool familyOpen;
      public string family; //Waveform family
      public string id;
      public string[] key;
      public BoltKey()
      {
        deviceOpen = false;
        idOpen = false;
        keyOpen = false;
        familyOpen = false;
        key = new string[10] { "", "", "", "", "", "", "", "", "", "" };
        id = "";
        family = "";
      }
    }
    public class BoltKeyChain
    {
      //public bool titleOpen;
      //public string titleString;
      public List<BoltKey> boltKeys;
      public BoltKeyChain()
      {
        //titleString = "";
        //titleOpen = false;
        boltKeys = new List<BoltKey>();
      }
    }

    public class BoltParameters   // This is information obtained from the BOLT and NIWC
    {
      public string brickNumber;  // Brick number retrieved from Bolt (used to index TEK Key for most waveforms)
      public string unitID;       // Unit ID retrieved from Bolt (AME transmit ID)
      public BoltKey tekKey;      // Tek key retrieved from TEK file (from file NIWC provides)
      public BoltParameters()
      {
        brickNumber = "";
        unitID = "";
        tekKey = new BoltKey();
      }
    }

    public class WaveformHWICommands
    {
      public string type;
      public int NumArguments;
      public List<string> ArgumentHelp;
      public List<string> ArgumentType;
      public WaveformHWICommands()
      {
        type = "";
        NumArguments = 0;
        ArgumentHelp = new List<string>();
        ArgumentType = new List<string>();
      }
    }
    // Hardware interface commands 
    public class WaveformHWIParams
    {
      public string appName;          // <name> in XML file - text to display in Red Tool Slot XML textbox
      public string title;            // <Waveform> in XML file - string from BOLT that IDs waveform
      public string passwordPrompt;   // <passwordPrompt> in XML file - string from BOLT prompting password entry
      public string password;         // <password> in XML file - wf specific password to send to BOLT
      public string brickPrompt;      // <BrickNumber> attribute in XML file - use to extract BN from BOLT output
      public string brickTrigger;     // <BrickNumber> attribute in XML file
      public string unitIDprompt;     // <UnitID> attribute in XML file - use to extract TX ID from BOLT output
      public string unitIDtrigger;    // <UnitID> attribute in XML file
      public string cmdPrompt;        // <prompt> in XML file - string from BOLT, command line prompt
      public string loadtekCommand;   // <TEKLoad> in XML file - wf specific command to load TEK(s)
      public string tekFile;          // <TEKFile> in XML file - name of local file in which TEKs are stored
      public string tekOffset;        // <TEKOffset> in XML file - TEK file keying index offset from Brick Number
      public List<Int32> keyOrder;    // <KeyOrder> in XML file - order in which to grab TEK from list of TEKs for given Keying ID
      public WaveformHWICommands keyingIDPrompt;  // <KeyingID> in XML file - stuff to allow user input of Keying ID
      public bool lightningWFChange;  // <LightningWFChange> in XML file - indicates if Lightning changes wf via User Port
      public string resetCheckCmd;    // <ResetCheckCmd> in XML file (never used as of 2022/05/01, if ever)  TODO
      public string tekNumKey;        // Number of Keys to Load (unused symbol, delete? - when have time to test!!)
      public WaveformHWIParams()
      {
        appName = "";
        cmdPrompt = "";
        password = "";
        passwordPrompt = "";
        loadtekCommand = "";
        tekFile = "AME"; //Initialized to AME
        tekNumKey = "1"; //Initialized to One Key
        tekOffset = "0"; //Initialized to No offset
        unitIDprompt = "";
        unitIDtrigger = "";
        brickPrompt = "";
        brickTrigger = "";
        resetCheckCmd = "";
        lightningWFChange = false;
                //Initialize Transmit ID
        keyingIDPrompt = new WaveformHWICommands();
        keyingIDPrompt.ArgumentHelp = new List<String>();
        keyingIDPrompt.ArgumentType = new List<String>();
        keyingIDPrompt.NumArguments = 0;
        keyingIDPrompt.type = "";
        //Set Key Order 
        keyOrder = new List<Int32>();
      }
    }
    public class WaveformHWICommandSlots
    {
      public string type;
      public List<WaveformHWICommands> commands;
      public WaveformHWICommandSlots()
      {
        commands = new List<WaveformHWICommands>();
      }
    }
    public class WaveformParams
    {
      public int wfidx; // index based off of BOLT's indicated wfmid
      public List<WaveformHWIParams> hwiParams; // Parameters used interfacing with waveforms
      public List<WaveformHWICommandSlots> commandSlot; // Commands sent post Tek Key load
      public List<BoltParameters> boltParams;
      public WaveformParams()
      {
        hwiParams = new List<WaveformHWIParams>(new WaveformHWIParams[4]);
        commandSlot = new List<WaveformHWICommandSlots>(new WaveformHWICommandSlots[4]);
        boltParams = new List<BoltParameters>(new BoltParameters[4]);
        for (int i = 0; i < 4; i++)
        {
          hwiParams[i] = (new WaveformHWIParams());
          boltParams[i] = (new BoltParameters());
          commandSlot[i] = (new WaveformHWICommandSlots());
        }
      }
      // Do this on waveform wipe
      public void ClearParameters()
      {
        commandSlot = new List<WaveformHWICommandSlots>(new WaveformHWICommandSlots[4]);
        hwiParams = new List<WaveformHWIParams>(new WaveformHWIParams[4]);
        boltParams = new List<BoltParameters>(new BoltParameters[4]);
      }
      public void ClearBoltParameters()
      {
        boltParams = new List<BoltParameters>(new BoltParameters[4]);
      }
      public bool LoadCommandsXml(string xmlAddress, int index)
      {
        string msg;
        string cap = " Configuration File Error ";
        MessageBoxButtons btn = MessageBoxButtons.OK;
        MessageBoxIcon icon = MessageBoxIcon.Asterisk;
        XmlTextReader xmlReader = null;
        WaveformHWIParams tempHwiparam = new WaveformHWIParams();
        commandSlot[index].commands.Clear();
        bool result = false;  // did we at least find a XML element ???
        try
        {
          xmlReader = new XmlTextReader(xmlAddress);
          xmlReader.ReadStartElement("HWIWaveformConfig");
          while (xmlReader.Read())
          {
            if (xmlReader.NodeType.Equals(XmlNodeType.Element))
            {
              /*
               * Read stuff to 'Load TEK Keys' for BOLT waveform slot
               * Not all LocalName items are required by every waveform
               */
              if (xmlReader.LocalName.Equals("name"))
              { // Short text descriptor to display in GUI 'Slot' textbox
                xmlReader.Read();
                tempHwiparam.appName = xmlReader.Value;
              }
              else if (xmlReader.LocalName.Equals("Waveform"))
              { // Expected text from BOLT IDing waveform application
                xmlReader.Read();
                tempHwiparam.title = xmlReader.Value;
              }
              else if (xmlReader.LocalName.Equals("passwordPrompt"))
              { // Expected text emitted by BOLT as a password prompt
                xmlReader.Read();
                tempHwiparam.passwordPrompt = xmlReader.Value;
              }
              else if (xmlReader.LocalName.Equals("password"))
              { // The wf specific password we send to BOLT
                xmlReader.Read();
                tempHwiparam.password = xmlReader.Value;
              }
              else if (xmlReader.LocalName.Equals("BrickNumber"))
              { // Expected label emitted by BOLT prior to <Brick Number>
                while (xmlReader.MoveToNextAttribute())
                {
                  if (xmlReader.LocalName.Equals("Prompt"))
                  {
                    tempHwiparam.brickPrompt = xmlReader.Value;
                  }
                  else if (xmlReader.LocalName.Equals("Trigger"))
                  {
                    tempHwiparam.brickTrigger = xmlReader.Value;
                  }
                }
              }
              else if (xmlReader.LocalName.Equals("UnitID"))
              { // Expected label emitted by BOLT prior to wf <Transmit ID> 
                while (xmlReader.MoveToNextAttribute())
                {
                  if (xmlReader.LocalName.Equals("Prompt"))
                  {
                    tempHwiparam.unitIDprompt = xmlReader.Value;
                  }
                  else if (xmlReader.LocalName.Equals("Trigger"))
                  {
                    tempHwiparam.unitIDtrigger = xmlReader.Value;
                  }
                }
              }
              else if (xmlReader.LocalName.Equals("prompt"))
              { // Expected text emitted by BOLT as a command line prompt
                xmlReader.Read();
                tempHwiparam.cmdPrompt = xmlReader.Value;
              }
              else if (xmlReader.LocalName.Equals("TEKLoad"))
              { // The wf specific command we use to send TEK to BOLT
                xmlReader.Read();
                tempHwiparam.loadtekCommand = xmlReader.Value;
              }
              else if (xmlReader.LocalName.Equals("TEKFile"))
              { // Name of local file in which TEKs are stored
                xmlReader.Read();
                tempHwiparam.tekFile = xmlReader.Value;
              }
              else if (xmlReader.LocalName.Equals("TEKOffset"))
              { // TEK file index offset from brick number (used to calculate guess at Keying ID)
                xmlReader.Read();
                tempHwiparam.tekOffset = xmlReader.Value;
              }
              else if (xmlReader.LocalName.Equals("KeyOrder"))
              { // Order in which to grab TEK from within list of TEKs for a given Keying ID
                xmlReader.Read();
                string[] order = xmlReader.Value.Split(',');
                for (int i = 0; i < order.Length; i++)
                {
                  tempHwiparam.keyOrder.Add(Int32.Parse(order[i]));
                }
              }
              else if(xmlReader.LocalName.Equals("KeyingID"))
              { // Stuff to allow user to input Keying ID as it doesn't always correlate to Brick Number + Std Offset
                tempHwiparam.keyingIDPrompt = new WaveformHWICommands();
                while (xmlReader.MoveToNextAttribute())
                {
                  if (xmlReader.LocalName.Equals("ArgumentHelp"))
                  {
                    tempHwiparam.keyingIDPrompt.ArgumentHelp.Add(xmlReader.Value);
                  }
                  else if (xmlReader.LocalName.Equals("ArgumentType"))
                  {
                    tempHwiparam.keyingIDPrompt.ArgumentType.Add(xmlReader.Value);
                  }
                  else if (xmlReader.LocalName.Equals("Type"))
                  {
                    tempHwiparam.keyingIDPrompt.type = xmlReader.Value;
                  }
                  else if (xmlReader.LocalName.Equals("NumArguments"))
                  {
                    tempHwiparam.keyingIDPrompt.NumArguments = Convert.ToInt32(xmlReader.Value);
                  }
                }
              }
              else if (xmlReader.LocalName.Equals("LightningWFChange"))
              { // Boolean flag indicating if Lightning, via BOLT User Port, commands change to next wf slot
                xmlReader.Read();
                tempHwiparam.lightningWFChange = true;
              }
              else if (xmlReader.LocalName.Equals("ResetCheckCmd"))
              { // ??? Suspect this was to do with sending BOLT Debug Port "reset" command but BOLT wf inconsistent wrt command
                xmlReader.Read();
                tempHwiparam.resetCheckCmd = xmlReader.Value;
              }

              /*
               * Read stuff to 'Configure BOLT' using Debug Port <system> or <waveform unique> commands for BOLT waveform slot
               * Not every waveform needs, or due to password restrictions can even accept, Debug Port commands from Red Tool
               */
              if (xmlReader.LocalName.Equals("Command"))
              {
                WaveformHWICommands tempHwicommands = new WaveformHWICommands();
                while (xmlReader.MoveToNextAttribute())
                {
                  if (xmlReader.LocalName.Equals("ArgumentHelp"))
                  {
                    tempHwicommands.ArgumentHelp.Add(xmlReader.Value);
                  }
                  else if (xmlReader.LocalName.Equals("ArgumentType"))
                  {
                    tempHwicommands.ArgumentType.Add(xmlReader.Value);
                  }
                  else if (xmlReader.LocalName.Equals("Type"))
                  {
                    tempHwicommands.type = xmlReader.Value;
                  }
                  else if (xmlReader.LocalName.Equals("NumArguments"))
                  {
                    tempHwicommands.NumArguments = Convert.ToInt32(xmlReader.Value);
                  }
                }
                commandSlot[index].commands.Add(tempHwicommands);
              }

              result = true; // an XML Element was found (but not neccessarily LocalName matched to be processed)
            } // end if XmlNodeType.Element
          } // end while xmlReader.Read()
          hwiParams[index] = tempHwiparam;
        }
        catch (FileNotFoundException)
        {
          cap = "ERROR";
          msg = "Configuration file\n"
              + xmlAddress
              + "\nnot found.";
          MessageBox.Show(msg, cap, btn, icon);
        }
        catch (XmlException ex)
        {
          cap = "ERROR: Bad format";
          msg = "The xml configuration file\n"
              + xmlAddress
              + "\nis not in the correct format.\n"
              + ex.Message;
          MessageBox.Show(msg, cap, btn, icon);
        }
        catch (Exception ex)
        {
          cap = "ERROR";
          msg = "Exception:\n" + ex.Message;
          MessageBox.Show(msg, cap, btn, icon);
        }
        finally
        {
          if (null != xmlReader)
          {
            xmlReader.Close();
          }
        }
        return result; // true if an XML element was is the file
      }
    }

    public class InfilGroupKey
    {
      public string index;
      public string GroupKey;
      public string GroupName;
      public string SponsorName;
      public string timestamp;
    }
    public class InfilKeyChain
    {
      public int numgrpkeys;
      public bool retrievedKeys;
      public int deviceID;
      public string uniqueKey;
      public List<InfilGroupKey> GroupKeys;
      public InfilKeyChain()
      {
        GroupKeys = new List<InfilGroupKey>();
        numgrpkeys = 0;
        retrievedKeys = false;
        deviceID = 0;
        uniqueKey = "";
      }
    }

    public struct ProgParams
    {
      // User Configurator Options
      public bool connected;
      public string deviceId;
      public string InfilKey1;
      public string InfilKey2;
      public string InfilKey3;
      public string InfilKey4;
      public string InfilKey5;
      public int    Selectedinfilkey;
      public string SelectedinfilkeyString;
      public string InfilNone;

      public string[,] waveformInfo;
      public int actOnWaveform;

      public bool infilNameack;   // Used by MakeHWIconnection in main thread and parseLoResponse in LoMsgThread
      public bool infilActiveack; // Used by MakeHWIconnection in main thread and parseLoResponse in LoMsgThread

      // Programming Tool Options
      public string cfgInfil1index;   // Index in KeyChain Document
      public string cfgInfil1Name;    // Key Name associated with key at index
      public int    cfgInfil1grpindx; // Index in group associations array
      public string cfgInfil2index;
      public string cfgInfil2Name;
      public int    cfgInfil2grpindx;
      public string cfgInfil3index;
      public string cfgInfil3Name;
      public int    cfgInfil3grpindx;
      public string cfgInfil4index;
      public string cfgInfil4Name;
      public int    cfgInfil4grpindx;
      public string cfgInfil5index;
      public string cfgInfil5Name;
      public int    cfgInfil5grpindx;
      public int    keyLoad;            // indicates for which key to BuildLoadInfilKey()
      // Command Menu
      public string cmdTextbox;         // cmdTxtBox.Text from Comms Log tab for use by LtMsgThread
      // Password
      public string PASSWORD;

      public void InitToHardCodedDefaults()
      { // Used on initialization or first power cycle to fill out user interface
        SelectedinfilkeyString = "0No Key";
        InfilNone = "0No Key";
        InfilKey1 = "1No Key Found";
        InfilKey2 = "2No Key Found";
        InfilKey3 = "3No Key Found";
        InfilKey4 = "4No Key Found";
        InfilKey5 = "5No Key Found";
        PASSWORD = "I 8 krypton!te";

        waveformInfo = new string[5, 5];
        actOnWaveform = 0;

        Selectedinfilkey = 0;
        connected = false;
        //
        keyLoad = 0;
      } // InitToHardCodedDefaults

      // Make sure returned information matches set information
      public bool InfilEquals(ProgParams other)
      {
        bool result = true;

        result &= (deviceId.Equals(other.deviceId));
        result &= (InfilKey1.Equals(other.InfilKey1) || (InfilKey1.Contains("NoKey") && other.InfilKey1.Contains("NoKey")));
        result &= (InfilKey2.Equals(other.InfilKey2) || (InfilKey2.Contains("NoKey") && other.InfilKey2.Contains("NoKey")));
        result &= (InfilKey3.Equals(other.InfilKey3) || (InfilKey3.Contains("NoKey") && other.InfilKey3.Contains("NoKey")));
        result &= (InfilKey4.Equals(other.InfilKey4) || (InfilKey4.Contains("NoKey") && other.InfilKey4.Contains("NoKey")));
        result &= (InfilKey5.Equals(other.InfilKey5) || (InfilKey5.Contains("NoKey") && other.InfilKey5.Contains("NoKey")));
        return result;
      } // end InfilEquals

      public void LoadOptions(ref ProgParams readValues)
      {
        deviceId = readValues.deviceId;
        InfilNone = "0No Key";
        InfilKey1 = readValues.InfilKey1;
        InfilKey2 = readValues.InfilKey2;
        InfilKey3 = readValues.InfilKey3;
        InfilKey4 = readValues.InfilKey4;
        InfilKey5 = readValues.InfilKey5;
        SelectedinfilkeyString = readValues.SelectedinfilkeyString;
        Selectedinfilkey = readValues.Selectedinfilkey;
        actOnWaveform = readValues.actOnWaveform;
        waveformInfo = readValues.waveformInfo;
      } // end LoadOptions

    } // end struct ProgParams

    public class ObfKeyChain
    {
      public int keyIdx;
      public List<string> keys;   // really C.OLT command parameters but treat like keys
      public ObfKeyChain()
      {
        keyIdx = 0;
        keys = new List<string>();
      }
    }
    const int numObfKeys = 62;
    public ObfKeyChain obfKeys = new ObfKeyChain(); // 

    public List<string> groupNames = new List<string>();
    public List<string> sponsorNames = new List<string>();
    public string[,] waveformConfigData;

    // Serial port 
    SerialPort comPort = new SerialPort();            // Used for all communications, new sets as COM1
                                                      // Thread Globals
    const bool HwIntfcEnable = false;                 // RtsEnable = false sets CMOS RTS/DTR to 3.3 VDC
    const bool HwIntfcDisable = true;                 // RtsEnable = true  sets CMOS RTS/DTR to 0.0 VDC
    LoRespStatus loRespStatus = LoRespStatus.Null;
    bool runMsgThreads = false;                       // boolean to keep msg threads running, set "false" to exit
    bool runDiagThread = false;                       // boolean to keep diag sniff & display thread running, set 'false' to exit
    ProgParams gpiParams = new ProgParams();          // parameters as displayed in the GUI
    ProgParams cfgParams = new ProgParams();          // parameters used by device
    WaveformParams waveforms = new WaveformParams();
    // Signals for the various threads - this allows us to make sure the threads will close
    private ManualResetEvent ltMsgEvent = new ManualResetEvent(false);  // start with false (non-signalled state)
    private ManualResetEvent loMsgEvent = new ManualResetEvent(false);  // start with false (non-signalled state)
    private ManualResetEvent diagEvent = new ManualResetEvent(false);   // start with false (non-signalled state)
    public LightningMsgType activeLtMsg;    // Message type
    private string appDir = Environment.CurrentDirectory;
    // Need to add Configuration save directory
    private string cfgDir;
    private string diagDir;
    private string exportDir;
    private string tekDir;
    private string portCfgFile;
    private string waveformCfgFile;
    private string currentPort;
    public string boltDbg;
    public string LightningDiag;
    // Define the timer 
    public System.Timers.Timer boltTimer;

    Image CheckMrkimg;
    Image minusImg;
    Image xmarkImg;

    public InfilKeyChain infilKeys;

    enum BackgroundUpdate                   // send via ProgressChangedEventArgs.ProgressPercentage
    {
      Null,
      MsgThreadFatality,
      StartLoRespTimer,
      DisableLoRespTimer,
      WriteLoToRtb,
      WriteLtToRtb,
      ShowModelessInsight,
      ShowModelessNotification,
    }

    enum InfoType
    {
      Insight,
      Notification,
    }

    // Lightning Originated Response
    enum LoRespStatus
    {
      Null,               // nothing expected or received - lowest value
      GAKreq,                 // required 
      GKNreq,                 // required 
//    WFSreq,// (as of 2022/05/01 this no longer used except in dead code)
//    WFNreq,// (as of 2022/05/01 this no longer used except in dead code)
      SWFIreq,
      CWFIreq,
      CMDreq,
      GLKreq,
      GOKreq,                 // required
      OLTreq,
      SWFreq,                 // required 
      GZKreq,
      ECPreq,                 // 
      ATE1req,
      ATE0req
    }

    #endregion DECLARATIONS

    private void GetProgParamsFromGpi(ref ProgParams progParams)
    {
      /*
       * Using the values/states of the forms controls relating to programmable parameters,
       * set the programmable parameters within provided ProgParams structure.
       * Must set normalRate before motionRate and movingRate as it may be used to set them.
       */
      if (noKey0.Checked)
      {
        progParams.Selectedinfilkey = 0;
        progParams.SelectedinfilkeyString = "No Key";
      }
      else if (infilKey1.Checked)
      {
        progParams.Selectedinfilkey = 1;
        progParams.SelectedinfilkeyString = progParams.InfilKey1.ToString();
      }
      else if (infilKey2.Checked)
      {
        progParams.Selectedinfilkey = 2;
        progParams.SelectedinfilkeyString = progParams.InfilKey2.ToString();
      }
      else if (infilKey3.Checked)
      {
        progParams.Selectedinfilkey = 3;
        progParams.SelectedinfilkeyString = progParams.InfilKey3.ToString();
      }
      else if (infilKey4.Checked)
      {
        progParams.Selectedinfilkey = 4;
        progParams.SelectedinfilkeyString = progParams.InfilKey4.ToString();
      }
      else if (infilKey5.Checked)
      {
        progParams.Selectedinfilkey = 5;
        progParams.SelectedinfilkeyString = progParams.InfilKey5.ToString();
      }
      else
      {
        progParams.Selectedinfilkey = 0;
        progParams.SelectedinfilkeyString = "0None";
      }
   } // end method GetProgParamsFromGpi

    private void SetGpiFromProgParams(ref ProgParams gpiParams)
    {
      infilgrpBox.Enabled = true;
      noKey0.Enabled = true;
      infilKey1.Enabled = true;
      infilKey1.Text = gpiParams.InfilKey1;
      infilKey2.Enabled = true;
      infilKey2.Text = gpiParams.InfilKey2;
      infilKey3.Enabled = true;
      infilKey3.Text = gpiParams.InfilKey3;
      infilKey4.Enabled = true;
      infilKey4.Text = gpiParams.InfilKey4;
      infilKey5.Enabled = true;
      infilKey5.Text = gpiParams.InfilKey5;
      devInfilUniqueIdBox.Text = gpiParams.deviceId;
      configureInfilButton.Enabled = true;
      if (gpiParams.SelectedinfilkeyString == gpiParams.InfilNone)
      {
        noKey0.Checked = true;
        gpiParams.Selectedinfilkey = 0;
      }
      else if (gpiParams.SelectedinfilkeyString == gpiParams.InfilKey1)
      {
        infilKey1.Checked = true;
        gpiParams.Selectedinfilkey = 1;
      }
      else if (gpiParams.SelectedinfilkeyString == gpiParams.InfilKey2)
      {
        infilKey2.Checked = true;
        gpiParams.Selectedinfilkey = 2;
      }
      else if (gpiParams.SelectedinfilkeyString == gpiParams.InfilKey3)
      {
        infilKey3.Checked = true;
        gpiParams.Selectedinfilkey = 3;
      }
      else if (gpiParams.SelectedinfilkeyString == gpiParams.InfilKey4)
      {
        infilKey4.Checked = true;
        gpiParams.Selectedinfilkey = 4;
      }
      else if (gpiParams.SelectedinfilkeyString == gpiParams.InfilKey5)
      {
        infilKey5.Checked = true;
        gpiParams.Selectedinfilkey = 5;
      }

      //-----------------WaveFormInfo Config Data------------------------
      string data;
      List<string[]> splitData = new List<string[]>();
      try
      {
        if (File.Exists(waveformCfgFile))
        {                                             // Read waveform family/name options from file
          using (StreamReader sr = new StreamReader(waveformCfgFile))
          {
            while ((data = sr.ReadLine()) != null)
            {
              string[] split = data.Split(' ');       // family_enum, family_name, wf_name(s)
              Int32 splitLen = split.GetLength(0);
              char  famChar  = data[0];               // see Lightning End User Software ICD, C.WFI : p2
              int   famEnum   = 0;                    // famChar enum converted to base62 number{0,...,61}

              if (splitLen > 1)
              {
                if (Char.IsDigit(famChar))        {   // '0' - '9'
                  famEnum = famChar - 48;         }   //            0,...,9
                else if (Char.IsUpper(famChar))   {   // 'A' - 'Z'
                  famEnum = famChar - 55;         }   //           10,...,35
                else if (Char.IsLower(famChar))   {   // 'a' - 'z'
                  famEnum = famChar - 61;         }   //           36,...,61
              }

              if ((splitLen > 1) &&                                 // ensure min of family_enum & family_name
                  (famEnum == splitData.Count))                   { // ensure family_enum in order: 0,...
                splitData.Add(split);                             }
              else                                                {
                throw new Exception("ERROR: file corrupted");     }
            }
          }

          resetWfInfoTab();
          for (int i = 0; i < splitData.Count; i++)
          {
            waveformConfigData[i, 0] = splitData[i][1];
            w0p2.Items.Add(splitData[i][1]);
            w1p2.Items.Add(splitData[i][1]);
            w2p2.Items.Add(splitData[i][1]);
            w3p2.Items.Add(splitData[i][1]);
            w4p2.Items.Add(splitData[i][1]);
            if (splitData[i][0] != "0")
            {                                                       // When not dealing with "No Waveform"
              for (int n = 2; n < splitData[i].Length; n++)       { // copy all pertinent waveform names
                waveformConfigData[i, n - 1] = splitData[i][n];   } // read from file into what will be used
            }                                                       // to populate WF Info tab's dropdown boxes
          }
        }
        else
        {
          string cap = "ERROR: MISSING FILE";
          string msg = "ERROR: Did not find required file\n" + waveformCfgFile;
          ShowModelessMessageBox(msg, cap, InfoType.Notification);
        }

        bool nonNull = true;
        //Ensure gpiParams for waveformInfo are not Null
        for (int wfParam = 1; wfParam < 5 && nonNull; wfParam++)
        { // wfParam == 0 is the fixed slot number {0,...,4}; wfParam == 1 is Family attribute,..., wfParam == 4 is Name attribute
          for (int wfSlot = 0; wfSlot < 5 && nonNull; wfSlot++)
          {
            if (gpiParams.waveformInfo[wfSlot, wfParam] == null)
            {
              nonNull = false;  // this is the case when have yet to 'Query LBHH' for WF Info
            }
          }
        }

        //// Waveform Configuration Tab
        if (nonNull)
        {
          int index = -1;
          char wc = '0';  // waveform class (a base62hex enumeration using characters 0-9,A-Z,a-z)
          //// w0
          wc = gpiParams.waveformInfo[0, 1][0];
          if (Char.IsLetterOrDigit(wc))
          {
            if      (wc <= '9')  { index = (int)(wc - '0');      }
            else if (wc <= 'Z')  { index = (int)(wc - 'A' + 10); }
            else if (wc <= 'z')  { index = (int)(wc - 'a' + 36); }
            if (w0p2.Items.Count > index) {
              w0p2.SelectedIndex = index; }               // 'Family'
            w0p3.Text = gpiParams.waveformInfo[0, 2];     // 'Keying ID'
            w0p4.Text = gpiParams.waveformInfo[0, 3];     // 'Transmit ID'
            w0p5.Text = gpiParams.waveformInfo[0, 4];     // 'WF Name'
          }
          //// w1
          wc = gpiParams.waveformInfo[1, 1][0];
          if (Char.IsLetterOrDigit(wc))
          {
            if      (wc <= '9')  { index = (int)(wc - '0');      }
            else if (wc <= 'Z')  { index = (int)(wc - 'A' + 10); }
            else if (wc <= 'z')  { index = (int)(wc - 'a' + 36); }
            if (w1p2.Items.Count > index) {
              w1p2.SelectedIndex = index; }
            w1p3.Text = gpiParams.waveformInfo[1, 2];
            w1p4.Text = gpiParams.waveformInfo[1, 3];
            w1p5.Text = gpiParams.waveformInfo[1, 4];
          }
          //// w2
          wc = gpiParams.waveformInfo[2, 1][0];
          if (Char.IsLetterOrDigit(wc))
          {
            if      (wc <= '9')  { index = (int)(wc - '0');      }
            else if (wc <= 'Z')  { index = (int)(wc - 'A' + 10); }
            else if (wc <= 'z')  { index = (int)(wc - 'a' + 36); }
            if (w2p2.Items.Count > index) {
              w2p2.SelectedIndex = index; }
            w2p3.Text = gpiParams.waveformInfo[2, 2];
            w2p4.Text = gpiParams.waveformInfo[2, 3];
            w2p5.Text = gpiParams.waveformInfo[2, 4];
          }
          //// w3
          wc = gpiParams.waveformInfo[3, 1][0];
          if (Char.IsLetterOrDigit(wc))
          {
            if      (wc <= '9')  { index = (int)(wc - '0');      }
            else if (wc <= 'Z')  { index = (int)(wc - 'A' + 10); }
            else if (wc <= 'z')  { index = (int)(wc - 'a' + 36); }
            if (w3p2.Items.Count > index) {
              w3p2.SelectedIndex = index; }
            w3p3.Text = gpiParams.waveformInfo[3, 2];
            w3p4.Text = gpiParams.waveformInfo[3, 3];
            w3p5.Text = gpiParams.waveformInfo[3, 4];
          }
          //// w4
          wc = gpiParams.waveformInfo[4, 1][0];
          if (Char.IsLetterOrDigit(wc))
          {
            if      (wc <= '9')  { index = (int)(wc - '0');      }
            else if (wc <= 'Z')  { index = (int)(wc - 'A' + 10); }
            else if (wc <= 'z')  { index = (int)(wc - 'a' + 36); }
            if (w4p2.Items.Count > index) {
              w4p2.SelectedIndex = index; }
            w4p3.Text = gpiParams.waveformInfo[4, 2];
            w4p4.Text = gpiParams.waveformInfo[4, 3];
            w4p5.Text = gpiParams.waveformInfo[4, 4];
          }
        }
      }
      catch (Exception)
      {
      }
    } // end SetGpiFromProgParams

    public LBHH_HWI()
    { //Initialize the GUI and various dependencies (e.g. directories)
      InitializeComponent();
      CheckMrkimg = (Image)Properties.Resources.ResourceManager.GetObject("CheckMark");
      minusImg = (Image)Properties.Resources.ResourceManager.GetObject("minusImage");
      xmarkImg = (Image)Properties.Resources.ResourceManager.GetObject("xmark");
      lightningStatus.Image = minusImg;
      infilKeys = new InfilKeyChain();
      // TODO - there has to be a better way than hardcoding [62, 16] but to do other will likely cause a ripple of changes I don't have time for right now!
      waveformConfigData = new string[62, 16];    // up to 62 waveform class/family IDs having 1 class and up to 15 waveform names per class read from waveforms.txt
      cfgDir = appDir + ConfigFolder;
      diagDir = appDir + DiagnosticsFolder;
      exportDir = appDir + ExportFolder;
      tekDir = appDir + TekFolder;
      portCfgFile = cfgDir + "\\port.txt";
      waveformCfgFile = cfgDir + "\\waveforms.txt";

      if ( ! Directory.Exists(cfgDir))
      {
        Directory.CreateDirectory(cfgDir);
      }
      if ( ! Directory.Exists(diagDir))
      {
        Directory.CreateDirectory(diagDir);
      }
      if ( ! Directory.Exists(exportDir))
      {
        Directory.CreateDirectory(exportDir);
      }
      if ( ! Directory.Exists(tekDir))
      {
        Directory.CreateDirectory(tekDir);
      }

      { // Note when, where, and by whom this app was started
        String wwwStr = System.DateTime.UtcNow.ToString("u") + "\n"
                      + Environment.CommandLine + "\nstarted by "
                      + Environment.UserName + "\n\n";
        WriteRichTextBox(diagRichTextBox, wwwStr);
      }

      //--------------------------------------------------
      DisconnectedTitle = this.Text;
      ConnectedTitle = this.Text.Remove(Text.IndexOf("Tool")) + "- CONNECTED : ";
      //diagRichTextBox.MaxLength = 16777216;  // Limit size to 16 MB; 1-2 days worth
      PopulateComPortComboBox(comPortComboBox);

      //-----------------Port Config Data------------------------
      try
      {
        if (File.Exists(portCfgFile))
        {
          using (StreamReader sr = new StreamReader(portCfgFile))
          {
            currentPort = sr.ReadLine();
          }
          comPortComboBox.Text = currentPort;  // this will set .SelectedItem and .SelectedIndex if port listed in ComboBox
        }
      }
      catch (Exception)
      {

      }
    } // end LBHH_HWI

    private bool ReadInfilKeyFile()
    { // Read password protected Excel file that contains INFIL and OBFUSCATION key information and init related GUI controls
      bool isOk = false;

      String keyCombo1Txt = keyCombo1.Text;
      String keyCombo2Txt = keyCombo2.Text;
      String keyCombo3Txt = keyCombo3.Text;
      String keyCombo4Txt = keyCombo4.Text;
      String keyCombo5Txt = keyCombo5.Text;
      //Pull Device ID, all allocated Group Keys, and C.OLT obfuscation commands
      if (GetInfilKeys(0, ref infilKeys, ref obfKeys.keys))
      {
        if (0 > uniqueIDupdown.Value) { uniqueIDupdown.Value = 0; }
        keyCombo1.Items.Clear();
        keyCombo2.Items.Clear();
        keyCombo3.Items.Clear();
        keyCombo4.Items.Clear();
        keyCombo5.Items.Clear();
        keyCombo1.Items.Add("None");
        keyCombo2.Items.Add("None");
        keyCombo3.Items.Add("None");
        keyCombo4.Items.Add("None");
        keyCombo5.Items.Add("None");
        keyCombo1.SelectedIndex = 0;
        keyCombo2.SelectedIndex = 0;
        keyCombo3.SelectedIndex = 0;
        keyCombo4.SelectedIndex = 0;
        keyCombo5.SelectedIndex = 0;
        //Application found key table and filled GPI
        for (int i = 0; i < infilKeys.GroupKeys.Count; i++)
        {
          if (infilKeys.GroupKeys[i].GroupName != "")
          {
            keyCombo1.Items.Add(infilKeys.GroupKeys[i].GroupName);
            keyCombo2.Items.Add(infilKeys.GroupKeys[i].GroupName);
            keyCombo3.Items.Add(infilKeys.GroupKeys[i].GroupName);
            keyCombo4.Items.Add(infilKeys.GroupKeys[i].GroupName);
            keyCombo5.Items.Add(infilKeys.GroupKeys[i].GroupName);
            if (infilKeys.GroupKeys[i].GroupName == keyCombo1Txt) { gpiParams.cfgInfil1grpindx = i + 1; keyCombo1.Text = keyCombo1Txt; }
            if (infilKeys.GroupKeys[i].GroupName == keyCombo2Txt) { gpiParams.cfgInfil2grpindx = i + 1; keyCombo2.Text = keyCombo2Txt; }
            if (infilKeys.GroupKeys[i].GroupName == keyCombo3Txt) { gpiParams.cfgInfil3grpindx = i + 1; keyCombo3.Text = keyCombo3Txt; }
            if (infilKeys.GroupKeys[i].GroupName == keyCombo4Txt) { gpiParams.cfgInfil4grpindx = i + 1; keyCombo4.Text = keyCombo4Txt; }
            if (infilKeys.GroupKeys[i].GroupName == keyCombo5Txt) { gpiParams.cfgInfil5grpindx = i + 1; keyCombo5.Text = keyCombo5Txt; }
          }
        }
        isOk = true;
      }

      return isOk;
    } // end ReadInfilKeyFile

    private bool MakeHwiPrgmConnection(string portName)
    { // Assume one LBHH plugged in at a time; connect to the first thing that responds appropriately to queries
      var rm = new System.Resources.ResourceManager(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name + ".Properties.Resources", System.Reflection.Assembly.GetExecutingAssembly());
      Image img = (Bitmap)rm.GetObject("CheckMark.jpg");
      lightningStatus.Image = minusImg;
      cfgParams.connected = false;
      
      if (comPort.IsOpen || (portName != comPort.PortName))
      {                                         // Changing ports, so disconnect from connected port
        StopDiagThread();                       //No need to run diag while connecting
        StopMessagingThreads();                 // Shutdown messaging threads before changing port
        Thread.Sleep(200);                      // Give device more time to see 'Disconnect'
        Application.DoEvents();                 // and process any message request/user inputs
        comPort.Close();                        // before closing port (which unfortunately
      }                                         // sets RTS back to HwIntfcEnabled)
      if ( ! comPort.IsOpen)                   {
        OpenLightningComPortProperty(portName);}// when just now opening the COM port, fully configure it

      //Use the method below for finding device
      loRespStatus = (cfgParams.connected)      // If a LBHH is already connected
                   ? LoRespStatus.Null          // then no LO msg expected (interface is open)
                   : LoRespStatus.GAKreq;       // otherwise we are looking for an LO ACK msg
      runMsgThreads = true;                     // Set flag so message threads will run
      if ( ! ltBackgroundWorker.IsBusy)       { // Send LT message thread shouldn't be running
        ltBackgroundWorker.RunWorkerAsync();  } // so start it now (if not already running)
      Thread.Sleep(0);                          // Give LT background worker thread time to startup
      Application.DoEvents();                   // then process events that may have occurred
      if ( ! loBackgroundWorker.IsBusy)       { // Receive LO message thread shouldn't be running
        loBackgroundWorker.RunWorkerAsync();  } // so start it now (if not already running)
      Thread.Sleep(20);                         // Give background worker threads time to startup
      Application.DoEvents();                   // then process events that may have occurred

      int countDown = 102;                      // 102 counts of 1/20 second ea. yields 5.1 sec to connect
      bool hasMoreTime = true;                  // and must accomplish two query tasks within that time
      cfgParams.infilNameack = false;           // Clear flag to break while loop after
      QueueLtMessage(LightningMsgType.qKeyNames);// queue Lightning Terminated query of LBHH infil key names
      while (( ! cfgParams.infilNameack) && hasMoreTime)
      {                                         // Wait for a query response which proves connection
        loMsgEvent.Set();                       // Signal 'Receive LO Messages' thread to run
        Thread.Sleep(50);                       // and delay at least 1/10 second each
        Application.DoEvents();                 // time the 'while loop' is executed, then
        countDown--;                            // decrement the countdown and hope we don't
        if (0 >= countDown)                   { // use it all during this first query but if so
            hasMoreTime = false;              } // note that we have idled away all our time
      }                                         // Getting a response will set cfgParam.connected
      if (cfgParams.connected && hasMoreTime)
      {                                         // When 1st query-response proved timely HWI connection
        cfgParams.connected = false;            // clear connected flag as way to enforce 2nd response and
        cfgParams.infilActiveack = false;       // clear flag to break from while loop after
        QueueLtMessage(LightningMsgType.qActvKeyNames);  // queue second query (for active infil keys)
        while (( ! cfgParams.infilActiveack) && hasMoreTime)
        {                                       // Wait for a query response which proves connection
          loMsgEvent.Set();                     // Signal 'Receive LO Messages' thread to run
          Thread.Sleep(50);                     // and delay at least 1/10 second each
          Application.DoEvents();               // time the 'while loop' is executed, then
          countDown--;                          // decrement the countdown and hope we don't
          if (0 >= countDown)                 { // use it all during this second query but if so
              hasMoreTime = false;            } // note that we have idled away all our time
        }                                       // Getting a response will set cfgParam.connected
      }
      return cfgParams.connected;
    } // end MakeHwiPrgmConnection

    private void DisconnectDevice()
    { // Normally called in the event of a failure, when LBHH to be pwr cycled, or when done with a LBHH
      connectButton.Text = "Connect";
      comPortComboBox.Enabled = true;
      connectButton.Refresh();
      this.Text = DisconnectedTitle; // Set window/form title bar to its default
      cfgParams.connected = false;

      StopMessagingThreads();               // Ensure shutdown of messaging threads
      infilKeysGrpBox.Enabled = false;
      boltSetupGrpBox.Enabled = false;

      ltngMsgStartLbl.Enabled = false;
      cmdTxtBox.Enabled = false;
      ltngMsgStopLbl.Enabled = false;
      sendCmdBtn.Enabled = false;
      connectBoltBtn.Enabled = false;

      lightningWaveformConfigurationGroup.Enabled = false;
      resetConfigTabs();
      antiTamperBox.Enabled = false;
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
          //WriteRichTextBox(diagRichTextBox, (System.DateTime.UtcNow.ToString("HH':'mm':'ss.fff") + " "
          // + ex.ToString() + "\n"));
        }
      }
      Monitor.Exit(comPort);                // Ensure we release the lock on accessing COM port
      comPort.Dispose();
      comPort = new SerialPort();           // Ensure comPort is left neither connected nor NULL
      infilgrpBox.Enabled = false;          // Kill infil group box
      configureInfilButton.Enabled = false;
      connectButton.Enabled = true;
      lightningStatus.Image = minusImg;
      gpiParams.InitToHardCodedDefaults();
      cfgParams.InitToHardCodedDefaults();
    } // end DisconnectDevice

    //Name: GetBoltKeys
    //Arguments: string keyIndex - index into TEK file, often related to Brick Number or TX ID
    //           string fileLocation - Tek File Location
    //Output: string key - Device Key to be sent to the BOLT
    //Description: Gets TEK Key needed for device
    private BoltKey GetBoltKeys(string keyIndex, string tekFile)
    {
      BoltKeyChain boltKeyChain = new BoltKeyChain();
      BoltKey tempboltKey = new BoltKey();
      List<string> tekFileList = new List<string>();
      string[] files = Directory.GetFiles(tekDir, "*.*", SearchOption.TopDirectoryOnly);
      foreach (string file in files)
      {
        if (Path.GetExtension(file) == ".tek")
        {
          if (file != "" && file.Replace(tekDir, "").Replace("\\", "") == tekFile)
          {
            XmlTextReader reader = new XmlTextReader(file);
            string tempfamily = "";
            int activeKey = 0;
            if (reader.IsStartElement())
            {
              while (reader.Read())
              {
                switch (reader.NodeType)
                {
                  case XmlNodeType.Element:
                    if (reader.Name == "Family")
                    {
                      tempboltKey.familyOpen = true;
                    }
                    if (reader.Name == "Device")
                    {
                      tempboltKey = new BoltKey();
                      tempboltKey.deviceOpen = true;
                      tempboltKey.family = tempfamily;
                    }
                    else if (reader.Name == "ID")
                    {
                      tempboltKey.idOpen = true;
                    }
                    else if (reader.Name.Contains("TEK_"))
                    {
                      tempboltKey.keyOpen = true;
                      activeKey = Int32.Parse(reader.Name.Split('_')[1]);
                    }
                    break;
                  case XmlNodeType.Text:
                    if (tempboltKey.idOpen)
                    {
                      tempboltKey.id = reader.Value;
                    }
                    else if (tempboltKey.keyOpen)
                    {
                      if (activeKey < 10)                                 {
                        tempboltKey.key[activeKey - 1] = reader.Value;    }
                    }
                    else if (tempboltKey.familyOpen)
                    {
                      tempfamily = reader.Value;
                    }
                    break;
                  case XmlNodeType.EndElement:
                    if (reader.Name == "Device")
                    { //Load into list
                      boltKeyChain.boltKeys.Add(tempboltKey);
                      tempboltKey.deviceOpen = false;
                    }
                    else if (reader.Name == "ID")
                    {
                      tempboltKey.idOpen = false;
                    }
                    else if (reader.Name.Contains("TEK_"))
                    {
                      tempboltKey.keyOpen = false;
                    }
                    else if (reader.Name == "Family")
                    {
                      tempboltKey.familyOpen = false;
                    }
                    break;
                }
              }
            }
            //go through and pull data we want
            for (int i = 0; i < boltKeyChain.boltKeys.Count; i++)
            {
              if (boltKeyChain.boltKeys[i].id == int.Parse(keyIndex).ToString())
              {
                if (boltKeyChain.boltKeys[i].key[0] != "")    {
                  return boltKeyChain.boltKeys[i];            }
              }
            }
          }
        }
      }
      return null;
    } // end GetBoltKeys

    //Name: SendBoltCommandSlowly
    //Arguments: string command- string to send 
    //Description: Slowly send commands to Bolt (the firmware must be polling characters)
    private void SendBoltCommandSlowly(string command, bool maskInRtb)
    {
      command = command + "\r\n";
      if (maskInRtb)                                                                                                      {
        WriteRichTextBox(diagRichTextBox, "\r\n" + DateTime.UtcNow.ToString("HH':'mm':'ss.fff") + "-> CMD: ...\r\n");     }
      else                                                                                                                {
        WriteRichTextBox(diagRichTextBox, "\r\n" + DateTime.UtcNow.ToString("HH':'mm':'ss.fff") + "-> CMD: " + command);  }
      for (int k = 0; k < (command.ToCharArray()).Length; k++)
      {
        string temp = command[k].ToString();
        comPort.Write(temp.ToCharArray(), 0, (temp.ToCharArray()).Length);
        Thread.Sleep(5);  // Inter-Character delay so BOLT will capture UART data
      }
      Thread.Sleep(100);  // Line delay so BOLT can process command string
    } // end SendBoltCommandSlowly

    //Name: SendLightningCommands   TODO - 2022/05/01 Get rid of this and use QueueLtMessage() so that Main GUI thread is not blocked
    //Arguments: string command- string to send 
    //Description: Sends commands to the lightning
    private void SendLightningCommands(string command)
    {
      for (int k = 0; k < (command.ToCharArray()).Length; k++)
      {
        string temp = command[k].ToString();
        comPort.Write(temp.ToCharArray(), 0, (temp.ToCharArray()).Length);
        Thread.Sleep(100);
      }
    } // end SendLightningCommands


    //Name: MakeLightningSwitchWaveform
    //Arguments: int index
    //Output: Boolean - indicates if action successfully switched waveforms O.SWF
    //Description: Responsible for changing Bolt Waveforms via the lightning
    public bool MakeLightningSwitchWaveform(int waveForm, string portName)
    {
      ltngMsgStartLbl.Visible = true;      //Indicate to the user that the we are no longer talking to the lightning
      ltngMsgStopLbl.Visible = true;       //Indicate to the user that the we are no longer talking to the lightning
      gpiParams.actOnWaveform = waveForm;
      QueueLtMessage(LightningMsgType.sWaveForm);                 // Request Lightning command the BOLT to wf 
      boltTimer = new System.Timers.Timer(3000);                  // Setup a new 3 second timer
      boltTimer.AutoReset = false;                                // for a single shot
      boltTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent); // Hook up the Elapsed event for the timer.
      boltTimer.Stop();                                           // [Re]Start the timer
      boltTimer.Start();
      boltTimer.Enabled = true;
      while (boltTimer.Enabled)
      {                                                           // While we have not timed out
        Thread.Sleep(1000);                                       // delay at least 1 second each time
        Application.DoEvents();                                   // the 'while loop' is executed, then
        if (LightningDiag.Contains("+C.SWF"))                     // check data from Lightning.
        {                                                         // When got command ACK from Lightning
        //boltTimer.Enabled = false;                              // turn off the timer and // TODO test if this 'good idea' breaks anything
          return true;                                            // return with indication of SUCCESS
        }
        QueueLtMessage(LightningMsgType.sWaveForm);               // Re-request Lightning command the BOLT to wf if no ACK yet
      }
      return false;                                               // Timeout occurred; return indicating error
    } // end MakeLightningSwitchWaveform

    //Name: ConnectToBoltDebugPort
    //Arguments: object sender
    //           EventArgs e
    //Output: Boolean - indicates if action occurred successfully
    //Description: Request Lightning to expose BOLT Debug Port thru COMx connection to which this program is connected
    public void ConnectToBoltDebugPort()
    { //Make sure no operations are in progress
      //Tell lightning to open Bolt
      QueueLtMessage(LightningMsgType.sCOMxBoltDbg);// Send Lightning command to expose the Bolt Debug port
      infilKeysGrpBox.Enabled = false;              // Disable all lightning options for the gui is no longer communicating with it
      ltngMsgStartLbl.Visible = false;              // Indicate to the user that the we are no longer talking to the lightning
      ltngMsgStopLbl.Visible = false;               // Indicate to the user that the we are no longer talking to the lightning
      Thread.Sleep(500);                            // Delay to let Red Tool send Set COMx and the LBHH start to act on it
      if (comPort.IsOpen)
      {                                             // Changing ports, so disconnect from connected port
        StopDiagThread();                           // No need to run diag while connecting (avoid garbage)
        StopMessagingThreads();                     // Shutdown messaging threads before changing port
        Thread.Sleep(200);                          // Give device more time to see 'Disconnect'
        Application.DoEvents();                     // and process any message request/user inputs
      }
      boltDbg = "";                                 // Reinitialize BOLT Debug stream data
      OpenBoltComPort(comPort.PortName);            // Reconfigure COM port for BOLT Debug Port communications
      if ( ! runDiagThread)
      {                                             // If diag thread not running,
        runDiagThread = true;                       // ensure it will and then
        diagBackgroundWorker.RunWorkerAsync();      // kick it off
      }
      Thread.Sleep(0);                              // This can probably be removed
      Application.DoEvents();                       // then process events that may have occurred
    } // end ConnectToBoltDebugPort

    //Name: GetWfmidFromBolt
    //Arguments: object sender
    //           EventArgs e
    //Output: Boolean - indicates if action occurred successfully
    //Description: Responsible for retrieving the Bolts current waveform slot 
    public bool GetWfmidFromBolt()
    {
      if (boltTimer != null)
      {
        boltTimer.Stop();
      }
      Application.DoEvents();
      boltTimer = new System.Timers.Timer(30000);                 // Setup a new 30-second timer
      boltTimer.AutoReset = false;                                // for a single shot
      boltTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent); // Hook up the Elapsed event for the timer.
      boltTimer.Stop();                                           // [Re]Start the timer
      boltTimer.Start();
      boltTimer.Enabled = true;
      while (boltTimer.Enabled)                                   // While we have not timed out
      {
        Thread.Sleep(100);                                        // and delay at least 1/10 second
        Application.DoEvents();                                   // each time the 'while loop' is executed, then
        if (boltDbg.Contains("WFMID"))                            // Search for identifier in BOLT configuration printout
        {
          string[] parsedDiag = boltDbg.Split('\n');
          for (int i = 0; i < parsedDiag.Length; i++)             // Find Line where WFMID is mentioned
          {
            if (parsedDiag[i].Contains("WFMID"))
            {                                                     // Extract the waveform ID/Slot
              int wfmid = Convert.ToInt16((parsedDiag[i].Split(' ')[1]).Replace("\r", ""));
              waveforms.wfidx = wfmid - 1;                        // Indexes current selected waveform info based on BOLT reported wfmid
              waveformtxtbox.Text = wfmid.ToString();
              boltTimer.Enabled = false;
              if (waveforms.hwiParams[waveforms.wfidx].appName != "")
              {             
                return true;                                      // Return SUCCESS; XML is loaded for waveform slot
              }
              else
              {
                string msg = "ERROR: No XML loaded for waveform slot\nPlease load compatible XML!";
                string cap = "ERROR";
                MessageBoxButtons btn = MessageBoxButtons.OK;
                MessageBoxIcon icon = MessageBoxIcon.Asterisk;
                MessageBox.Show(msg, cap, btn, icon);
                return false;                                     // Return error
              }
            }
          }
        }
      }
      return false;                                               // Return error; timeout occurred
    } // end GetWfmidFromBolt

    //Name: HandlePassword
    //Arguments: object sender
    //           EventArgs e
    //Output: Boolean - indicates if action occurred successfully
    //Description: The Password Phase of the bolt programming process
    //basically the Bolt has a password to load
    //the TEK keys and change device parameters
    //some waveforms do not have passwords though and this step can be skipped 
    public bool HandlePassword()
    {
      if ("" == waveforms.hwiParams[waveforms.wfidx].password || "" == waveforms.hwiParams[waveforms.wfidx].passwordPrompt)
      { // When XML file did not provide info needed to send waveform password there is nothing to do
        return true;                                             // Return indicating success by lack of requirement
      }
      if (boltTimer != null)    {
        boltTimer.Stop();       }
      Application.DoEvents();
      boltTimer = new System.Timers.Timer(30000);                 // Setup a new 30-second timer
      boltTimer.AutoReset = false;                                // for a single shot
      boltTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent); // Hook up the Elapsed event for the timer.
      boltTimer.Stop();                                           // [Re]Start the timer
      boltTimer.Start();
      boltTimer.Enabled = true;
      bool passwordSent = false;                                  // We must act based on initially not having sent pswd
      while (boltTimer.Enabled)                                   // While we have not timed out
      {
        Thread.Sleep(10);                                         // delay at least 1/100 second each time
        Application.DoEvents();                                   // the 'while loop' is executed, then
        if (( ! passwordSent) && boltDbg.Contains(waveforms.hwiParams[waveforms.wfidx].passwordPrompt))
        {                                                         // When found password prompt described in XML file
          SendBoltCommandSlowly(waveforms.hwiParams[waveforms.wfidx].password, true); // Immediately send password to BOLT -- time critical!!
          passwordSent = true;                                    // Note that password was sent so immediately start to
        }                                                         // look for BOLT command line prompt described in XML file
        else if (passwordSent && boltDbg.Contains(waveforms.hwiParams[waveforms.wfidx].cmdPrompt))
        {                                                         // When BOLT emitted its wf specific command line prompt
          boltTimer.Enabled = false;                              // turn off the timer
          return true;                                            // and return with indication of SUCCESS
        }
      }
      return false;                                               // Timeout occurred; return indicating error
    } // end HandlePassword

    //Name: GetUnitID
    //Arguments: object sender
    //           EventArgs e
    //Output: Boolean - indicates userID was recovered successfully (or not required of the waveform)
    //Description: Gets the userID (helpful for programmatics)
    public bool GetUnitID()
    {
      if ("" == waveforms.hwiParams[waveforms.wfidx].unitIDprompt || "" == waveforms.hwiParams[waveforms.wfidx].unitIDtrigger)
      { // When XML file did not provide info needed to extract unit TX ID there is nothing to do
        return true;                                             // Return indicating success by lack of requirement
      }
      if (boltTimer != null)    {
        boltTimer.Stop();       }
      Application.DoEvents();
      boltTimer = new System.Timers.Timer(30000);                 // Setup a new 30-second timer
      boltTimer.AutoReset = false;                                // for a single shot
      boltTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent); // Hook up the Elapsed event for the timer.
      boltTimer.Stop();                                           // [Re]Start the timer
      boltTimer.Start();
      boltTimer.Enabled = true;
      while (boltTimer.Enabled)                                   // While we have not timed out or found Unit TX ID in BOLT output
      {                                                           // (we can probably remove the trigger) 
        if (boltDbg.Contains(waveforms.hwiParams[waveforms.wfidx].unitIDtrigger))
        {                                                         // When the trigger was found start parsing out the unit id
          string[] parsedDiag = boltDbg.Split('\n');
          for (int i = 0; i < parsedDiag.Length; i++)
          {
            if (parsedDiag[i].Contains(waveforms.hwiParams[waveforms.wfidx].unitIDprompt))
            {
              string unitID = (parsedDiag[i].Split(' ')[1]).Replace("\r", "");  // Extract Unit TX ID and save for various uses
              waveforms.boltParams[waveforms.wfidx].unitID = unitID;
              gpiParams.deviceId = unitID;
              cfgParams.deviceId = unitID;
              txIdTxtBox.Text = unitID;                                         // 'Keying' tab, 'Bolt Setup', 'Transmit ID'
              devInfilUniqueIdBox.Text = unitID;                                // 'Keying' tab, 'Infil Keys', 'Unique ID'
              decimal infilId;
              if (decimal.TryParse(unitID, out infilId))
              {
                uniqueIDupdown.Value = infilId;
              }
              break;
            }
          }
          boltTimer.Enabled = false;
          return true;                                            // Return with indication of SUCCESS
        }
      }
      return false;                                               // Return device timed out FAILURE
    }

    //Name: GetAndLoadTek
    //Arguments: object sender
    //           EventArgs e
    //Output: Boolean - indicates if action occurred successfully
    //Description: Upon connecting to a wiped/uninitialized Bolt the TEK key must be loaded
    //before the rest of the device can be configured
    //  First- The Brick number must be read from the Bolt output
    //  Second- The TEK key must be retrieved from the relevant XML file (should be static and only one file should be used)
    //  Third- The TEK key must be sent to the BOLT
    //NOTE - this empties boltDbg (i.e. BOLT Debug Port output we have collected)
    public bool GetAndLoadTek()
    {
      if (boltTimer != null)    {
        boltTimer.Stop();       }
      Application.DoEvents();
      boltTimer = new System.Timers.Timer(30000);                 // Setup a new 30-second timer
      boltTimer.AutoReset = false;                                // for a single shot
      boltTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent); // Hook up the Elapsed event for the timer.
      boltTimer.Stop();                                           // [Re]Start the timer
      boltTimer.Start();
      boltTimer.Enabled = true;
      bool tekKeySent = false;                                    // We must act on initially not having sent TEK
      while (boltTimer.Enabled)                                   // While we have not timed out or load BOLT's TEK(s)
      {
        Thread.Sleep(100);                                        // delay at least 1/10 second each time
        Application.DoEvents();                                   // the 'while loop' is executed, then
        if (boltDbg.Contains(waveforms.hwiParams[waveforms.wfidx].cmdPrompt)) //Ensure the Bolt is ready for commands
        {
          string[] parseBoltLoad = (boltDbg.Replace(waveforms.hwiParams[waveforms.wfidx].cmdPrompt, "|")).Split('|');//Parse out the console and make sure error was not thrown
          int submsg = (parseBoltLoad[1].Split('\n')).Length;
          if (boltDbg.Contains(waveforms.hwiParams[waveforms.wfidx].brickTrigger) && !tekKeySent) //Search for key word
          {
            string[] parsedDiag = boltDbg.Split('\n');
            for (int i = 0; i < parsedDiag.Length; i++)
            {
              if (parsedDiag[i].Contains(waveforms.hwiParams[waveforms.wfidx].brickPrompt)) //Parse buffer until Brick number is found 
              {
                string brickNum = (parsedDiag[i].Split(' ')[1]).Replace("\r", ""); //Get bricknumber
                string keyIndex = brickNum;
                //If user included Keying ID in XML then Prompt user for input
                string promptInput = "";
                Boolean correctType = false;
                if(waveforms.hwiParams[waveforms.wfidx].keyingIDPrompt.NumArguments != 0)
                {
                  SystemSounds.Beep.Play();
                  while ((promptInput.Length != waveforms.hwiParams[waveforms.wfidx].keyingIDPrompt.NumArguments || !correctType) && boltTimer.Enabled)
                  {
                    //Restart Timer
                    boltTimer.Stop();
                    boltTimer.Start();
                    boltTimer.Enabled = true;
                    promptInput = Interaction.InputBox(waveforms.hwiParams[waveforms.wfidx].keyingIDPrompt.ArgumentHelp[0] + "\nNumber of Digits: " + waveforms.hwiParams[waveforms.wfidx].keyingIDPrompt.NumArguments, "Keying ID", (int.Parse(brickNum) + int.Parse(waveforms.hwiParams[waveforms.wfidx].tekOffset)).ToString(), -1, -1);
                    if (waveforms.hwiParams[waveforms.wfidx].keyingIDPrompt.ArgumentType[0] == "int")
                    { //Check if string can be converted to integer
                      if ( ! Regex.IsMatch(promptInput, @"^\d+$"))
                      { // Return that input is wrong type
                        string msg = "WARNING: Improper Argument Format please try again... BOLT is expecting type " + waveforms.hwiParams[waveforms.wfidx].keyingIDPrompt.ArgumentType[0] + " \n";
                        string cap = "WARNING: BAD FORMAT!";
                        MessageBoxButtons btn = MessageBoxButtons.OK;
                        MessageBoxIcon icon = MessageBoxIcon.Asterisk;
                        MessageBox.Show(msg, cap, btn, icon);
                        correctType = false;
                      }
                      else
                      {
                        correctType = true;
                        keyIndex = promptInput;
                      }
                    }
                    else if (waveforms.hwiParams[waveforms.wfidx].keyingIDPrompt.ArgumentType[0] == "hex")
                    {
                      if ( ! Regex.IsMatch(promptInput, @"\A\b[0-9a-fA-F]+\b\Z"))
                      {
                        string msg = "WARNING: Improper Argument Format please try again... BOLT is expecting type " + waveforms.hwiParams[waveforms.wfidx].keyingIDPrompt.ArgumentType[0] + " \n";
                        string cap = "WARNING: BAD FORMAT!";
                        MessageBoxButtons btn = MessageBoxButtons.OK;
                        MessageBoxIcon icon = MessageBoxIcon.Asterisk;
                        MessageBox.Show(msg, cap, btn, icon);
                        correctType = false;
                      }
                      else
                      {
                        correctType = true;
                        keyIndex = promptInput;
                      }
                    }
                    if (promptInput.Length != waveforms.hwiParams[waveforms.wfidx].keyingIDPrompt.NumArguments)
                    {
                      string msg = "WARNING: Improper Number of Arguments please try again...\n BOLT is expecting " + waveforms.hwiParams[waveforms.wfidx].keyingIDPrompt.NumArguments.ToString() + " arguments\n";
                      string cap = "WARNING: BAD FORMAT!";
                      MessageBoxButtons btn = MessageBoxButtons.OK;
                      MessageBoxIcon icon = MessageBoxIcon.Asterisk;
                      MessageBox.Show(msg, cap, btn, icon);
                      correctType = false;
                    }
                  }
                }
                waveforms.boltParams[waveforms.wfidx].tekKey = GetBoltKeys(keyIndex, waveforms.hwiParams[waveforms.wfidx].tekFile); //Get TEK Key from file
                waveforms.boltParams[waveforms.wfidx].brickNumber = brickNum; //Store brick Number
                bricknumtxtbox.Text = brickNum;                               //Show Brick Number in relevant Textbox
              }
            }
            if (null == waveforms.boltParams[waveforms.wfidx].tekKey)
            { // Sometimes it is okay if we didn't find get tekKey, just quit  // TODO - is this really true (for anything other than some test wf versions in slot 0) ???
              return true;
            }
            else if (waveforms.hwiParams[waveforms.wfidx].keyOrder.Count == 0)
            { // When wf requires loading only one TEK
              if (waveforms.boltParams[waveforms.wfidx].tekKey != null)
              {
                SendBoltCommandSlowly((waveforms.hwiParams[waveforms.wfidx].loadtekCommand + " " + waveforms.boltParams[waveforms.wfidx].tekKey.key[0]), false);
              }
            }
            else
            { // When wf requires loading multiple TEKs
              for (int k = 0; k < waveforms.hwiParams[waveforms.wfidx].keyOrder.Count; k++)
              {
                if (waveforms.boltParams[waveforms.wfidx].tekKey != null)
                { // Clear BOLT Debug data captured then send TEK key and await BOLT response
                  boltDbg = "";
                  if (waveforms.hwiParams[waveforms.wfidx].loadtekCommand.Contains("%d"))
                  {
                    if (waveforms.boltParams[waveforms.wfidx].tekKey.key[waveforms.hwiParams[waveforms.wfidx].keyOrder[k] - 1] != "")
                    {
                      SendBoltCommandSlowly((waveforms.hwiParams[waveforms.wfidx].loadtekCommand.Replace("%d", (k + 1).ToString()) + " " + waveforms.boltParams[waveforms.wfidx].tekKey.key[waveforms.hwiParams[waveforms.wfidx].keyOrder[k] - 1]), false);
                    }
                    else
                    {
                      break;
                    }
                  }
                  else
                  {
                    SendBoltCommandSlowly((waveforms.hwiParams[waveforms.wfidx].loadtekCommand + " " + waveforms.boltParams[waveforms.wfidx].tekKey.key[0]), false);  // TODO - would [k] be more correct ???
                  }
                  const int _50ms = 50;                   // use 50ms short sleeps
                  int countDown = (5000/_50ms);           // for up to 5 seconds which is required
                  while (0 < countDown)                   // for BOLT to get, process, and complete a
                  {                                       // TEK key load to include emitting cmd prompt
                    Thread.Sleep(_50ms);                  // use short sleeps so progressBar scrolls.
                    Application.DoEvents();               // Each time the 'while loop' is executed
                    countDown--;                          // decrement the countdown until reach 0
                    if (boltDbg.Contains(waveforms.hwiParams[0].cmdPrompt))
                    {
                      countDown = 0;
                    }
                  }
                }
                else
                {
                  return false;
                }
              } // end for (each TEK to load)
            }
            tekKeySent = true;
          }
          //once commands are sent look for prompt to reset
          if (tekKeySent)
          {
            boltTimer.Enabled = false;
            if (submsg == 2)          {
              return true;            }
            else if (submsg != 1)     { //there was an error
              return false;           }
            else                      {
              return true;            }
          }
        }
      }
      //Time out occurred
      return false;
    } // end GetAndLoadTek

    //Name: OnTimedEvent
    //Arguments: object sender
    //           EventArgs e
    //Description: When the timer has elapsed store result into relevant variable 
    private void OnTimedEvent(object source, ElapsedEventArgs e)
    {
      boltTimer.Enabled = false;
    } // end OnTimedEvent

    //Name: CheckResetTEK
    //Arguments: object sender
    //           EventArgs e
    //Output: true-reset success keys stuck
    //        false-key did not stick error 
    //Description: Each XML File has a command options that either prompt the user for information or send information directly to the Bolt
    public bool CheckResetTEK()
    {
      //Kill the Bolt timer if it running
      if (boltTimer != null)
      { //Stop timer if it is declared
        boltTimer.Stop();
      }
      Application.DoEvents();
      //In general the Bolt takes a long time to process commands
      boltTimer = new System.Timers.Timer(30000);
      boltTimer.AutoReset = false;
      // Hook up the Elapsed event for the timer.
      boltTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
      boltTimer.Stop();
      boltTimer.Start();
      boltTimer.Enabled = true;
      boltDbg = "";
      //Send Bolt Command (It must be sent extremely slowly)
      SendBoltCommandSlowly(waveforms.hwiParams[waveforms.wfidx].resetCheckCmd, false);
      while (boltTimer.Enabled)
      { // Search the Diagnostic data for the text indicating the waveform has changed
        if (boltDbg.Contains("Password:"))
        {
          return true;
        }
        Thread.Sleep(10);
        Application.DoEvents();
      }
      //Process timed out report error
      return false;
    } // end CheckResetTEK


    //Name: SendXmlCommands
    //Arguments: object sender
    //           EventArgs e
    //Output: true-commands sent successfully
    //        false-commands were not sent successfully 
    //Description: Each XML File has a command options that either prompt the user for information or send information directly to the Bolt
    public bool SendXmlCommands()
    {
      //Clear BOLT Debug Buffer
      boltDbg = "";
      int userInput = 0;
      //for every command send to bolt and ensure it is ACK'd
      for (int i = 0; i < waveforms.commandSlot[waveforms.wfidx].commands.Count; i++)
      {
        string promptInput = "";
        string command = waveforms.commandSlot[waveforms.wfidx].commands[i].type;
        //Handle bad number of arguments
        //Prompt user for command argument
        bool correctType = true;
        if(0 != waveforms.commandSlot[waveforms.wfidx].commands[i].NumArguments) {
          SystemSounds.Beep.Play();                                              }
        while (promptInput.Length != waveforms.commandSlot[waveforms.wfidx].commands[i].NumArguments || !correctType)
        {
          promptInput = Interaction.InputBox(waveforms.commandSlot[waveforms.wfidx].commands[i].ArgumentHelp[0] + "\nNumber of characters: " + waveforms.commandSlot[waveforms.wfidx].commands[i].NumArguments, "User Argument", "Default", -1, -1);
          if (waveforms.commandSlot[waveforms.wfidx].commands[i].ArgumentType[0] == "int")
          {
            //Check if string can be converted to integer
            if ( ! Regex.IsMatch(promptInput, @"^\d+$"))
            {
              //Return that input is wrong type
              string msg = "WARNING: Improper Argument Format please try again... BOLT is expecting type " + waveforms.commandSlot[waveforms.wfidx].commands[i].ArgumentType[0] + " \n";
              string cap = "WARNING: BAD FORMAT!";
              MessageBoxButtons btn = MessageBoxButtons.OK;
              MessageBoxIcon icon = MessageBoxIcon.Asterisk;
              MessageBox.Show(msg, cap, btn, icon);
              correctType = false;
            }
            //Format is correct
            else { correctType = true; }
          }
          else if (waveforms.commandSlot[waveforms.wfidx].commands[i].ArgumentType[0] == "hex")
          {
            if ( ! Regex.IsMatch(promptInput, @"\A\b[0-9a-fA-F]+\b\Z"))
            {
              string msg = "WARNING: Improper Argument Format please try again... BOLT is expecting type " + waveforms.commandSlot[waveforms.wfidx].commands[i].ArgumentType[0] + " \n";
              string cap = "WARNING: BAD FORMAT!";
              MessageBoxButtons btn = MessageBoxButtons.OK;
              MessageBoxIcon icon = MessageBoxIcon.Asterisk;
              MessageBox.Show(msg, cap, btn, icon);
              correctType = false;
            }
            else { correctType = true; }
          }
          if (promptInput.Length != waveforms.commandSlot[waveforms.wfidx].commands[i].NumArguments)
          {
            string msg = "WARNING: Improper Number of Arguments please try again...\n BOLT is expecting " + waveforms.commandSlot[waveforms.wfidx].commands[i].NumArguments.ToString() + " arguments\n";
            string cap = "WARNING: BAD FORMAT!";
            MessageBoxButtons btn = MessageBoxButtons.OK;
            MessageBoxIcon icon = MessageBoxIcon.Asterisk;
            MessageBox.Show(msg, cap, btn, icon);
            correctType = false;
          }
        }
        //Append input from the prompt to the command
        command += promptInput;
        //Send Command
        SendBoltCommandSlowly(command, false);
        //Wait for NAK/ACK from Bolt
        // Initialize the timer with a five second interval.
        if (boltTimer != null)
        {
          boltTimer.Stop();
        }
        Application.DoEvents();
        boltTimer = new System.Timers.Timer(5000);
        boltTimer.AutoReset = false;
        // Hook up the Elapsed event for the timer.
        boltTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
        //Start the timer
        boltTimer.Stop();
        boltTimer.Start();
        boltTimer.Enabled = true;
        //Give the user 3 attempts to properly send commands to the Bolt
        while (boltTimer.Enabled && (userInput < 4))
        {
          Application.DoEvents();                 // each time the 'while loop' is executed, then
          if (boltDbg.Contains(waveforms.hwiParams[waveforms.wfidx].cmdPrompt))
          {
            string[] parseBoltLoad = (boltDbg.Replace(waveforms.hwiParams[waveforms.wfidx].cmdPrompt, " | ")).Split('|'); //Parse out the console prompt and ensure error is not thrown (this should probably be changed )
            int submsg = (parseBoltLoad[0].Split('\n')).Length;
            if (submsg == 2)
            { // When the Bolt accepted command it will return nothing but ACK and another prompt; move on to next XML command
              boltTimer.Enabled = false;
              userInput = int.MaxValue;
            }
            else if (submsg != 1 && waveforms.commandSlot[waveforms.wfidx].commands[i].NumArguments == 0)
            { // When the BOLT unhappy with a command that had no user entered parameters return FAILURE indicator for higher level code to deal with
              boltTimer.Enabled = false;
              return false;
            }
            else if (submsg != 1 && userInput < 3)
            { // When BOLT not happy with command but we still have patience with the user...
              string msg = "WARNING: BOLT NAK'd Command please try again...\n Normally this is due to improper formatting please consult code issuer if this occurs";
              string cap = "WARNING: BOLT NAK'd!";
              MessageBoxButtons btn = MessageBoxButtons.OK;
              MessageBoxIcon icon = MessageBoxIcon.Asterisk;
              MessageBox.Show(msg, cap, btn, icon);
              i--; //Decrement for loop-control indexing XML commands so user retry this command
              userInput++;
            }
            else if (userInput >= 3)
            { // When user keeps sending bad commands and all patience is expended... return FAILURE indicator for higher leverl code to deal with
              string msg = "WARNING: BOLT NAK'd the command too many times, please try again later\n ";
              string cap = "WARNING: BOLT NAK'd!";
              MessageBoxButtons btn = MessageBoxButtons.OK;
              MessageBoxIcon icon = MessageBoxIcon.Asterisk;
              MessageBox.Show(msg, cap, btn, icon);
              boltTimer.Enabled = false;
              return false;
            }
          }
        } // end while attempting a command
        boltDbg = "";  // dispose reams of now unneeded data before sending next command
      } // end for more commands
      boltTimer.Stop();
      return true;
    } // end SendXmlCommands

    //Name: ConnectToLightning  TODO - 202205/01 call QueueLtMessage() instead of SendLightningCommands(); also call this from MakeHwiPrgmConnection() and get rid of duplicating code
    //Arguments: object sender
    //           EventArgs e
    //Output: true-Lightning connected
    //        false- Lightning not connected
    //Description: Due to the way the BOLT firmware is written we need to power-cycle the Bolt everytime a TEK Key is loaded
    //That means we need to change COMx from BOLT to Lightning and re-detect LBHH after the power-cycle 
    //This function is typically called once the user has notified the programmer that the bolt has been power cycled
    public bool ConnectToLightning(string portName)
    {
      cfgParams.connected = false;
      //Assume one Bolt plugged in at a time 
      if (comPort.IsOpen || (portName != comPort.PortName))
      {                                         // Changing ports, so disconnect from connected port
        StopDiagThread();                       //No need to run diag while connecting
        StopMessagingThreads();                 // Shutdown messaging threads before changing port
        Thread.Sleep(200);                      // Give device more time to see 'Disconnect'
        Application.DoEvents();                 // and process any message request/user inputs
        comPort.Close();                        // before closing port (which unfortunately
      }                                         // sets RTS back to HwIntfcEnabled)
                                                //If the comPort has been closed open comPort
      if ( ! comPort.IsOpen)
      {     // if just now opening the COM port,
        OpenLightningComPortProperty(portName);
      }     // fully configure it

      //Use the method below for finding device restart
      loRespStatus = (cfgParams.connected)           // If a Lightning is already connected
                   ? LoRespStatus.Null          // then no LO msg expected (interface is open)
                   : LoRespStatus.GAKreq;          // otherwise we are looking for an LO ACK msg
      runMsgThreads = true;                     // Set flag so message threads will run
      if ( ! ltBackgroundWorker.IsBusy)
      { // Send LT message thread shouldn't be running
        ltBackgroundWorker.RunWorkerAsync();
      } // so start it now (if not already running)
      Thread.Sleep(0);                          // Give LT background worker thread time to startup
      Application.DoEvents();                   // then process events that may have occurred
      if ( ! loBackgroundWorker.IsBusy)
      { // Receive LO message thread shouldn't be running
        loBackgroundWorker.RunWorkerAsync();
      } // so start it now (if not already running)
      Thread.Sleep(20);                         // Give background worker threads time to startup
      Application.DoEvents();                   // then process events that may have occurred
      if (boltTimer != null)
      {
        boltTimer.Stop();
      }
      Application.DoEvents();
      boltTimer = new System.Timers.Timer(10000);
      boltTimer.AutoReset = false;
      // Hook up the Elapsed event for the timer.
      boltTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
      //Start the timer
      boltTimer.Stop();
      boltTimer.Start();
      boltTimer.Enabled = true;
      //Poll Lightning for ID
      //Poll Lighting to Ensure Connection
      while (boltTimer.Enabled) //While we have not timed out or recieved WFMID
      {
        //Send command to check if lightning is connected
        SendLightningCommands("¡S.GAK¶");
        //Look for response
        //if response found set to true
        Thread.Sleep(100);
        Application.DoEvents();
        //If ACK is received return that lightning is connected
        if (LightningDiag.Contains("+S.GAK"))
        {
          return true;
        }
      }
      //Process Timed out Alert user something is wrong 
      return false;
    } // end ConnectToLightning

    //Name: SendBoltWfchangeCmd
    //Arguments: object sender
    //           EventArgs e
    //Output: bool - was waveform change successful
    //Description: Send the waveform change command to the BOLT
    public bool SendBoltWfchangeCmd(int waveform)
    {
      if (boltTimer != null)
      { //Stop timer if it is declared
        boltTimer.Stop();
      }
      Application.DoEvents();
      //In general the Bolt takes a long time to process commands
      boltTimer = new System.Timers.Timer(30000);
      boltTimer.AutoReset = false;
      // Hook up the Elapsed event for the timer.
      boltTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
      boltTimer.Stop();
      boltTimer.Start();
      boltTimer.Enabled = true;
      boltDbg = "";
      //Send Bolt Command (It must be sent extremely slowly)
      SendBoltCommandSlowly(("wfchange " + waveform.ToString()), false);
      while (boltTimer.Enabled)
      { // While op hasn't time out look for text indicating 'wfchange' was sent
        if (boltDbg.Contains("ChangeToWaveform(" + waveform.ToString() + ")"))
        {
          return true;
        }
        Thread.Sleep(10);
        Application.DoEvents();
      }
      //Process timed out report error
      return false;
    } // end SendBoltWfchangeCmd

  }

}

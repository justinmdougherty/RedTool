using System;
using System.Xml;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; //Needed for operation with excel
using System.Collections.Generic;


namespace LBHH_Red {
  public partial class LBHH_HWI {
    //## REGION ################################################################
    #region Defines
    internal const string DefEncryptionFileName = "\\KeyChain.xlsx";
    internal const string EncryptionPassword = "I 8 krypton!te";
    //extern InfilKeyChain infilKeys;
    #endregion

    //## REGION ################################################################
    #region FieldsAndProperties
    internal string EncryptionFilePathName = "";
    #endregion FieldsAndProperties

    //## REGION ################################################################
    #region Methods

    /*Name:  GetUniqueIdKey
     *Inputs:     int uniqueID- unique Key ID for which encryption key is required
     *Outputs:    ref InfilKeyChain infilKeys- If successful loads key and id into keychain
     *            returns true if key file was read
     *Notes: Retrieves user requested unique key and encryption information *Read Only*
     *       For efficiency purposes this is only to be called if user changes device ID and group ID's have been pulled
     */
    private bool GetUniqueIdKey(int KeyID, ref InfilKeyChain keys)
    {
      bool fileWasRead = false;
      EncryptionFilePathName = appDir + DefEncryptionFileName;
      if (File.Exists(EncryptionFilePathName))
      { // Should try-catch this code-block but...
        //--------------------------------------------------Open Excel Document-------------------------------------------------------------------
        Excel.Application appExcel = new Excel.Application();
        appExcel.Visible = false;
        Excel.Workbook WB = appExcel.Workbooks.Open(EncryptionFilePathName, ReadOnly: true, Password: EncryptionPassword);
        Excel.Sheets excelSheets = WB.Worksheets;
        string currentSheet = "UniqueKeys";
        Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
        //-----------------------------------------------------------------------------------------------------------------------------------------
        //Check worksheet formatting... read data only if column headers are correct otherwise inform user of issue and halt device programming
        if ("ID" == (string)(excelWorksheet.Cells[1, 1] as Excel.Range).Value && "Encryption" == (string)(excelWorksheet.Cells[1, 2] as Excel.Range).Value) 
        { //Retrieve only the desired unique key
          //RowIndex = uniqueID + 2; // index needs +2 since uniqueID is 0-based (though operationally 0 is invalid) and first row is column headers
          //ColumnIndex = 2;      // 2nd column contains encryption keys (1st column is unique ID)
          keys.uniqueKey = (string)((excelWorksheet.Cells[KeyID + 2, 2] as Excel.Range).Value).ToString();
          keys.deviceID = KeyID;
          fileWasRead = true;
        }
        else
        {	//Inform User that File linked is incorrect format
          string msg = "FAILED: Encryption Key File is not formatted properly";
          string cap = "Encryption File";
          MessageBoxButtons btn = MessageBoxButtons.OK;
          MessageBoxIcon icon = MessageBoxIcon.Asterisk;
          MessageBox.Show(msg, cap, btn, icon);
        }
        //---------------------------------------------------------Close Excel Document--------------------------------------------------------------
        appExcel.Workbooks.Close();
        appExcel.Quit();
        GC.Collect();
        GC.WaitForPendingFinalizers();
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(WB);
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(appExcel);
        //-----------------------------------------------------------------------------------------------------------------------------------------
      }
      else
      { //Inform User that File was not found
        string msg = "FAILED: Encryption Key File cannot be found at expected location: " + EncryptionFilePathName;
        string cap = "Key Retrieval";
        MessageBoxButtons btn = MessageBoxButtons.OK;
        MessageBoxIcon icon = MessageBoxIcon.Asterisk;
        MessageBox.Show(msg, cap, btn, icon);
      }
      return fileWasRead;
    } // end GetUniqueIdKey

    /*Name:  GetInfilKeys
     *Inputs:     int uniqueID- device unique ID to look for unique encryption key
     *Outputs:    ref InfilKeyChain infilKeys - If successful loads unique and group key info into keychain
     *            ref obfKeys - If successful loads C.OLT obfuscation key commands into obfKeys
     *            returns true if key file was read
     *Notes: Retrieves user requested unique AND group keys and encryption information AND obfuscation keys
     */
    private bool GetInfilKeys(int uniqueID, ref InfilKeyChain infilKeys, ref List<string> obfKeys)
    {
      bool fileWasRead = false;
      EncryptionFilePathName = appDir + DefEncryptionFileName;
      if (File.Exists(EncryptionFilePathName))
      { // Should try-catch this code-block but...
        //--------------------------------------------------Open Excel Document-------------------------------------------------------------------
        Excel.Application appExcel = new Excel.Application();
        appExcel.Visible = false;
        Excel.Workbook WB = appExcel.Workbooks.Open(EncryptionFilePathName, ReadOnly: true, Password: EncryptionPassword);
        Excel.Sheets excelSheets = WB.Worksheets;

        // Retrieve desired single Unique ID key first and use validity checks of that worksheet to determine if we proceed to extract any data
        string currentSheet = "UniqueKeys";
        Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
        if ("ID" == (string)(excelWorksheet.Cells[1, 1] as Excel.Range).Value && "Encryption" == (string)(excelWorksheet.Cells[1, 2] as Excel.Range).Value)
        { // RowIndex = uniqueID + 2; // index needs +2 since uniqueID is 0-based (though operationally 0 is invalid) and first row is column headers
          // ColumnIndex = 2;         // 2nd column contains encryption keys (1st column is unique ID)
          infilKeys.uniqueKey = (string)((excelWorksheet.Cells[uniqueID + 2, 2] as Excel.Range).Value).ToString();
          infilKeys.deviceID = uniqueID;

          // Retrieve ALL activated group keys (it is possible, but unrealistic, that no group keys have been activated)
          {
            currentSheet = "ActiveGroupKeys";
            excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
            if (((excelWorksheet.Cells[2, 2] as Excel.Range).Value) != null)
            {	//When at least 1 group key exisits, go through all group keys in the ActiveGroupKeys worksheet
              int rowIdx = 2;
              infilKeys.GroupKeys.Clear();
              while ( ! string.IsNullOrEmpty(Convert.ToString((excelWorksheet.Cells[rowIdx, 2] as Excel.Range).Value)))
              {	// If group key is populated then store it until Red Tool is closed
                InfilGroupKey tempGroupkey = new InfilGroupKey();
                tempGroupkey.index = Convert.ToString((excelWorksheet.Cells[rowIdx, 1] as Excel.Range).Value);
                tempGroupkey.GroupKey = Convert.ToString((excelWorksheet.Cells[rowIdx, 2] as Excel.Range).Value);
                tempGroupkey.GroupName = Convert.ToString((excelWorksheet.Cells[rowIdx, 3] as Excel.Range).Value);
                tempGroupkey.SponsorName = Convert.ToString((excelWorksheet.Cells[rowIdx, 4] as Excel.Range).Value);
                tempGroupkey.timestamp = Convert.ToString((excelWorksheet.Cells[rowIdx, 5] as Excel.Range).Value);
                infilKeys.GroupKeys.Add(tempGroupkey);
                rowIdx++;
              } // end while a group key to read
              infilKeys.retrievedKeys = true;
            }
          } // end if group keys exist

          // Retrieve ALL Lightning C.OLT Obfuscation Look-up Table commands
          {
            currentSheet = "C_OLT";
            excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
            int rowIdx;
            obfKeys.Clear();
            for (rowIdx = 1; numObfKeys >= rowIdx; rowIdx++)
            { // get the 10 chars required for 2 parameters of ¡C.OLT 0 12345678¶ commands -- 62 times
              string keyStr = Convert.ToString((excelWorksheet.Cells[rowIdx,1] as Excel.Range).Value);
              if (10 == keyStr.Length)    {           // Check for proper length and add
                obfKeys.Add(keyStr);      }           // only keys/commands of proper length
              else                        {           // Cell values of improper length
                break;                    }           // cause for-loop to break prematurely
            }
            fileWasRead = (numObfKeys+1 == rowIdx);   // Only when got all obfKeys do we consider file read
          }
        }
        else
        {	//Inform user that File linked is incorrect format
          string msg = "FAILED: Encryption Key File is not formatted properly\r\n"
                     + "You cannot 'Program Infil Keys' until a valid Key File is read";
          string cap = "Encryption File";
          MessageBoxButtons btn = MessageBoxButtons.OK;
          MessageBoxIcon icon = MessageBoxIcon.Asterisk;
          MessageBox.Show(msg, cap, btn, icon);
        }
        //---------------------------------------------------------Close Excel Document--------------------------------------------------------------
        appExcel.Workbooks.Close();
        appExcel.Quit();
        GC.Collect();
        GC.WaitForPendingFinalizers();
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(WB);
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(appExcel);
      }
      else
      { //Inform User that File was not found
        string msg = "FAILED: Encryption Key File cannot be found at expected location: " + EncryptionFilePathName;
        string cap = "Key Retrieval";
        MessageBoxButtons btn = MessageBoxButtons.OK;
        MessageBoxIcon icon = MessageBoxIcon.Asterisk;
        MessageBox.Show(msg, cap, btn, icon);
      }
      return fileWasRead;
    }

    // AddEditActiveKeys
    // Actions: Opens Excel Document 
    //         1. Pulls Current Active Keys and loads them into Key Chain
    //         2. See's if index is being used
    //         3. a. If index in use take current name and timestamp concatenate and add to current previous associations column
    //         3. b. If the index is not in use add it to the list and write to excel file
    //         3. c. If deleting key (delete checkbox clicked) - WRITE !DELETED! and do same actions as step a.
    //         4. Prompt user that device will be disconnected and all active keys will need to be reloaded
    // Inputs:  string grpName - Desired Group Name associated with index
    //          int index - Index to be edited
    //          bool deletechkbox - Does user want to delete values
    // Outputs: None
    private void AddRemoveActiveKeys(string grpName, string sponsorName, int index) {
      string currentSheet = "ActiveGroupKeys";
      Excel.Worksheet excelWorksheet;
      EncryptionFilePathName = appDir + DefEncryptionFileName;
      if (File.Exists(EncryptionFilePathName)) {
        //---------------------------------------------------------Open Excel Document-------------------------------------------------------------
        // Open Document and set global variable for reading
        Excel.Application appExcel = new Excel.Application();
        appExcel.Visible = false;
        Excel.Workbook WB = appExcel.Workbooks.Open(EncryptionFilePathName, ReadOnly: false, Password: EncryptionPassword);
        Excel.Sheets excelSheets = WB.Worksheets;
        currentSheet = "ActiveGroupKeys";
        excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
        InfilKeyChain tempkeychain = new InfilKeyChain();
        //-----------------------------------------------------------------------------------------------------------------------------------------
        try {
          excelWorksheet.Unprotect(EncryptionPassword); // Unlock for editing
          excelWorksheet.Cells[2, 2].NumberFormat = "@"; // Ensure format is string
          //if ( ! string.IsNullOrEmpty(Convert.ToString((excelWorksheet.Cells[2, 2] as Excel.Range).Value))) {
            // Pull information from Excel
            int j = 0;
            while ( ! string.IsNullOrEmpty(Convert.ToString((excelWorksheet.Cells[j + 2, 2] as Excel.Range).Value))) {
              excelWorksheet.Cells[j + 2, 2].NumberFormat = "@"; // Ensure format is string
              excelWorksheet.Cells[j + 2, 3].NumberFormat = "@"; // Ensure format is string
              excelWorksheet.Cells[j + 2, 4].NumberFormat = "@"; // Ensure format is string
              excelWorksheet.Cells[j + 2, 5].NumberFormat = "@"; // Ensure format is string 
              InfilGroupKey tempGroupkey = new InfilGroupKey();
              // If the name is populated then store it
              tempGroupkey.index = Convert.ToString((excelWorksheet.Cells[j + 2, 1] as Excel.Range).Value);
              tempGroupkey.GroupKey = Convert.ToString((excelWorksheet.Cells[j + 2, 2] as Excel.Range).Value);
              tempGroupkey.GroupName = Convert.ToString((excelWorksheet.Cells[j + 2, 3] as Excel.Range).Value);
              tempGroupkey.SponsorName = Convert.ToString((excelWorksheet.Cells[j + 2, 4] as Excel.Range).Value);
              tempGroupkey.timestamp = Convert.ToString((excelWorksheet.Cells[j + 2, 5] as Excel.Range).Value);
              tempkeychain.GroupKeys.Add(tempGroupkey);
              j++;
            }				

            // Make sure we are not deleting
            if ( ! string.IsNullOrEmpty(grpName)) {
              // Check if desired name is taken if it is tell user to either delete the group name association or try a different name
              for (int i = 0; i < tempkeychain.GroupKeys.Count; i++) {
                // If this name is in use else where complain unless it is deleting the value
                if (tempkeychain.GroupKeys[i].GroupName == grpName) {
                  // Close excel worksheet
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
                  // prompt user that name is taken
                  string msgwindow = "";

                  // check if existing name is the one in use
                  if (tempkeychain.GroupKeys[i].index == index.ToString()) {
                    msgwindow = "ACTION FAILED: Group Name: " + tempkeychain.GroupKeys[i].GroupName + " is identical to current Group Name at Index: " + tempkeychain.GroupKeys[i].index + ". Please use a different name!";
                  } else {
                    msgwindow = "ACTION FAILED: Group Name: " + tempkeychain.GroupKeys[i].GroupName + " is already in use. Please delete the group name at Index: " + tempkeychain.GroupKeys[i].index + ", or use a different name!";
                  }
                    
                  string capwindow = "Group Key Action";
                  MessageBoxButtons btnwindow = MessageBoxButtons.OK;
                  MessageBoxIcon iconwindow = MessageBoxIcon.Asterisk;
                  MessageBox.Show(msgwindow, capwindow, btnwindow, iconwindow);
                    
                  return;
                }
              }
            } else {
              if (tempkeychain.GroupKeys.Count == 0) {
                // No valid Group Names present in excel file
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
                string msgwindow = "ACTION FAILED: No Group Names exist in Keychain.";
                string capwindow = "Group Key Action";
                MessageBoxButtons btnwindow = MessageBoxButtons.OK;
                MessageBoxIcon iconwindow = MessageBoxIcon.Asterisk;
                MessageBox.Show(msgwindow, capwindow, btnwindow, iconwindow);

                return;
              }
            }

            // See if index exists in Active group Keys
            for (int i = 0; i < tempkeychain.GroupKeys.Count; i++) {
              if (tempkeychain.GroupKeys[i].index == index.ToString()) {
                // Check if deleting or not
                if (string.IsNullOrEmpty(grpName)) {
                  // Check if group name is empty or not
                  if ( ! string.IsNullOrEmpty(Convert.ToString((excelWorksheet.Cells[i + 2, 3] as Excel.Range).Value))) {
                    // Delete the group name and update the previous association
                    string associationValues = "";
                    if ( ! string.IsNullOrEmpty(Convert.ToString((excelWorksheet.Cells[i + 2, 6] as Excel.Range).Value))) {
                      associationValues = Convert.ToString((excelWorksheet.Cells[i + 2, 6] as Excel.Range).Value);
                    }
                    excelWorksheet.Cells[i + 2, 3].NumberFormat = "@";
                    excelWorksheet.Cells[i + 2, 4].NumberFormat = "@";
                    excelWorksheet.Cells[i + 2, 5].NumberFormat = "@";
                    excelWorksheet.Cells[i + 2, 6].NumberFormat = "@";
                    excelWorksheet.Cells[i + 2, 3].Value = "";
                    excelWorksheet.Cells[i + 2, 4].Value = "";
                    excelWorksheet.Cells[i + 2, 5].Value = DateTime.Now.ToString("yyyy.MMMM.dd(HH:mm:ss)");
                    if (associationValues != "") {
                      excelWorksheet.Cells[i + 2, 6].Value = associationValues + "/" + tempkeychain.GroupKeys[i].timestamp + "_" + tempkeychain.GroupKeys[i].GroupName + "[" + tempkeychain.GroupKeys[i].SponsorName + "]";
                    } else {
                      excelWorksheet.Cells[i + 2, 6].Value = tempkeychain.GroupKeys[i].timestamp + "_" + tempkeychain.GroupKeys[i].GroupName + "[" + tempkeychain.GroupKeys[i].SponsorName + "]";
                    }

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
                    string msgwindow = "ACTION COMPLETED: Key Name has been deleted.";
                    string capwindow = "Group Key Action";
                    MessageBoxButtons btnwindow = MessageBoxButtons.OK;
                    MessageBoxIcon iconwindow = MessageBoxIcon.Asterisk;
                    MessageBox.Show(msgwindow, capwindow, btnwindow, iconwindow);

                    return;

                  } else {
                    // Do nothing
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
                    string msgwindow = "ACTION FAILED: Group Name is already deleted.";
                    string capwindow = "Group Key Action";
                    MessageBoxButtons btnwindow = MessageBoxButtons.OK;
                    MessageBoxIcon iconwindow = MessageBoxIcon.Asterisk;
                    MessageBox.Show(msgwindow, capwindow, btnwindow, iconwindow);

                    return;
                  }
                } else {
                  if (string.IsNullOrEmpty(Convert.ToString((excelWorksheet.Cells[i + 2, 3] as Excel.Range).Value))) {
                    string associationValues = Convert.ToString((excelWorksheet.Cells[i + 2, 6] as Excel.Range).Value);
                    excelWorksheet.Cells[i + 2, 3].NumberFormat = "@";
                    excelWorksheet.Cells[i + 2, 4].NumberFormat = "@";
                    excelWorksheet.Cells[i + 2, 5].NumberFormat = "@";
                    excelWorksheet.Cells[i + 2, 6].NumberFormat = "@";
                    excelWorksheet.Cells[i + 2, 3].Value = grpName;
                    excelWorksheet.Cells[i + 2, 4].Value = sponsorName;
                    excelWorksheet.Cells[i + 2, 5].Value = DateTime.Now.ToString("yyyy.MMMM.dd(HH:mm:ss)");
                    excelWorksheet.Cells[i + 2, 6].Value = associationValues + "/" + tempkeychain.GroupKeys[i].timestamp + "_DELETED";

                    // See if group name needs to be added to list
                    if ( ! groupNames.Contains(grpName)) {
                      groupNames.Add(grpName);
                      groupNameKeyPairString.Items.Clear();
                      foreach (string group in groupNames) {
                        groupNameKeyPairString.Items.Add(group);
                      }
                    }

                    // Check if sponsor name needs to be added to list
                    if ( ! sponsorNames.Contains(sponsorName)) {
                      excelWorksheet.UsedRange.Columns.AutoFit();
                      excelWorksheet.Protect(EncryptionPassword); // Lock Editing
                      currentSheet = "Sponsor Names";
                      excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
                      excelWorksheet.Unprotect(EncryptionPassword);
                      excelWorksheet.Cells[sponsorNames.Count + 1, 1].NumberFormat = "@";
                      excelWorksheet.Cells[sponsorNames.Count + 1, 1].Value = sponsorName;
                      sponsorNames.Add(sponsorName);
                      sponsorNameKeyPairString.Items.Clear();
                      foreach (string sponsor in sponsorNames) {
                        sponsorNameKeyPairString.Items.Add(sponsor);
                      }
                    }

                    // Close Excel
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
                    string diagStrwindow2 = (DateTime.UtcNow.ToString("u") + " Key changes have been made to excel sheet\n");
                    string msgwindow2 = "ACTION COMPLETED: Key Name has been added.";
                    string capwindow2 = "Group Key Action";
                    MessageBoxButtons btnwindow2 = MessageBoxButtons.OK;
                    MessageBoxIcon iconwindow2 = MessageBoxIcon.Asterisk;
                    MessageBox.Show(msgwindow2, capwindow2, btnwindow2, iconwindow2);

                    return;
                  } else {
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
                    string msgwindow = "ACTION FAILED: Index: " + tempkeychain.GroupKeys[i].index + " is already in use by Group Name: " + tempkeychain.GroupKeys[i].GroupName + ". Please delete the Group Name first, or specify a different Index!";
                    string capwindow = "Group Key Action";
                    MessageBoxButtons btnwindow = MessageBoxButtons.OK;
                    MessageBoxIcon iconwindow = MessageBoxIcon.Asterisk;
                    MessageBox.Show(msgwindow, capwindow, btnwindow, iconwindow);
                    return;
                  }
                }									
              }
            }							
            // If index key pair not created (happens when it is not found in excel file)
            if ( ! string.IsNullOrEmpty(grpName)) {
              // Switch to other sheet
              excelWorksheet.UsedRange.Columns.AutoFit();
              excelWorksheet.Protect(EncryptionPassword); // Lock Editing
              string GroupKeySheetname = "GroupKeys";
              excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(GroupKeySheetname);
              excelWorksheet.Unprotect(EncryptionPassword);
              string Groupindex = Convert.ToString((excelWorksheet.Cells[index + 2, 1] as Excel.Range).Value);
              string GroupKey = Convert.ToString((excelWorksheet.Cells[index + 2, 2] as Excel.Range).Value);

              // Set Active Group Keys
              excelWorksheet.UsedRange.Columns.AutoFit();
              excelWorksheet.Protect(EncryptionPassword); // Lock Editing
              string ActiveGroupKeySheetname = "ActiveGroupKeys";
              excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(ActiveGroupKeySheetname);
              excelWorksheet.Unprotect(EncryptionPassword); // Unlock for editing
              excelWorksheet.Cells[tempkeychain.GroupKeys.Count + 2, 1].NumberFormat = "@";
              excelWorksheet.Cells[tempkeychain.GroupKeys.Count + 2, 2].NumberFormat = "@";
              excelWorksheet.Cells[tempkeychain.GroupKeys.Count + 2, 3].NumberFormat = "@";
              excelWorksheet.Cells[tempkeychain.GroupKeys.Count + 2, 4].NumberFormat = "@";
              excelWorksheet.Cells[tempkeychain.GroupKeys.Count + 2, 5].NumberFormat = "@";
              excelWorksheet.Cells[tempkeychain.GroupKeys.Count + 2, 6].NumberFormat = "@";
              excelWorksheet.Cells[tempkeychain.GroupKeys.Count + 2, 1].Value = Groupindex;
              excelWorksheet.Cells[tempkeychain.GroupKeys.Count + 2, 2].Value = GroupKey;
              excelWorksheet.Cells[tempkeychain.GroupKeys.Count + 2, 3].Value = grpName;
              excelWorksheet.Cells[tempkeychain.GroupKeys.Count + 2, 4].Value = sponsorName;
              excelWorksheet.Cells[tempkeychain.GroupKeys.Count + 2, 5].Value = DateTime.Now.ToString("yyyy.MMMM.dd(HH:mm:ss)");

              // See if group name needs to be added to list
              if ( ! groupNames.Contains(grpName)) {
                groupNames.Add(grpName);
                groupNameKeyPairString.Items.Clear();
                foreach (string group in groupNames) {
                  groupNameKeyPairString.Items.Add(group);
                }
              }

              // Check if sponsor name needs to be added to list
              if ( ! sponsorNames.Contains(sponsorName)) {
                excelWorksheet.UsedRange.Columns.AutoFit();
                excelWorksheet.Protect(EncryptionPassword); // Lock Editing
                currentSheet = "Sponsor Names";
                excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
                excelWorksheet.Unprotect(EncryptionPassword);
                excelWorksheet.Cells[sponsorNames.Count + 1, 1].NumberFormat = "@";
                excelWorksheet.Cells[sponsorNames.Count + 1, 1].Value = sponsorName;
                sponsorNames.Add(sponsorName);
                sponsorNameKeyPairString.Items.Clear();
                foreach (string sponsor in sponsorNames) {
                  sponsorNameKeyPairString.Items.Add(sponsor);
                }
              }

              string msg = "ACTION COMPLETED: Key Name has been added";
              string cap = "Group Key Action";
              MessageBoxButtons btn = MessageBoxButtons.OK;
              MessageBoxIcon icon = MessageBoxIcon.Asterisk;
              MessageBox.Show(msg, cap, btn, icon);
            } else {
              string msg = "ACTION FAILED: No key at this index exists, or no keys in keychain!";
              string cap = "Group Key Action";
              MessageBoxButtons btn = MessageBoxButtons.OK;
              MessageBoxIcon icon = MessageBoxIcon.Asterisk;
              MessageBox.Show(msg, cap, btn, icon);
            }
          /*} else {
            string msg = "ACTION FAILED: No keys exist in the keychain!";
            string cap = "Group Key Action";
            MessageBoxButtons btn = MessageBoxButtons.OK;
            MessageBoxIcon icon = MessageBoxIcon.Asterisk;
            MessageBox.Show(msg, cap, btn, icon);
          }*/
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
        } catch (Exception) {
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
      }
    }

    private void AddSponsorName(string groupName, string sponsorName, int index) {
      string currentSheet = "ActiveGroupKeys";
      Excel.Worksheet excelWorksheet;
      EncryptionFilePathName = appDir + DefEncryptionFileName;
      if (File.Exists(EncryptionFilePathName)) {
        //---------------------------------------------------------Open Excel Document-------------------------------------------------------------
        // Open Document and set global variable for reading
        Excel.Application appExcel = new Excel.Application();
        appExcel.Visible = false;
        Excel.Workbook WB = appExcel.Workbooks.Open(EncryptionFilePathName, ReadOnly: false, Password: EncryptionPassword);
        Excel.Sheets excelSheets = WB.Worksheets;
        currentSheet = "ActiveGroupKeys";
        excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
        InfilKeyChain tempkeychain = new InfilKeyChain();
        try {
          excelWorksheet.Unprotect(EncryptionPassword); // Unlock for editing
          excelWorksheet.Cells[2, 2].NumberFormat = "@"; // Ensure format is string
          if ( ! string.IsNullOrEmpty(Convert.ToString((excelWorksheet.Cells[2, 2] as Excel.Range).Value))) {
            // Pull information from Excel
            int j = 0;
            while ( ! string.IsNullOrEmpty(Convert.ToString((excelWorksheet.Cells[j + 2, 2] as Excel.Range).Value))) {
              excelWorksheet.Cells[j + 2, 2].NumberFormat = "@"; // Ensure format is string
              excelWorksheet.Cells[j + 2, 3].NumberFormat = "@"; // Ensure format is string
              excelWorksheet.Cells[j + 2, 4].NumberFormat = "@"; // Ensure format is string
              excelWorksheet.Cells[j + 2, 5].NumberFormat = "@"; // Ensure format is string 
              InfilGroupKey tempGroupkey = new InfilGroupKey();
              // If the name is populated then store it
              tempGroupkey.index = Convert.ToString((excelWorksheet.Cells[j + 2, 1] as Excel.Range).Value);
              tempGroupkey.GroupKey = Convert.ToString((excelWorksheet.Cells[j + 2, 2] as Excel.Range).Value);
              tempGroupkey.GroupName = Convert.ToString((excelWorksheet.Cells[j + 2, 3] as Excel.Range).Value);
              tempGroupkey.SponsorName = Convert.ToString((excelWorksheet.Cells[j + 2, 4] as Excel.Range).Value);
              tempGroupkey.timestamp = Convert.ToString((excelWorksheet.Cells[j + 2, 5] as Excel.Range).Value);
              tempkeychain.GroupKeys.Add(tempGroupkey);
              j++;
            }						

            // Check if ID Exists
            for (int i = 0; i < tempkeychain.GroupKeys.Count; i++) {
              // Check ID
              if (tempkeychain.GroupKeys[i].index == index.ToString()) {
                // Check if group name is empty or not
                string tempSponsorName = Convert.ToString((excelWorksheet.Cells[i + 2, 4] as Excel.Range).Value);
                if ( ! string.IsNullOrEmpty(tempSponsorName)) {
                  // Check if Group contains Sponsor Name already
                  if ( ! tempSponsorName.Contains(sponsorName)) {
                    // Update Sponsor Name
                    excelWorksheet.Cells[i + 2, 4].NumberFormat = "@";
                    excelWorksheet.Cells[i + 2, 4].Value = tempSponsorName + "/" + sponsorName;

                    // Update associationValues
                    string associationValues = "";
                    if ( ! string.IsNullOrEmpty(Convert.ToString((excelWorksheet.Cells[i + 2, 6] as Excel.Range).Value))) {
                      associationValues = Convert.ToString((excelWorksheet.Cells[i + 2, 6] as Excel.Range).Value);
                    }
                    excelWorksheet.Cells[i + 2, 5].NumberFormat = "@";
                    excelWorksheet.Cells[i + 2, 6].NumberFormat = "@";
                    excelWorksheet.Cells[i + 2, 5].Value = DateTime.Now.ToString("yyyy.MMMM.dd(HH:mm:ss)");
                    if (associationValues != "") {
                      excelWorksheet.Cells[i + 2, 6].Value = associationValues + "/" + tempkeychain.GroupKeys[i].timestamp + "_" + tempkeychain.GroupKeys[i].GroupName + "[" + tempkeychain.GroupKeys[i].SponsorName + "]";
                    } else {
                      excelWorksheet.Cells[i + 2, 6].Value = tempkeychain.GroupKeys[i].timestamp + "_" + tempkeychain.GroupKeys[i].GroupName + "[" + tempkeychain.GroupKeys[i].SponsorName + "]";
                    }

                    // Update Sponsor Names Tab if Sponsor Name is new & add to sponsorNames & add to sponsorNameKeyPairString
                    if ( ! sponsorNames.Contains(sponsorName)) {
                      excelWorksheet.UsedRange.Columns.AutoFit();
                      excelWorksheet.Protect(EncryptionPassword); // Lock Editing
                      currentSheet = "Sponsor Names";
                      excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
                      excelWorksheet.Unprotect(EncryptionPassword);
                      excelWorksheet.Cells[sponsorNames.Count + 1, 1].NumberFormat = "@";
                      excelWorksheet.Cells[sponsorNames.Count + 1, 1].Value = sponsorName;
                      sponsorNames.Add(sponsorName);
                      sponsorNameKeyPairString.Items.Clear();
                      foreach (string sponsor in sponsorNames) {
                        sponsorNameKeyPairString.Items.Add(sponsor);
                      }
                    }

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
                    string msgwindow = "ACTION COMPLETED: Sponsor Name has been added.";
                    string capwindow = "Group Key Action";
                    MessageBoxButtons btnwindow = MessageBoxButtons.OK;
                    MessageBoxIcon iconwindow = MessageBoxIcon.Asterisk;
                    MessageBox.Show(msgwindow, capwindow, btnwindow, iconwindow);

                    return;
                  } else {
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
                    string msgwindow = "ACTION FAILED: Group already contains Sponsor Name.";
                    string capwindow = "Group Key Action";
                    MessageBoxButtons btnwindow = MessageBoxButtons.OK;
                    MessageBoxIcon iconwindow = MessageBoxIcon.Asterisk;
                    MessageBox.Show(msgwindow, capwindow, btnwindow, iconwindow);

                    return;
                  }						
                } else {
                  // Do nothing
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
                  string msgwindow = "ACTION FAILED: Group Name is null.";
                  string capwindow = "Group Key Action";
                  MessageBoxButtons btnwindow = MessageBoxButtons.OK;
                  MessageBoxIcon iconwindow = MessageBoxIcon.Asterisk;
                  MessageBox.Show(msgwindow, capwindow, btnwindow, iconwindow);

                  return;
                }
              }
            }

            string msg = "ACTION FAILED: ID does not exist in keychain!";
            string cap = "Group Key Action";
            MessageBoxButtons btn = MessageBoxButtons.OK;
            MessageBoxIcon icon = MessageBoxIcon.Asterisk;
            MessageBox.Show(msg, cap, btn, icon);

          } else {
            string msg = "ACTION FAILED: No keys exist in the keychain!";
            string cap = "Group Key Action";
            MessageBoxButtons btn = MessageBoxButtons.OK;
            MessageBoxIcon icon = MessageBoxIcon.Asterisk;
            MessageBox.Show(msg, cap, btn, icon);
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
        } catch (Exception) {
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
      }
    }

    private void RemoveSponsorName(string groupName, string sponsorName, int index) {
      string currentSheet = "ActiveGroupKeys";
      Excel.Worksheet excelWorksheet;
      EncryptionFilePathName = appDir + DefEncryptionFileName;
      if (File.Exists(EncryptionFilePathName)) {
        //---------------------------------------------------------Open Excel Document-------------------------------------------------------------
        // Open Document and set global variable for reading
        Excel.Application appExcel = new Excel.Application();
        appExcel.Visible = false;
        Excel.Workbook WB = appExcel.Workbooks.Open(EncryptionFilePathName, ReadOnly: false, Password: EncryptionPassword);
        Excel.Sheets excelSheets = WB.Worksheets;
        currentSheet = "ActiveGroupKeys";
        excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
        InfilKeyChain tempkeychain = new InfilKeyChain();
        try {
          excelWorksheet.Unprotect(EncryptionPassword); // Unlock for editing
          excelWorksheet.Cells[2, 2].NumberFormat = "@"; // Ensure format is string
          if ( ! string.IsNullOrEmpty(Convert.ToString((excelWorksheet.Cells[2, 2] as Excel.Range).Value))) {
            // Pull information from Excel
            int j = 0;
            while ( ! string.IsNullOrEmpty(Convert.ToString((excelWorksheet.Cells[j + 2, 2] as Excel.Range).Value))) {
              excelWorksheet.Cells[j + 2, 2].NumberFormat = "@"; // Ensure format is string
              excelWorksheet.Cells[j + 2, 3].NumberFormat = "@"; // Ensure format is string
              excelWorksheet.Cells[j + 2, 4].NumberFormat = "@"; // Ensure format is string
              excelWorksheet.Cells[j + 2, 5].NumberFormat = "@"; // Ensure format is string 
              InfilGroupKey tempGroupkey = new InfilGroupKey();
              // If the name is populated then store it
              tempGroupkey.index = Convert.ToString((excelWorksheet.Cells[j + 2, 1] as Excel.Range).Value);
              tempGroupkey.GroupKey = Convert.ToString((excelWorksheet.Cells[j + 2, 2] as Excel.Range).Value);
              tempGroupkey.GroupName = Convert.ToString((excelWorksheet.Cells[j + 2, 3] as Excel.Range).Value);
              tempGroupkey.SponsorName = Convert.ToString((excelWorksheet.Cells[j + 2, 4] as Excel.Range).Value);
              tempGroupkey.timestamp = Convert.ToString((excelWorksheet.Cells[j + 2, 5] as Excel.Range).Value);
              tempkeychain.GroupKeys.Add(tempGroupkey);
              j++;
            }
            // Check if ID Exists
            // Check ID
            // Check if Group name is empty or not
            // Check if Group contains Sponsor Name
            

            // Check if ID Exists
            for (int i = 0; i < tempkeychain.GroupKeys.Count; i++) {
              // Check ID
              if (tempkeychain.GroupKeys[i].index == index.ToString()) {
                // Check if group name is empty or not
                string tempSponsorName = Convert.ToString((excelWorksheet.Cells[i + 2, 4] as Excel.Range).Value);
                if ( ! string.IsNullOrEmpty(tempSponsorName)) {
                  // Check if Group contains Sponsor Name
                  if (tempSponsorName.Contains(sponsorName)) {
                    // Check if Sponsor Name is last Sponsor Name in Group
                    if (tempSponsorName.Contains("/")) {
                      // Update Sponsor Name
                      string tempString = "/" + sponsorName;
                      excelWorksheet.Cells[i + 2, 4].NumberFormat = "@";
                      excelWorksheet.Cells[i + 2, 4].Value = tempSponsorName.Replace(tempString, "");

                      // Update associationValues
                      string associationValues = "";
                      if ( ! string.IsNullOrEmpty(Convert.ToString((excelWorksheet.Cells[i + 2, 6] as Excel.Range).Value))) {
                        associationValues = Convert.ToString((excelWorksheet.Cells[i + 2, 6] as Excel.Range).Value);
                      }
                      excelWorksheet.Cells[i + 2, 5].NumberFormat = "@";
                      excelWorksheet.Cells[i + 2, 6].NumberFormat = "@";
                      excelWorksheet.Cells[i + 2, 5].Value = DateTime.Now.ToString("yyyy.MMMM.dd(HH:mm:ss)");
                      if (associationValues != "") {
                        excelWorksheet.Cells[i + 2, 6].Value = associationValues + "/" + tempkeychain.GroupKeys[i].timestamp + "_" + tempkeychain.GroupKeys[i].GroupName + "[" + tempkeychain.GroupKeys[i].SponsorName + "]";
                      } else {
                        excelWorksheet.Cells[i + 2, 6].Value = tempkeychain.GroupKeys[i].timestamp + "_" + tempkeychain.GroupKeys[i].GroupName + "[" + tempkeychain.GroupKeys[i].SponsorName + "]";
                      }

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
                      string msgwindow = "ACTION COMPLETED: Sponsor Name has been removed.";
                      string capwindow = "Group Key Action";
                      MessageBoxButtons btnwindow = MessageBoxButtons.OK;
                      MessageBoxIcon iconwindow = MessageBoxIcon.Asterisk;
                      MessageBox.Show(msgwindow, capwindow, btnwindow, iconwindow);

                      return;
                    } else {
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
                      string msgwindow5 = "ACTION FAILED: Cannot remove last Sponsor Name from Group.";
                      string capwindow5 = "Group Key Action";
                      MessageBoxButtons btnwindow5 = MessageBoxButtons.OK;
                      MessageBoxIcon iconwindow5 = MessageBoxIcon.Asterisk;
                      MessageBox.Show(msgwindow5, capwindow5, btnwindow5, iconwindow5);

                      return;
                    }
                  } else {
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
                    string msgwindow = "ACTION FAILED: Group does not contain Sponsor Name.";
                    string capwindow = "Group Key Action";
                    MessageBoxButtons btnwindow = MessageBoxButtons.OK;
                    MessageBoxIcon iconwindow = MessageBoxIcon.Asterisk;
                    MessageBox.Show(msgwindow, capwindow, btnwindow, iconwindow);

                    return;
                  }
                } else {
                  // Do nothing
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
                  string msgwindow = "ACTION FAILED: Group Name is null.";
                  string capwindow = "Group Key Action";
                  MessageBoxButtons btnwindow = MessageBoxButtons.OK;
                  MessageBoxIcon iconwindow = MessageBoxIcon.Asterisk;
                  MessageBox.Show(msgwindow, capwindow, btnwindow, iconwindow);

                  return;
                }
              }
            }

            string msg = "ACTION FAILED: ID does not exist in keychain!";
            string cap = "Group Key Action";
            MessageBoxButtons btn = MessageBoxButtons.OK;
            MessageBoxIcon icon = MessageBoxIcon.Asterisk;
            MessageBox.Show(msg, cap, btn, icon);

          } else {
            string msg = "ACTION FAILED: No keys exist in the keychain!";
            string cap = "Group Key Action";
            MessageBoxButtons btn = MessageBoxButtons.OK;
            MessageBoxIcon icon = MessageBoxIcon.Asterisk;
            MessageBox.Show(msg, cap, btn, icon);
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
        } catch (Exception) {
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
      }
    }

    private void ExportSponsorKeys(string sponsorName) {
      string NewEncryptionPassword = exportPassword.Text;
      string currentSheet = "ActiveGroupKeys";
      Excel.Worksheet excelWorksheet;
      EncryptionFilePathName = appDir + DefEncryptionFileName;
      if (File.Exists(EncryptionFilePathName)) {
        //---------------------------------------------------------Open Excel Document-------------------------------------------------------------
        // Open Document and set global variable for reading
        Excel.Application appExcel = new Excel.Application();
        appExcel.Visible = false;
        Excel.Workbook WB = appExcel.Workbooks.Open(EncryptionFilePathName, ReadOnly: false, Password: EncryptionPassword);
        Excel.Sheets excelSheets = WB.Worksheets;
        currentSheet = "ActiveGroupKeys";
        excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
        InfilKeyChain tempkeychain = new InfilKeyChain();
        List<string> LocalGroupNames = new List<string>();
        try {
          excelWorksheet.Unprotect(EncryptionPassword); // Unlock for editing
          excelWorksheet.Cells[2, 2].NumberFormat = "@"; // Ensure format is string
          if ( ! string.IsNullOrEmpty(Convert.ToString((excelWorksheet.Cells[2, 2] as Excel.Range).Value))) {
            // Pull information from Excel
            int j = 0;
            while ( ! string.IsNullOrEmpty(Convert.ToString((excelWorksheet.Cells[j + 2, 2] as Excel.Range).Value))) {
              excelWorksheet.Cells[j + 2, 2].NumberFormat = "@"; // Ensure format is string
              excelWorksheet.Cells[j + 2, 3].NumberFormat = "@"; // Ensure format is string
              excelWorksheet.Cells[j + 2, 4].NumberFormat = "@"; // Ensure format is string
              excelWorksheet.Cells[j + 2, 5].NumberFormat = "@"; // Ensure format is string 
              InfilGroupKey tempGroupkey = new InfilGroupKey();
              // If the name is populated then store it
              string existingSponsorName = Convert.ToString((excelWorksheet.Cells[j + 2, 4] as Excel.Range).Value);
              if (existingSponsorName.Contains(sponsorName)) {
                tempGroupkey.index = Convert.ToString((excelWorksheet.Cells[j + 2, 1] as Excel.Range).Value);
                tempGroupkey.GroupKey = Convert.ToString((excelWorksheet.Cells[j + 2, 2] as Excel.Range).Value);
                tempGroupkey.GroupName = Convert.ToString((excelWorksheet.Cells[j + 2, 3] as Excel.Range).Value);
                tempGroupkey.SponsorName = sponsorName;
                tempGroupkey.timestamp = Convert.ToString((excelWorksheet.Cells[j + 2, 5] as Excel.Range).Value);
                tempkeychain.GroupKeys.Add(tempGroupkey);
              }
              
              j++;
            }

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

            try {
              //---------------------------------------------------------Open Excel Document-------------------------------------------------------------
              // Open Document and set global variable for reading
              appExcel = new Excel.Application();
              appExcel.Visible = false;
              WB = appExcel.Workbooks.Add();
              excelSheets = WB.Worksheets;
              excelWorksheet = (Excel.Worksheet)excelSheets[1];
              excelWorksheet.Name = "ActiveGroupKeys";

              excelWorksheet.Cells[1, 1].NumberFormat = "@";
              excelWorksheet.Cells[1, 2].NumberFormat = "@";
              excelWorksheet.Cells[1, 3].NumberFormat = "@";
              excelWorksheet.Cells[1, 4].NumberFormat = "@";
              excelWorksheet.Cells[1, 5].NumberFormat = "@";
              excelWorksheet.Cells[1, 6].NumberFormat = "@";
              excelWorksheet.Cells[1, 1].Value = "Index";
              excelWorksheet.Cells[1, 2].Value = "Group Key";
              excelWorksheet.Cells[1, 3].Value = "Name";
              excelWorksheet.Cells[1, 4].Value = "Sponsor";
              excelWorksheet.Cells[1, 5].Value = "Date Currently Associated";
              excelWorksheet.Cells[1, 6].Value = "Previous Associations";

              for (int i = 0; i < tempkeychain.GroupKeys.Count; i++) {
                excelWorksheet.Cells[i + 2, 1].NumberFormat = "@";
                excelWorksheet.Cells[i + 2, 2].NumberFormat = "@";
                excelWorksheet.Cells[i + 2, 3].NumberFormat = "@";
                excelWorksheet.Cells[i + 2, 4].NumberFormat = "@";
                excelWorksheet.Cells[i + 2, 5].NumberFormat = "@";
                excelWorksheet.Cells[i + 2, 1].Value = tempkeychain.GroupKeys[i].index;
                excelWorksheet.Cells[i + 2, 2].Value = tempkeychain.GroupKeys[i].GroupKey;
                excelWorksheet.Cells[i + 2, 3].Value = tempkeychain.GroupKeys[i].GroupName;
                excelWorksheet.Cells[i + 2, 4].Value = tempkeychain.GroupKeys[i].SponsorName;
                excelWorksheet.Cells[i + 2, 5].Value = tempkeychain.GroupKeys[i].timestamp;

                if ( ! LocalGroupNames.Contains(tempkeychain.GroupKeys[i].GroupName)) {
                  LocalGroupNames.Add(tempkeychain.GroupKeys[i].GroupName);
                }
              }
              excelWorksheet.UsedRange.Columns.AutoFit();

              excelWorksheet.Protect(NewEncryptionPassword); // Lock Editing
              excelWorksheet = (Excel.Worksheet)excelSheets.Add(After: excelWorksheet);
              excelWorksheet.Name = "Sponsor Names";

              excelWorksheet.Cells[1, 1].NumberFormat = "@";
              excelWorksheet.Cells[1, 1].Value = sponsorName;
              excelWorksheet.UsedRange.Columns.AutoFit();

              excelWorksheet.Protect(NewEncryptionPassword); // Lock Editing
              currentSheet = "ActiveGroupKeys";
              excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
              excelWorksheet.Select(Type.Missing);

              WB.Password = NewEncryptionPassword;
              WB.Protect(NewEncryptionPassword); // Lock Editing
              WB.SaveAs(appDir + "\\Export\\" + sponsorName + ".xlsx");
              appExcel.Workbooks.Close();
              appExcel.Quit();
              GC.Collect();
              GC.WaitForPendingFinalizers();
              System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
              System.Runtime.InteropServices.Marshal.FinalReleaseComObject(WB);
              System.Runtime.InteropServices.Marshal.FinalReleaseComObject(appExcel);
              string msg = "ACTION COMPLETED: All keys with Sponsor Name: " + sponsorName + " have been exported to a file of that name in the Export directory";
              string cap = "Key Export Action";
              MessageBoxButtons btn = MessageBoxButtons.OK;
              MessageBoxIcon icon = MessageBoxIcon.Asterisk;
              MessageBox.Show(msg, cap, btn, icon);

              return;
            } catch (Exception ex) {
              excelWorksheet.Protect(NewEncryptionPassword); // Lock Editing
              WB.Password = NewEncryptionPassword;
              WB.Protect(NewEncryptionPassword); // Lock Editing
              WB.SaveAs(appDir + "\\Export\\" + sponsorName + ".xlsx");
              appExcel.Workbooks.Close();
              appExcel.Quit();
              GC.Collect();
              GC.WaitForPendingFinalizers();
              System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
              System.Runtime.InteropServices.Marshal.FinalReleaseComObject(WB);
              System.Runtime.InteropServices.Marshal.FinalReleaseComObject(appExcel);
              string msg = "UNEXPECTED ERROR: " + ex;
              string cap = "Unknown";
              MessageBoxButtons btn = MessageBoxButtons.OK;
              MessageBoxIcon icon = MessageBoxIcon.Asterisk;
              MessageBox.Show(msg, cap, btn, icon);

              return;
            }
          } else {
            string msg = "ACTION FAILED: No keys exist in the keychain!";
            string cap = "Group Key Action";
            MessageBoxButtons btn = MessageBoxButtons.OK;
            MessageBoxIcon icon = MessageBoxIcon.Asterisk;
            MessageBox.Show(msg, cap, btn, icon);

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

            return;
          }
        } catch (Exception ex) {
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
          string msg = "UNEXPECTED ERROR: " + ex;
          string cap = "Unknown";
          MessageBoxButtons btn = MessageBoxButtons.OK;
          MessageBoxIcon icon = MessageBoxIcon.Asterisk;
          MessageBox.Show(msg, cap, btn, icon);

          return;
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

        return;
      }			
    }

    #endregion
  }
}

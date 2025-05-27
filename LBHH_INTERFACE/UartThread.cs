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
    private void ReceiveDiagData(BackgroundWorker bw) {
      /*
       * Background worker thread to receive mobile orginated WF diagnostic data
       * on hardwire interface.  It should be started when a serial port (COMx) is
       * opened but device is not in "connected" state.  It should be stopped when
       * a device is connected (or application is closed).
       * Stop this thread by setting runDiagThread = false
       */
      string diagStr = "\n" + DateTime.UtcNow.ToString("HH':'mm':'ss.fff");
      diagStr += " STARTING THREAD 'Receive Debug Data'\n";
      bw.ReportProgress((int)BackgroundUpdate.WriteLtToRtb, diagStr);

      char[] diagData = new char[0];         // data read from the serial port COMx
      int bytesToRead = 0;

      while (runDiagThread && comPort.IsOpen)// run main loop of the Recieve Diag Data thread
      {
        diagEvent.WaitOne(50, false);        // wait for event or check every 1/20 second for received diag data
                                             // 1/20 second would allow about 48 characters max to be buffered
        if (runDiagThread && comPort.IsOpen)  // Ensure runDiagThread didn't change while thread was blocking/waiting
        {
          try {
            // Acquire exclusive lock on serial port, read available data, and release lock
            Monitor.Enter(comPort);
            try {
              bytesToRead = comPort.BytesToRead;
              if (0 < bytesToRead) {
                diagData = new char[bytesToRead];
                comPort.Read(diagData, 0, bytesToRead);
              }
            } catch (Exception ex) {
              runDiagThread = false;       // Do NOT let diag thread continue to run
              Thread.Sleep(0);
              string msg = "Error reading diagnostic data from serial port.\nException: " + ex.ToString();
              string cap = " COM Port ";
              MessageBoxButtons btn = MessageBoxButtons.OK;
              MessageBoxIcon icon = MessageBoxIcon.Exclamation;
              bw.ReportProgress((int)BackgroundUpdate.WriteLtToRtb,
                                (DateTime.UtcNow.ToString("HH':'mm':'ss.fff") + cap + msg + "\n"));
              MessageBox.Show(msg, cap, btn, icon);
              bw.ReportProgress((int)BackgroundUpdate.MsgThreadFatality);
            }
            Monitor.Exit(comPort);

            if (0 < diagData.Length) {
              diagStr = "";
              foreach (char diagChar in diagData) {
                diagStr += diagChar.ToString();
              }
              bw.ReportProgress((int)BackgroundUpdate.WriteLtToRtb, diagStr);
              diagData = new char[0];
            }
          } catch (Exception ex) {
            string msg = "Error processing Diagnostic data.\nException: " + ex.ToString();
            bw.ReportProgress((int)BackgroundUpdate.WriteLtToRtb,
                              (DateTime.UtcNow.ToString("HH':'mm':'ss.fff") + msg + "\n"));
            bw.ReportProgress((int)BackgroundUpdate.ShowModelessNotification, msg);
          } // end try-catch reading & processing serial port data

        }//end if runDiagThread

        diagEvent.Reset();//  await event at top of while loop

      } // end while runDiagThread

      diagStr = "\n\n" + System.DateTime.UtcNow.ToString("HH':'mm':'ss.fff");
      diagStr += " TERMINATING THREAD 'Receive Debug Data'\n\n";
      bw.ReportProgress((int)BackgroundUpdate.WriteLtToRtb, diagStr);
    } // end method ReceiveDiagData
  }
}

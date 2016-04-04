using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace SSMSQuotes_Log
{
    [ComVisible(true)]
    public partial class LogWindow : UserControl
    {
        public LogWindow()
        {
            InitializeComponent();
            LogMessage(string.Format("Log Window Loaded .NET {0} - {1}", Environment.Version, System.Reflection.Assembly.GetExecutingAssembly().Location));
        }

        public void LogMessage(string msg)
        {
            rtbLog.AppendText(string.Format("[{0:HH:mm:ss}]: {1}{2}", DateTime.Now, msg, System.Environment.NewLine));
        }

        public void LogRawMessage(string msg)
        {
            rtbLog.AppendText(msg);
            rtbLog.SelectionStart = rtbLog.TextLength;
            rtbLog.ScrollToCaret();
        }
    }
}

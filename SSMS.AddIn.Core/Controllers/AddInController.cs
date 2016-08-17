using EnvDTE;
using EnvDTE80;
using Extensibility;
using Microsoft.VisualStudio.CommandBars;
using SSMS.AddInHelp.Core.Interfaces;
using SSMSQuotes_Log;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace SSMS.AddInHelp.Core.Controller
{
    public class AddInController : IAddInController
    {
        [DllImport("user32.dll")]
        public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        IAddInAdapter _adapter;
        private AddIn _addInInstance;
        private DTE2 _applicationObject;
        IToolWindowContext _logWindow = null;
        LogWindow _logWindowControl = null;
        string _workingDirectory;
        string _uiDLL;
        string _logFileName;

        public AddInController(IAddInAdapter adapter)
        {
            _adapter = adapter;
            _adapter.OnMessage += new Adapter.AddInMessageEvent(_adapter_OnMessage);
            _adapter.OnWarning += new Adapter.AddInMessageEvent(_adapter_OnWarning);
            _adapter.OnError += new Adapter.AddInMessageEvent(_adapter_OnError);
            _adapter.OnBeforeExecute += new EnvDTE._dispCommandEvents_BeforeExecuteEventHandler(_adapter_OnBeforeExecute);
            _adapter.OnAfterExecute += new EnvDTE._dispCommandEvents_AfterExecuteEventHandler(_adapter_OnAfterExecute);

            _workingDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            _uiDLL = Path.Combine(_workingDirectory, "SSMSQuotes_Log.dll");
            _adapter.CommonUIAssemblyLocation = _uiDLL;
        }

        private void _adapter_OnMessage(object sender, string msg, MessageType msgType)
        {
            LogMessage(msg);
        }

        private void _adapter_OnWarning(object sender, string warning, MessageType msgType)
        {
            LogWarning(warning);
        }

        private void _adapter_OnError(object sender, string error, MessageType msgType)
        {
            LogError(error);
        }

        private void _adapter_OnAfterExecute(string Guid, int ID, object CustomIn, object CustomOut)
        {
            try
            {
                //AddInController.CommandEvents_AfterExecute(Guid, ID, CustomIn, CustomOut);
            }
            catch (Exception ex)
            {
                LogError("Error AddInController.CommandEvents_AfterExecute: " + ex.Message);
            }
        }

        void _adapter_OnBeforeExecute(string Guid, int ID, object CustomIn, object CustomOut, ref bool CancelDefault)
        {
            try
            {
                //AddInController.CommandEvents_BeforeExecute(Guid, ID, CustomIn, CustomOut, ref CancelDefault);
            }
            catch (Exception ex)
            {
                LogError("Error AddInController.CommandEvents_BeforeExecute: " + ex.Message);
            }
        }

        #region Logging

        private void CreateLogWindow()
        {
            _logWindow = _adapter.CreateCommonUIControl("SSMSQuotes_Log.LogWindow", "SSMSQuotes Log window", new Guid("3ADC13FF-DCF4-4C49-B2EF-3D78DECDC664"));
            _logWindowControl = ((LogWindow)_logWindow.ControlObject);
        }

        private void InternalLogMessage(string msg, MessageType msgType)
        {
            LogMessageToFile(msg, msgType);
            LogMessageToWindow(msg, msgType);
        }

        private void LogWarning(string warning)
        {
            InternalLogMessage(warning, MessageType.Warning);
        }

        private void LogMessage(string msg)
        {
            InternalLogMessage(msg, MessageType.Message);
        }

        private void LogError(string error)
        {
            LogMessageToFile(error, MessageType.Error);
            if (LogMessageToWindow(error, MessageType.Error))
            {
                return;
            }
            MessageBox.Show(error);
        }

        private string FormatMessage(string msg, MessageType msgType)
        {
            return string.Format("[{0:HH:mm:ss}][{1}]: {2}", DateTime.Now, MessageTypeToString(msgType), msg);
        }

        private bool LogMessageToWindow(string msg, MessageType msgType)
        {
            if (_logWindowControl != null)
            {
                string msgToLog = FormatMessage(msg, msgType) + System.Environment.NewLine;
                _logWindowControl.LogRawMessage(msgToLog);
                return true;
            }
            return false;
        }

        private void LogMessageToFile(string msg, MessageType msgType)
        {
            System.IO.StreamWriter sw = System.IO.File.AppendText(_logFileName);
            try
            {
                sw.WriteLine(FormatMessage(msg, msgType));
            }
            finally
            {
                sw.Close();
            }
        }

        public static string MessageTypeToString(MessageType msgType)
        {
            switch (msgType)
            {
                case MessageType.Error:
                    return "Error";
                case MessageType.Message:
                    return "Message";
                case MessageType.Warning:
                    return "Warning";
                default:
                    return "Unknown";
            }
        }

        private void CreateLogFile()
        {
            try
            {
                _logFileName = Path.Combine(Application.LocalUserAppDataPath, "SSMSQuotes.log"); //TODO allow this to be configurable
                LogMessage("Log File Created");
            }
            catch (Exception e)
            {
                LogError("Error CreateLogFile: " + e.Message);
            }
        }

        private void InitLogWindow()
        {
            try
            {
                CreateLogWindow();
                string assemblyName = Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                LogMessage(string.Format("Controller Loaded: {0}", _adapter.ToString()));
            }
            catch (Exception ex)
            {
                LogError("Error InitLogWindow: " + ex.Message);
            }
        }

        #endregion

        #region AddIn Interface Methods
        
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                _addInInstance = (AddIn)addInInst;
                _applicationObject = (DTE2)_addInInstance.DTE;

                CreateLogFile();
                InitLogWindow();

                try
                {
                    if (connectMode == ext_ConnectMode.ext_cm_Startup)
                    {
                        AddMenuItems();
                    }
                }
                catch (Exception ex)
                {
                    LogError("Error AddInController.OnConnection.AddMenuItems: " + ex.Message);
                }

                LogMessage("Controller Connected");
            }
            catch (Exception ex)
            {
                LogError("Error AddInController.OnConnection: " + ex.Message);
            }
        }

        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
            try
            {
                Commands2 commands = (Commands2)_applicationObject.Commands;
                Command addinCommand = commands.Item("SSMSQuotes.Connect.AddQuotes");
                addinCommand.Delete();
                Command addinCommand2 = commands.Item("SSMSQuotes.Connect.ReplaceQuotes");
                addinCommand2.Delete();
                Command cmdLogFile = commands.Item("SSMSQuotes.Connect.LogWindow");
                cmdLogFile.Delete();
            }
            catch (Exception ex)
            {
                LogError("Error AddInController.OnDisconnection: " + ex.Message);
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            try
            {

            }
            catch (Exception ex)
            {
                LogError("Error AddInController.OnAddInsUpdate: " + ex.Message);
            }
        }

        public void OnStartupComplete(ref Array custom)
        {
            try
            {
                if (_logWindow != null)
                    _logWindow.Window.Visible = true;
            }
            catch (Exception ex)
            {
                LogError("Error AddInController.OnStartupComplete: " + ex.Message);
            }
        }

        public void OnBeginShutdown(ref Array custom)
        {
            try
            {

            }
            catch (Exception ex)
            {
                LogError("Error AddInController.OnBeginShutdown: " + ex.Message);
            }
        }

        public void Exec(string CmdName, vsCommandExecOption ExecuteOption, ref object VariantIn, ref object VariantOut, ref bool Handled)
        {
            try
            {
                Handled = false;
                if (ExecuteOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
                {
                    switch (CmdName)
                    {
                        case "SSMSQuotes.Connect.AddQuotes":
                            LogMessage("Selected SSMSQuotes - AddQuotes Menu Item.");
                            AddQuotes();
                            Handled = true;
                            return;
                        case "SSMSQuotes.Connect.ReplaceQuotes":
                            LogMessage("Selected SSMSQuotes - ReplaceQuotes Menu Item.");
                            ReplaceQuotes();
                            Handled = true;
                            return;
                        case "SSMSQuotes.Connect.LogWindow":
                            LogMessage(string.Format("Selected LogWindow Menu Item. Setting visibility to {0}", (!_logWindow.Window.Visible).ToString()));
                            ShowHideLogwindow();
                            Handled = true;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("Error AddInController.Exec: " + ex.Message);
            }
        }

        public void QueryStatus(string CmdName, vsCommandStatusTextWanted NeededText, ref vsCommandStatus StatusOption, ref object CommandText)
        {
            if (NeededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
            {
                switch (CmdName)
                {
                    case "SSMSQuotes.Connect.AddQuotes":
                    case "SSMSQuotes.Connect.ReplaceQuotes":
                    case "SSMSQuotes.Connect.LogWindow":
                        StatusOption = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                        return;
                }
            }
        }

        #endregion

        private void AddMenuItems()
        {
            object[] contextGUIDS = new object[] { };
            Commands2 commands = (Commands2)_applicationObject.Commands; //TODO Add options menu of some sort
            string toolsMenuName = "Tools";

            //Find the MenuBar command bar, which is the top-level command bar holding all the main menu items:
            Microsoft.VisualStudio.CommandBars.CommandBar menuBarCommandBar = ((Microsoft.VisualStudio.CommandBars.CommandBars)_applicationObject.CommandBars)["MenuBar"];

            //Find the Tools command bar on the MenuBar command bar:
            CommandBarControl toolsControl = menuBarCommandBar.Controls[toolsMenuName];
            CommandBarPopup toolsPopup = (CommandBarPopup)toolsControl;

            CommandBarPopup myAddinMenu = menuBarCommandBar.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, Type.Missing, toolsControl.Index, true) as CommandBarPopup;
            myAddinMenu.Caption = "SSMSQuotes";

            //Code for the shortcut keys
            //EX: Global::ALT+A
            object[] bindings = new object[1];
            bindings[0] = "Text Editor::F8";

            try
            {
                Command cmdQuotes = commands.AddNamedCommand2(_addInInstance, "AddQuotes", "Add Quotes", "Adds quotes around the current selection", true, 59, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled, (int)vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton);

                if ((cmdQuotes != null) && (myAddinMenu != null))
                {
                    cmdQuotes.Bindings = bindings;
                    cmdQuotes.AddControl(myAddinMenu.CommandBar, 1);
                }
            }
            catch (System.ArgumentException)
            {
                //Command already exists, do not create it
            }

            try
            {
                Command cmdQuotes = commands.AddNamedCommand2(_addInInstance, "ReplaceQuotes", "Add Escape Quotes", "Adds escape quotes to all single quotes in selected text", true, 59, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled, (int)vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton);

                if ((cmdQuotes != null) && (myAddinMenu != null))
                {
                    cmdQuotes.AddControl(myAddinMenu.CommandBar, 2);
                }
            }
            catch (System.ArgumentException)
            {
                //Command already exists, do not create it
            }

            try
            {
                Command cmdLogWindow = commands.AddNamedCommand2(_addInInstance, "LogWindow", "Show/Hide Log Window", "Show or hide the log window", true, 12, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled, (int)vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton);

                if ((cmdLogWindow != null) && (myAddinMenu != null))
                {
                    cmdLogWindow.AddControl(myAddinMenu.CommandBar, 3);
                }
            }
            catch (System.ArgumentException)
            {
                //Command already exists, do not create it
            }
        }

        #region Menu Items

        private void AddQuotes()
        {
            string selectedText = "";
            try
            {
                selectedText = ((TextSelection)_applicationObject.ActiveWindow.Selection).Text;
            }
            catch (Exception ex)
            {
                LogMessage("No text selected.");
                LogError("Error - AddQuotes: " + ex.Message);
                return;
            }

            if (selectedText != null || selectedText != string.Empty)
            {
                LogMessage("Text is selected");

                string finalText = "";

                string[] textLines = selectedText.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < textLines.Length - 1; i++)
                {
                    finalText += "'" + textLines[i] + "', " + Environment.NewLine;
                }
                finalText += "'" + textLines[textLines.Length - 1] + "'";

                //Replace the text in the currently selected section
                ((TextSelection)_applicationObject.ActiveWindow.Selection).Insert(finalText, (int)EnvDTE.vsInsertFlags.vsInsertFlagsContainNewText);
            }
        }

        private void ReplaceQuotes()
        {
            string selectedText = "";
            try
            {
                selectedText = ((TextSelection)_applicationObject.ActiveWindow.Selection).Text;
            }
            catch (Exception ex)
            {
                LogMessage("No text selected.");
                LogError("Error - ReplaceQuotes: " + ex.Message);
                return;
            }

            if (selectedText != null || selectedText != string.Empty)
            {
                LogMessage("Text is selected");

                string finalText = "";

                finalText = "'" + selectedText.Replace("'", "''") + "'";

                //Replace the text in the currently selected section
                ((TextSelection)_applicationObject.ActiveWindow.Selection).Insert(finalText, (int)EnvDTE.vsInsertFlags.vsInsertFlagsContainNewText);
            }
        }

        private void ShowHideLogwindow()
        {
            if (_logWindow != null)
                _logWindow.Window.Visible = !_logWindow.Window.Visible;
        }

        #endregion
    }
}

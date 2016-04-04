using EnvDTE;
using EnvDTE80;
using SSMS.AddInHelp.Core.Interfaces;
using SSMS.AddInHelp.Core.Tools;
using SSMSQuotes_Log;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SSMS.AddInHelp.Core.Adapter
{
    public delegate void AddInMessageEvent (object sender, string msg, MessageType messageType);
    public class AddInAdapter_2008 : IAddInAdapter
    {
        private DTE2 _applicationObject;

        public DTE2 ApplicationObject
        {
            get { return _applicationObject; }
        }

        private EnvDTE.AddIn _addInInstance;

        public EnvDTE.AddIn AddInInstance
        {
            get { return _addInInstance; }
        }

        CommandEvents _commandEvents;

        public CommandEvents CommandEvents
        {
            get { return _commandEvents; }
        }

        public event AddInMessageEvent OnMessage;
        public event AddInMessageEvent OnWarning;
        public event AddInMessageEvent OnError;
        public event _dispCommandEvents_BeforeExecuteEventHandler OnBeforeExecute;
        public event _dispCommandEvents_AfterExecuteEventHandler OnAfterExecute;

        public AddInAdapter_2008(object addInInst)
        {
            _addInInstance = (EnvDTE.AddIn)addInInst;
            _applicationObject = (DTE2)_addInInstance.DTE;
            _commandEvents = _applicationObject.Events.get_CommandEvents("{52692960-56BC-4989-B5D3-94C47A513E8D}", 1); 
            _commandEvents.AfterExecute += new _dispCommandEvents_AfterExecuteEventHandler(_CommandEvents_AfterExecute);
            _commandEvents.BeforeExecute += new _dispCommandEvents_BeforeExecuteEventHandler(_CommandEvents_BeforeExecute);
        }

        private void _CommandEvents_AfterExecute(string Guid, int ID, object CustomIn, object CustomOut)
        {
            if (OnAfterExecute != null)
            {
                OnAfterExecute(Guid, ID, CustomIn, CustomOut);
            }
        }

        private void _CommandEvents_BeforeExecute(string Guid, int ID, object CustomIn, object CustomOut, ref bool CancelDefault)
        {
            if (OnBeforeExecute != null)
            {
                OnBeforeExecute(Guid, ID, CustomIn, CustomOut, ref CancelDefault);
            }
        }

        public void LogMessage(string msg)
        {
            if (OnMessage != null)
            {
                OnMessage(this, msg, MessageType.Message);
            }
        }

        public void LogWarning(string warning)
        {
            if (OnWarning != null)
            {
                OnWarning(this, warning, MessageType.Warning);
            }
        }

        public void LogError(string error)
        {
            if (OnError != null)
            {
                OnError(this, error, MessageType.Error);
            }
        }

        public string CommonUIAssemblyLocation { get; set; }

        public IToolWindowContext CreateCommonUIControl(string typeName, string caption, Guid guid)
        {
            return CreateToolWindow(CommonUIAssemblyLocation, typeName, caption, guid);
        }

        public void HostWindow(ControlHost hostWindow, Control controlToHost)
        {
            hostWindow.HostChildControl(new ControlContext(controlToHost));
        }

        public IToolWindowContext CreateHostWindow(Control controlToHost, string caption, string guid)
        {
            controlToHost.Dock = DockStyle.Fill;
            controlToHost.Visible = true;
            var window = CreateCommonUIControl("SSMSQuotes_Log.ControlHost", caption, new Guid(guid));
            HostWindow(((ControlHost)window.ControlObject), controlToHost);
            return window;
        }

        public IToolWindowContext CreateToolWindow(string assemblyLocation, string typeName, string caption, Guid uiTypeGuid)
        {
            Windows2 win2 = _applicationObject.Windows as Windows2;
            if (win2 != null)
            {
                object controlObject = null;

                Window toolWindow = win2.CreateToolWindow2(_addInInstance, assemblyLocation, typeName, caption, "{" + uiTypeGuid + "}", ref controlObject);
                Window2 toolWindow2 = (Window2)toolWindow;
                toolWindow.Linkable = false;

                try
                {
                    toolWindow.WindowState = vsWindowState.vsWindowStateMaximize;
                }
                catch
                {
                }

                toolWindow.Visible = true;

                return new ToolWindowContext(toolWindow, toolWindow2, controlObject);
            }
            return null;
        }

        public void LockAround(object lockObject, Action<object> callback, object arg)
        {
            lock (lockObject)
            {
                callback(arg);
            }
        }
    }
}

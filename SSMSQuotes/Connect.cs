using System;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using System.Resources;
using System.Reflection;
using System.Globalization;
using SSMS.AddInHelp.Core;
using SSMS.AddInHelp.Core.Interfaces;
using System.Windows.Forms;
using SSMS.AddInHelp.Core.Adapter;
using SSMS.AddInHelp.Core.Controller;

namespace SSMSQuotes
{
	/// <summary>The object for implementing an Add-in.</summary>
	/// <seealso class='IDTExtensibility2' />
	public class Connect : IDTExtensibility2, IDTCommandTarget
	{
        IAddInController _controller;
        IAddInAdapter _adapter;
        private AddIn _addInInstance;
        private DTE2 _applicationObject;

        int instanceID = 1;
		/// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
		public Connect()
		{
		}

		/// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
		/// <param term='application'>Root object of the host application.</param>
		/// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
		/// <param term='addInInst'>Object representing this Add-in.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
            _addInInstance = (AddIn)addInInst;
            _applicationObject = (DTE2)_addInInstance.DTE;

            switch (_applicationObject.RegistryRoot.ToString())
            {
                case SSMS.AddInHelp.Core.Tools.SSMSVersions.SSMS2008:
                    _adapter = new AddInAdapter_2008(addInInst);
                    break;
                case SSMS.AddInHelp.Core.Tools.SSMSVersions.SSMS2012:
                    break;
                case SSMS.AddInHelp.Core.Tools.SSMSVersions.SSMS2014:
                    _adapter = new AddInAdapter_2014(addInInst);
                    break;
                default:
                    ShowMessage("Error Incompatible SSMS Version " + _applicationObject.RegistryRoot.ToString());
                    return;
            }

            try
            {
                //TODO check the version of SSMS we are running against and load different Adapters and Controllers based on version.
                _controller = new AddInController(_adapter);
                _controller.OnConnection(application, connectMode, addInInst, ref custom);
            }
            catch (Exception ex)
            {
                ShowMessage("Error Connect.OnConnection: " + ex.Message);
            }
		}

		/// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
		/// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
		{
            _controller.OnDisconnection(disconnectMode, ref custom);
		}

		/// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />		
		public void OnAddInsUpdate(ref Array custom)
		{
            _controller.OnAddInsUpdate(ref custom);
		}

		/// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref Array custom)
		{
            _controller.OnStartupComplete(ref custom);
		}

		/// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref Array custom)
		{
            _controller.OnBeginShutdown(ref custom);
		}
		
		/// <summary>Implements the QueryStatus method of the IDTCommandTarget interface. This is called when the command's availability is updated</summary>
		/// <param term='commandName'>The name of the command to determine state for.</param>
		/// <param term='neededText'>Text that is needed for the command.</param>
		/// <param term='status'>The state of the command in the user interface.</param>
		/// <param term='commandText'>Text requested by the neededText parameter.</param>
		/// <seealso class='Exec' />
		public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
		{
            _controller.QueryStatus(commandName, neededText, ref status, ref commandText);
		}

		/// <summary>Implements the Exec method of the IDTCommandTarget interface. This is called when the command is invoked.</summary>
		/// <param term='commandName'>The name of the command to execute.</param>
		/// <param term='executeOption'>Describes how the command should be run.</param>
		/// <param term='varIn'>Parameters passed from the caller to the command handler.</param>
		/// <param term='varOut'>Parameters passed from the command handler to the caller.</param>
		/// <param term='handled'>Informs the caller if the command was handled or not.</param>
		/// <seealso class='Exec' />
		public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
		{
            _controller.Exec(commandName, executeOption, ref varIn, ref varOut, ref handled);
		}

        private void ShowMessage(string msg)
        {
            MessageBox.Show(string.Format("[{0}] {1}", instanceID, msg));
        }
	}
}
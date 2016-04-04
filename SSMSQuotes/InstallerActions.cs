using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;

namespace SSMSQuotes
{
    [RunInstaller(true)]
    public partial class InstallerActions : System.Configuration.Install.Installer
    {
        public InstallerActions()
        {
            InitializeComponent();
        }

        protected override void OnAfterInstall(IDictionary savedState)
        {
 	         base.OnAfterInstall(savedState);

             string installPath = Context.Parameters["targetdir"].ToString(); //TODO maybe update the .addin file with the path for the dll.
            //TODO create the options xml file?
        }
    }
}

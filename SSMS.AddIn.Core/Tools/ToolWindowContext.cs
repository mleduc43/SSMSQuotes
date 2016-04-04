using EnvDTE;
using EnvDTE80;
using SSMS.AddInHelp.Core.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SSMS.AddInHelp.Core.Tools
{
    public class ToolWindowContext : IToolWindowContext
    {
        public Window Window { get; set; }
        public Window Window2 { get; set; }
        public object ControlObject { get; set; }
        public ToolWindowContext(Window Window, Window2 Window2, object ControlObject)
        {
            this.Window = Window;
            this.Window2 = Window2;
            this.ControlObject = ControlObject;
        }
    }
}

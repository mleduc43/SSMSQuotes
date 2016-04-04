using SSMSQuotes_Log;
using System;

namespace SSMS.AddInHelp.Core.Interfaces
{
    public interface IToolWindowContext
    {
        object ControlObject { get; set; }
        EnvDTE.Window Window { get; set; }
        EnvDTE.Window Window2 { get; set; }
    }
}

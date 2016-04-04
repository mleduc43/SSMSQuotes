using System;
using System.Runtime.InteropServices;

namespace SSMSQuotes_Log
{
    [ComVisible(true)]
    public interface IControlContext
    {
        IntPtr Handle { get; }
    }
}

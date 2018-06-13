using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Quintessence.MotionControl
{
    public class PortResponse
    {
        public string Code;
        public string Message;
        public Exception ex;
        public PortResponse(string code, string message, Exception e)
        {
            this.Code = code;
            this.Message = message;
            this.ex = e;
        }
    }
}

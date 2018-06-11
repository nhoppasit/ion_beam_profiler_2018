using System;

namespace Quintessence.Meter
{
    public struct GpibResponse
    {
        public string Code;
        public string Message;
        public SystemException ex;
        public GpibResponse(string code, string message, SystemException e)
        {
            this.Code = code;
            this.Message = message;
            this.ex = e;
        }
    }
}

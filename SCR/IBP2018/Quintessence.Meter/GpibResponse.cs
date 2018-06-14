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

        public const string SUCCESS = "00";
        public const string DEMO = "01";
        public const string ERR_PORTNAME = "PE";
        public const string ERR_OPEN = "OE";
        public const string ERR_CLOSE = "CE";
        public const string ERR_WRITE = "WE";
        public const string ERR_READ = "RE";

    }
}

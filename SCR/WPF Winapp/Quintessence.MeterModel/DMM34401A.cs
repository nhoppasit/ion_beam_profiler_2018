using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Quintessence.MeterModel
{
    public class DMM34401A
    {
        /* ----------------------------------------------------------  
         * Measured Value
         * ----------------------------------------------------------  */
        private float _Current; public float Current { get { return _Current; } set { _Current = value; } }
        private float _Voltage; public float Voltage { get { return _Voltage; } set { _Voltage = value; } }

        /* ----------------------------------------------------------  
         * Read Interval
         * ----------------------------------------------------------  */
        private int _ReadIntervalMillisecond; public int ReadIntervalMillisecond { get { return _ReadIntervalMillisecond; } set { _ReadIntervalMillisecond = value; } }

        /* ----------------------------------------------------------  
         * Read Interval
         * ----------------------------------------------------------  */


    }
}

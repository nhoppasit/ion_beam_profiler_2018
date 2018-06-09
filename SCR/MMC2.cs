using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Quintessence.MMC2Model
{
    public class MMC2
    {
        /* ----------------------------------------------------------  
         * พอร์ต
         * ----------------------------------------------------------  */
        private string _SerialPortName; public string SerialPortName { get; set; }

        /* ----------------------------------------------------------  
         * ข้อมูลตำแหน่ง
         * ----------------------------------------------------------  */
        private int _ActualXStep; public int ActualXStep { get; set; }
        private int _ActualYStep; public int ActualYStep { get; set; }
        private bool _IsReady; public bool IsReady { get; set; }

        /* ----------------------------------------------------------  
         * การเคลื่อนที่แบบสัมพัทธ์
         * ----------------------------------------------------------  */
        private int _XRelativeStep; public string XRelativeStep { get; set; }
        private int _YRelativeStep; public string YRelativeStep { get; set; }

    }
}

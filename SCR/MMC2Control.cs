using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Quintessence.MMC2Model;

namespace Quintessence.MMC2Control
{
    public class MMC2Control
    {
        private MMC2 obj = new MMC2();
        public string SerialPortName
        {
            get { return obj.SerialPortName; }
            set { obj.SerialPortName = value; }
        }
        public string ActualXStep
        {
            get { return obj.ActualXStep.ToString("0"); }
            set { obj.ActualXStep = Convert.ToInt32(value); }
        }
        public string ActualYStep
        {
            get { return obj.ActualYStep.ToString("0"); }
            set { obj.ActualYStep = Convert.ToInt32(value); }
        }

    }
}

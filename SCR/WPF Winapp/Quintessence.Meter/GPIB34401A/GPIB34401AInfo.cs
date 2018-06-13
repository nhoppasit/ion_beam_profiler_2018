using Ivi.Visa.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace Quintessence.Meter.Gpib34401a
{
    public class Gpib34401aInfo : INotifyPropertyChanged
    {
        #region INotifyPropertyChanged Implementation
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion

        /* ----------------------------------------------------------  
         * Gpib properties
         * ---------------------------------------------------------- */
        private string _GpibInterfaceId; public string GpibInterfaceId { get { return _GpibInterfaceId; } }
        private string _VisaAddress; public string VisaAddress { get { return _VisaAddress; } }
        private int _GpibBoardNumber;
        public int GpibBoardNumber
        {
            get { return _GpibBoardNumber; }
            set
            {
                _GpibBoardNumber = value;
                _GpibInterfaceId = "GPIB" + _GpibBoardNumber.ToString();
                _VisaAddress = _GpibInterfaceId + "::" + _GpibAddress.ToString() + "::INSTR";
                OnPropertyChanged("GpibBoardNumber");
                OnPropertyChanged("GpibInterfaceId");
                OnPropertyChanged("VisaAddress");
            }
        }
        private int _GpibAddress;
        public int GpibAddress
        {
            get { return _GpibAddress; }
            set
            {
                _GpibAddress = value;
                _VisaAddress = _GpibInterfaceId + "::" + _GpibAddress.ToString() + "::INSTR";
                OnPropertyChanged("GpibAddress");
                OnPropertyChanged("VisaAddress");
            }
        }

        /* ----------------------------------------------------------  
         * GPIB interface
         * formattedIO interface
         *   1. Create IO object
         *   2. Intialize meter via GPIB
         *   3. Configure meter via GPIB
         *   4. Measure Value
         * ---------------------------------------------------------- */
        private Ivi.Visa.Interop.FormattedIO488 ioDmm;
        public GpibResponse CreateIO488Object()
        {
            try
            {
                ioDmm = new FormattedIO488Class();
                GpibResponse gr = new GpibResponse("00", "OK", null);
                return gr;
            }
            catch (SystemException ex)
            {
                GpibResponse gr = new GpibResponse("EX", "FormattedIO488Class object creation failure. " + ex.Source + "  " + ex.Message, ex);
                return gr;
            }
        }
        public GpibResponse InitializeMeter()
        {
            // Verify IO 
            if (ioDmm == null)
            {
                GpibResponse gr = CreateIO488Object();
                if (!gr.Code.Equals("00")) return gr;
            }

            throw new NotImplementedException();
        }
        public GpibResponse ConfigureMeter() { throw new NotImplementedException(); }
        public GpibResponse Measure() { throw new NotImplementedException(); }

        /* ----------------------------------------------------------  
         * Measured Value
         * ----------------------------------------------------------  */
        private double _Current; public double Current { get { return _Current; } set { _Current = value; OnPropertyChanged("Current"); } }
        private double _Voltage; public double Voltage { get { return _Voltage; } set { _Voltage = value; OnPropertyChanged("Voltage"); } }

        /* ----------------------------------------------------------  
         * Read Interval
         * ----------------------------------------------------------  */
        private int _ReadIntervalMillisecond; public int ReadIntervalMillisecond { get { return _ReadIntervalMillisecond; } set { _ReadIntervalMillisecond = value; OnPropertyChanged("ReadIntervalMillisecond"); } }

        /* ----------------------------------------------------------  
         * Demo mode
         * There is no interface but can get random data 
         * for measureing demonstration
         * ----------------------------------------------------------  */
        static int demoIdForCurrent = 0;
        public double PeriodicRangeOfCurrent = 30;
        public void ResetDemoIdForCurrent() { demoIdForCurrent = 0; }
        public void GenerateNewDemoCurrent()
        {
            Current = Math.Sin(2 * Math.PI * demoIdForCurrent / PeriodicRangeOfCurrent);
            demoIdForCurrent += 1;
        }
        static int demoIdForVoltage = 0;
        public double PeriodicRangeOfVoltage = 30;
        public void ResetDemoIdForVoltage() { demoIdForVoltage = 0; }
        public void GenerateDemoVoltage()
        {
            Voltage = Math.Sin(2 * Math.PI * demoIdForVoltage / PeriodicRangeOfVoltage);
            demoIdForVoltage += 1;
        }
    }
}

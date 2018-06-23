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
        private string _GpibInterfaceId = "GPIB0";
        public string GpibInterfaceId { get { return _GpibInterfaceId; } }
        private string _VisaAddress;
        public string VisaAddress { get { return _VisaAddress; } set { _VisaAddress = value; } }
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
        private int _GpibTimeout = 7000;
        public int GpibTimeout { get { return _GpibTimeout; } set { _GpibTimeout = value; OnPropertyChanged("GpibTimeout"); } }
        private bool _IsDemo = false; public bool IsDemo { get { return _IsDemo; } set { _IsDemo = value; OnPropertyChanged("IsDemo"); } }

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
        public GpibResponse InitializeMeterForCurrent()
        {
            // Verify IO 
            if (ioDmm == null)
            {
                GpibResponse gr = CreateIO488Object();
                if (!gr.Code.Equals("00")) return gr;
            }

            try
            {
                //create the resource manager and open a session with the instrument specified on txtAddress
                ResourceManager grm = new ResourceManager();
                ioDmm.IO = (IMessage)grm.Open(_VisaAddress, AccessMode.NO_LOCK, 2000, "");
                ioDmm.IO.Timeout = 7000;
                GpibResponse gr = new GpibResponse("00", "Initialize the instrument on " + _VisaAddress, null);
                return gr;
            }
            catch (SystemException ex)
            {
                ioDmm.IO = null;
                GpibResponse gr = new GpibResponse("IE", "Open failed on " + _VisaAddress + " " + ex.Source + "  " + ex.Message, ex);
                return gr;
            }
        }
        public GpibResponse ConfigureMeterForCurrent()
        {
            try
            {
                ioDmm.WriteString("*RST", true);//Reset the dmm                
                ioDmm.WriteString("*CLS", true);//Clear the dmm registers                
                //ioDmm.WriteString("CALC:DBM:REF 50", true);//Set 50 ohm reference for dBm
                ioDmm.WriteString("Conf:CURR:DC 1, 0.001", true);// Set dmm to 1 amp ac range                
                //ioDmm.WriteString(":Det:Band 200", true);// Select the 200 Hz (fast) ac filter                
                //ioDmm.WriteString("Trig:Coun 5", true);//dmm will accept 5 triggers                
                //ioDmm.WriteString("Trig:Sour IMM", true);//Trigger source is IMMediate                
                //ioDmm.WriteString("Calc:Func DBM", true);//Select dBm function                
                //ioDmm.WriteString("Calc:Stat ON", true);//Enable math and request operation complete                
                ioDmm.WriteString("Read?", true);//Take readings; send to output buffer

                // Get readings and parse into array of doubles
                // Enter will wait until all readings are completed
                //' print to Text box
                double[] Readings = new double[5];
                string sText = "";
                Readings = (double[])ioDmm.ReadList(IEEEASCIIType.ASCIIType_R8, ",");
                for (int iIndex = 0; iIndex < Readings.Length; iIndex++)
                {
                    sText = sText + Readings[iIndex].ToString() + " A" + "\r\n";
                }
                //Current = 0.00;
                GpibResponse gr = new GpibResponse("00", "Readed", null);
                return gr;
            }
            catch (SystemException ex)
            {
                GpibResponse gr = new GpibResponse("CE", "Configure command failed. " + ex.Source + "  " + ex.Message, ex);
                return gr;
            }
        }
        public GpibResponse MeasureCurrent()
        {
            try
            {
                ioDmm.WriteString("*RST", true);//Reset the dmm                
                ioDmm.WriteString("*CLS", true);//Clear the dmm registers                
                ioDmm.WriteString("Measure:Current:DC? 1A,0.001MA", true);// Set meter to 1 amp dc range, 0.001mA resolution
                Current = (double)ioDmm.ReadNumber(IEEEASCIIType.ASCIIType_R4, true);
                GpibResponse gr = new GpibResponse(GpibResponse.SUCCESS, Current.ToString(), null);
                return gr;
            }
            catch (SystemException ex)
            {
                GpibResponse gr = new GpibResponse(GpibResponse.ERR_MEAS, "Measure current query failed. " + ex.Source + "  " + ex.Message, ex);
                return gr;
            }
        }

        /* ----------------------------------------------------------  
         * Measured Value
         * ----------------------------------------------------------  */
        private double _Current; public double Current { get { return _Current; } set { _Current = value; OnPropertyChanged("Current"); } }
        private double _Voltage; public double Voltage { get { return _Voltage; } set { _Voltage = value; OnPropertyChanged("Voltage"); } }

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

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
         * Gpib Name
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
        

        /* ----------------------------------------------------------  
         * Measured Value
         * ----------------------------------------------------------  */
        private float _Current; public float Current { get { return _Current; } set { _Current = value; OnPropertyChanged("Current"); } }
        private float _Voltage; public float Voltage { get { return _Voltage; } set { _Voltage = value; OnPropertyChanged("Voltage"); } }

        /* ----------------------------------------------------------  
         * Read Interval
         * ----------------------------------------------------------  */
        private int _ReadIntervalMillisecond; public int ReadIntervalMillisecond { get { return _ReadIntervalMillisecond; } set { _ReadIntervalMillisecond = value; OnPropertyChanged("ReadIntervalMillisecond"); } }

    }
}

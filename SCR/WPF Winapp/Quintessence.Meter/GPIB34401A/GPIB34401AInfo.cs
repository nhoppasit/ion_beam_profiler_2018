using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace Quintessence.Meter.Gpib34401a
{
    public class Gpib34401aInfo : INotifyPropertyChanged
    {
        #region Implementation
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion

        public Gpib34401aInfo(int BoardNumber, int address)
        {

        }

        /* ----------------------------------------------------------  
         * Gpib Name
         * ----------------------------------------------------------  */
        private string _GpibInterfaceId;
        public string GpibInterfaceId
        {
            get
            {
                return _GpibInterfaceId;
            }
        }
        private int _GpibBoardNumber;
        public int GpibBoardNumber
        {
            get { return _GpibBoardNumber; }
            set
            {
                _GpibBoardNumber = value;
                _GpibInterfaceId = "GPIB" + _GpibBoardNumber.ToString();
                OnPropertyChanged("GpibBoardNumber");
                OnPropertyChanged("GpibInterfaceId");
            }
        }
        private string _GpibAddress;
        public string GpibAddress
        {
            get { return _GpibAddress; }
            set
            {
                _GpibAddress = value;
                OnPropertyChanged("GpibAddress");
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

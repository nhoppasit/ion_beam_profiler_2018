using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Collections.ObjectModel;
using System.Windows.Input;
using System.Windows;

namespace Quintessence.Meter.Gpib34401a
{
    public class TwoAmmeterViewModel : INotifyPropertyChanged
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

        // Model
        private IList<Gpib34401aInfo> _Ammeters;
        public IList<Gpib34401aInfo> Ammeters { get { return _Ammeters; } set { _Ammeters = value; } }

        // ViewModel Properties 
        public int A1GpibAddress
        {
            get { return _Ammeters[0].GpibAddress; }
            set
            {
                _Ammeters[0].GpibAddress = value;
                OnPropertyChanged("A1GpibAddress");
                OnPropertyChanged("A1VisaAddress");
            }
        }

        public string A1VisaAddress
        {
            get { return _Ammeters[0].VisaAddress; }
        }

        public int A2GpibAddress
        {
            get { return _Ammeters[1].GpibAddress; }
            set
            {
                _Ammeters[1].GpibAddress = value;
                OnPropertyChanged("A2GpibAddress");
                OnPropertyChanged("A2VisaAddress");
            }
        }

        public string A2VisaAddress
        {
            get { return _Ammeters[1].VisaAddress; }
        }

        public string Current1 { get { return _Ammeters[0].Current.ToString(); } set { _Ammeters[0].Current = Convert.ToSingle(value); OnPropertyChanged("Current1"); } }
        public string Current2 { get { return _Ammeters[1].Current.ToString(); } set { _Ammeters[1].Current = Convert.ToSingle(value); OnPropertyChanged("Current2"); } }

        private bool isBusy = false;

        public ICommand InitializeMeter { get; set; }
        public ICommand ConfigureMeter { get; set; }
        public ICommand Measure { get; set; }

        public TwoAmmeterViewModel()
        {
            _Ammeters = new List<Gpib34401aInfo>();
            Gpib34401aInfo a1 = new Gpib34401aInfo();
            a1.GpibBoardNumber = 0;
            a1.GpibAddress = 26;
            a1.ReadIntervalMillisecond = 333;
            _Ammeters.Add(a1);
            Gpib34401aInfo a2 = new Gpib34401aInfo();
            a2.GpibBoardNumber = 0;
            a2.GpibAddress = 27;
            a2.ReadIntervalMillisecond = 333;
            _Ammeters.Add(a2);

            Current1 = "100.01";
            Current2 = "200.02";

            InitializeMeter = new Gpib34401aRelayCommand(ExecuteInitializeMeterMethod, CanExecuteInitializeMeterMethod);
            ConfigureMeter = new Gpib34401aRelayCommand(ExecuteConfigureMeterMethod, CanExecuteConfigureMeterMethod);
            Measure = new Gpib34401aRelayCommand(ExecuteMeasureMethod, CanExecuteMeasureMethod);
        }

        private bool CanExecuteInitializeMeterMethod(object parameter) { return !isBusy; }

        private void ExecuteInitializeMeterMethod(object parameter)
        {
            MessageBox.Show("Hello...  Ammeter");
        }

        private bool CanExecuteConfigureMeterMethod(object parameter) { return !isBusy; }

        private void ExecuteConfigureMeterMethod(object parameter)
        {
            MessageBox.Show("Configured");
        }

        private bool CanExecuteMeasureMethod(object parameter) { return !isBusy; }

        private void ExecuteMeasureMethod(object parameter)
        {
            Current1 = "111.11";
            Current2 = "222.22";
            MessageBox.Show("Measured");
        }
    }
}

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
    public class AmmeterViewModel : INotifyPropertyChanged
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
        public Gpib34401aInfo gpib34401AInfo { get; set; }

        // ViewModel Properties 
        public int GpibAddress
        {
            get { return gpib34401AInfo.GpibAddress; }
            set
            {
                gpib34401AInfo.GpibAddress = value;
                OnPropertyChanged("GpibAddress");
                OnPropertyChanged("VisaAddress");
            }
        }

        public string VisaAddress
        {
            get { return gpib34401AInfo.VisaAddress; }
        }

        private bool isBusy = false;

        public ICommand InitializeMeter { get; set; }
        public ICommand ConfigureMeter { get; set; }
        public ICommand Measure { get; set; }

        public AmmeterViewModel()
        {
            gpib34401AInfo = new Gpib34401aInfo();
            gpib34401AInfo.GpibBoardNumber = 0;
            gpib34401AInfo.GpibAddress = 26;

            InitializeMeter = new Gpib34401aRelayCommand(ExecuteInitializeMeterMethod, CanExecuteInitializeMeterMethod);
            ConfigureMeter= new Gpib34401aRelayCommand(ExecuteConfigureMeterMethod, CanExecuteConfigureMeterMethod);
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
            MessageBox.Show("Measured");
        }
    }
}

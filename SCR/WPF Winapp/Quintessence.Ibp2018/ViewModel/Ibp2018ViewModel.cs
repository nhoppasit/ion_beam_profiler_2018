using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using Quintessence.Meter;
using Quintessence.Meter.Gpib34401a;
using Quintessence.MotionControl.MMC2;
using Quintessence.WpfCommands;
using System.Windows.Input;
using System.Windows;
using System.Threading;
using System.Windows.Threading;
using GalaSoft.MvvmLight.Threading;
using GalaSoft.MvvmLight;

namespace Quintessence.Ibp2018.ViewModel
{
    public class Ibp2018ViewModel : ViewModelBase, INotifyPropertyChanged
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

        /*
         * MODELS
         *  - Two Ammeter, Gpib
         *  - Two MMC2, Serial port
         */
        private IList<Gpib34401aInfo> _Ammeters;
        public IList<Gpib34401aInfo> Ammeters { get { return _Ammeters; } set { _Ammeters = value; } }

        /*
         * Ammeter 1 Properties
         */
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

        /*
         * Ammeter 2 Properties
         */
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
        public string A2VisaAddress { get { return _Ammeters[1].VisaAddress; } }

        public string Current1Text { get { return _Ammeters[0].Current.ToString(); } set { _Ammeters[0].Current = Convert.ToSingle(value); OnPropertyChanged("Current1Text"); } }
        public string Current2 { get { return _Ammeters[1].Current.ToString(); } set { _Ammeters[1].Current = Convert.ToSingle(value); OnPropertyChanged("Current2"); } }
        public float Current1 { get { return _Ammeters[0].Current; } set { _Ammeters[0].Current = value; OnPropertyChanged("Current1Text"); } }

        private bool isBusy = false;

        public ICommand InitializeMeter { get; set; }
        public ICommand ConfigureMeter { get; set; }
        public ICommand Measure { get; set; }

        public Ibp2018ViewModel()
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

            Current1Text = "100.01";
            Current2 = "200.02";

            InitializeMeter = new DelegateCommand(ExecuteInitializeMeterMethod, CanExecuteInitializeMeterMethod);
            ConfigureMeter = new RelayCommand(ExecuteConfigureMeterMethod, CanExecuteConfigureMeterMethod);
            Measure = new RelayCommand(ExecuteMeasureMethod, CanExecuteMeasureMethod);
        }

        private bool CanExecuteInitializeMeterMethod(object parameter)
        {
            return !isBusy;
        }

        private void ExecuteInitializeMeterMethod(object parameter)
        {
            ThreadPool.QueueUserWorkItem(
                o =>
                {
                    // This is a background operation!

                    DispatcherHelper.CheckBeginInvokeOnUI(
                            () =>
                            {
                                // Dispatch back to the main thread
                                //isBusy = false;
                                Current1 += 1;
                                System.Threading.Thread.Sleep(1000);
                                GpibResponse gr = _Ammeters[0].CreateIO488Object();
                                Current1 += 1;
                                //isBusy = false;
                            });
                });
        }

        private bool CanExecuteConfigureMeterMethod(object parameter) { return !isBusy; }

        private void ExecuteConfigureMeterMethod(object parameter)
        {
            MessageBox.Show("Configured");
        }

        private bool CanExecuteMeasureMethod(object parameter) { return !isBusy; }

        private void ExecuteMeasureMethod(object parameter)
        {
            ThreadPool.QueueUserWorkItem(
                o =>
                {
                    // This is a background operation!

                    DispatcherHelper.CheckBeginInvokeOnUI(
                            () =>
                            {
                                // Dispatch back to the main thread
                                //isBusy = false;
                                Current1Text = "111.11";
                                Current2 = "222.22";
                                MessageBox.Show("Measured");
                                //isBusy = false;
                            });
                });
        }
    }
}

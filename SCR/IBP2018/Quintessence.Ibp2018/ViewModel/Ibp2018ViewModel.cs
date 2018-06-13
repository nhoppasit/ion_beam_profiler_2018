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
using System.Collections.ObjectModel;
using System.Windows.Controls;
using System.Data;
using System.Windows.Media;
using Quintessence.Ibp2018.Model;
using System.Windows.Data;
using Quintessence.MotionControl;

namespace Quintessence.Ibp2018.ViewModel
{
    public class Ibp2018ViewModel : INotifyPropertyChanged
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
         * MODELS
         *  - Two Ammeter, Gpib
         *  - Two MMC2, Serial port
         * ---------------------------------------------------------- */
        private IList<Gpib34401aInfo> _Ammeters;
        public IList<Gpib34401aInfo> Ammeters { get { return _Ammeters; } set { _Ammeters = value; } }

        /* ----------------------------------------------------------
         * Ammeter-1 Properties
         * ---------------------------------------------------------- */
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
        public string A1VisaAddressText
        {
            get { return _Ammeters[0].VisaAddress; }
        }
        public double Current1 { get { return _Ammeters[0].Current; } set { _Ammeters[0].Current = value; OnPropertyChanged("Current1Text"); } }
        public string Current1Text { get { return _Ammeters[0].Current.ToString(); } set { _Ammeters[0].Current = Convert.ToSingle(value); OnPropertyChanged("Current1Text"); } }
        private bool canMeasure = true, canPause = false, canStop = false;

        /* ----------------------------------------------------------
         * Ammeter-2 Properties
         * ---------------------------------------------------------- */
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
        public string A2VisaAddressText { get { return _Ammeters[1].VisaAddress; } }
        public double Current2 { get { return _Ammeters[1].Current; } set { _Ammeters[1].Current = value; OnPropertyChanged("Current2Text"); } }
        public string Current2Text { get { return _Ammeters[0].Current.ToString(); } set { _Ammeters[0].Current = Convert.ToSingle(value); OnPropertyChanged("Current1Text"); } }

        /* ----------------------------------------------------------
         * X-Y scanner and Z axis object
         * ---------------------------------------------------------- */
        private MMC2Info _XyMmc;
        public MMC2Info XyMmc { get { return _XyMmc; } set { _XyMmc = value; } }
        public string XyMmcPortName { get { return _XyMmc.SerialPortName; } set { _XyMmc.SerialPortName = value; OnPropertyChanged("XyMmcPortName"); } }
        public string XLPosText { get { return _XyMmc.ActualX.ToString("F2"); } }
        public string YLPosText { get { return _XyMmc.ActualY.ToString("F2"); } }
        private MMC2Info _ZMmc;
        public MMC2Info ZMmc { get { return _ZMmc; } set { _ZMmc = value; } }
        public string ZMmcPortName { get { return _ZMmc.SerialPortName; } set { _ZMmc.SerialPortName = value; OnPropertyChanged("ZMmcPortName"); } }
        public string ZLPosText { get { return _ZMmc.ActualX.ToString("F2"); } }

        /* ----------------------------------------------------------
         * Current tables
         * ---------------------------------------------------------- */
        private IList<Ibp2018DataTableModel> _CurrentTables;
        public IList<Ibp2018DataTableModel> CurrentTables { get { return _CurrentTables; } set { _CurrentTables = value; } }

        /* ----------------------------------------------------------
         * File / Current-1 Properties
         * ---------------------------------------------------------- */
        public DataTable CurrentTable1
        {
            get
            {
                return _CurrentTables[0].Datatable;
            }
            set
            {
                if (_CurrentTables[0].Datatable != value)
                {
                    _CurrentTables[0].Datatable = value;
                }
                OnPropertyChanged("CurrentTable1");
            }
        }

        /* ----------------------------------------------------------
         * ColumnCollection designed for dynamic datagrid columns
         * ---------------------------------------------------------- */
        private ObservableCollection<DataGridColumn> _columnCollection = new ObservableCollection<DataGridColumn>();
        public ObservableCollection<DataGridColumn> ColumnCollection
        {
            get
            {
                return this._columnCollection;
            }
            set
            {
                _columnCollection = value;
                OnPropertyChanged("ColumnCollection");
                //Error
                //base.OnPropertyChanged<ObservableCollection<DataGridColumn>>(() => this.ColumnCollection);
            }
        }

        /* ----------------------------------------------------------
         * Commands for binding to buttons
         * ---------------------------------------------------------- */
        public ICommand InitializeMeter { get; set; }
        public ICommand ConfigureMeter { get; set; }
        public ICommand Measure { get; set; }
        public ICommand GenerateNewDemoData { get; set; }

        // Commands of xy and z mmc reconnect
        public ICommand ReconnectXyMmc { get; set; }
        public ICommand ReconnectZMmc { get; set; }
        public ICommand QueryScanner { get; set; }
        public ICommand UnqueryScanner { get; set; }

        /* ----------------------------------------------------------
         * CONSTRUCTOR
         * ---------------------------------------------------------- */
        public Ibp2018ViewModel()
        {
            // Create meters object
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

            // Create XYZ scanner object
            _XyMmc = new MMC2Info();
            _XyMmc.SerialPortName = "COM1";
            _ZMmc = new MMC2Info();
            _ZMmc.SerialPortName = "COM2";

            // Create data tables object
            _CurrentTables = new List<Ibp2018DataTableModel>();
            _CurrentTables.Add(new Ibp2018DataTableModel());
            _CurrentTables.Add(new Ibp2018DataTableModel());

            // Initialize meter objects
            Current1Text = "100.01";
            Current2Text = "200.02";

            // Create commands
            InitializeMeter = new RelayCommand(ExecuteInitializeMeterMethod, CanExecuteInitializeMeterMethod);
            ConfigureMeter = new RelayCommand(ExecuteConfigureMeterMethod, CanExecuteConfigureMeterMethod);
            Measure = new RelayCommand(ExecuteMeasureMethod, CanExecuteMeasureMethod);
            GenerateNewDemoData = new RelayCommand(ExecuteGenerateDemoDataMethod, CanExecuteGenerateDemoDataMethod);

            // Serial port of xy and z mmc
            ReconnectXyMmc = new RelayCommand(ExecuteReconnectXyMmcMethod, CanExecuteReconnectXyMmcMethod);
            ReconnectZMmc = new RelayCommand(ExecuteReconnectZMmcMethod, CanExecuteReconnectZMmcMethod);

            // Query and unquery scanner 
            QueryScanner = new RelayCommand(ExecuteQueryScannerMethod, CanExecuteQueryScannerMethod);
            UnqueryScanner = new RelayCommand(ExecuteUnqueryScannerMethod, CanExecuteUnqueryScannerMethod);
        }

        /* ----------------------------------------------------------
         * Query and unquery scanner
         * ---------------------------------------------------------- */
        private bool canQueryScanner = true;
        private bool _continueQueryScanner = true;
        private bool CanExecuteQueryScannerMethod(object parameter) { return canQueryScanner; }
        private void ExecuteQueryScannerMethod(object parameter)
        {
            ThreadPool.QueueUserWorkItem(
                o =>
                {
                    canQueryScanner = true;
                    _continueQueryScanner = true;
                    while (_continueQueryScanner)
                    {
                        try
                        {
                            XyMmc.QueryPosition(true);
                            OnPropertyChanged("XLPosText");
                            OnPropertyChanged("YLPosText");
                            ZMmc.QueryPosition(true);
                            OnPropertyChanged("ZLPosText");
                        }
                        catch (Exception ex) { }
                        Thread.Sleep(500);
                    }
                    MessageBox.Show("Auto query turned OFF.", "Query scanner", MessageBoxButton.OK, MessageBoxImage.Information);
                });
        }
        private bool CanExecuteUnqueryScannerMethod(object parameter) { return true; }
        private void ExecuteUnqueryScannerMethod(object parameter)
        {
            _continueQueryScanner = false;
            canQueryScanner = true;
        }


        /* ----------------------------------------------------------
         * Initialize Meter
         * ---------------------------------------------------------- */
        private bool CanExecuteInitializeMeterMethod(object parameter) { return canMeasure; }
        private bool _continueInitMeter = true;
        private int scanX_Idx, scanY_Idx;
        private void ExecuteInitializeMeterMethod(object parameter)
        {
            ThreadPool.QueueUserWorkItem(
                o =>
                {
                    _continueInitMeter = true;
                    canMeasure = false;
                    canPause = true;
                    canStop = true;
                    while (_continueInitMeter)
                    {
                        try
                        {
                            Current1 += 1;
                            //DataTable dt = CurrentTable1;
                            DataRow r = CurrentTable1.Rows[scanY_Idx];
                            r[scanX_Idx + 1] = Current1;

                            Thread.Sleep(1000);
                        }
                        catch (Exception ex) { }
                    }
                });
        }

        /* ----------------------------------------------------------
         * Configure meter
         * ---------------------------------------------------------- */
        private bool CanExecuteConfigureMeterMethod(object parameter) { return canPause; }
        private void ExecuteConfigureMeterMethod(object parameter)
        {
            _continueInitMeter = false;
            canMeasure = true;
            canPause = false;
            canStop = true;
        }

        /* ----------------------------------------------------------
         * Measurement
         * ---------------------------------------------------------- */
        private bool CanExecuteMeasureMethod(object parameter) { return canStop; }
        private void ExecuteMeasureMethod(object parameter)
        {
            _continueInitMeter = false;
            Current1Text = "111.11";
            MessageBox.Show("Measured");
            canMeasure = true;
            canPause = false;
            canStop = false;
        }

        /* ----------------------------------------------------------
         * DEMO DATA
         * ---------------------------------------------------------- */
        private void GenerateDemoData()
        {
            // Wait dialog can show in view 
            // Wait for columns.count

            canDemo = false;
            // Current-1 data table
            _CurrentTables[0].GenerateNewDemoColumns(0.02, 0.02, 0, 20, 0, 20);

            // Binding columns name and header
            for (int i = 0; i < _CurrentTables[0].ColumnNames.Count; i++)
            {
                Binding binding = new Binding(_CurrentTables[0].ColumnNames[i]);
                DataGridTextColumn textColumn = new DataGridTextColumn();
                textColumn.Header = _CurrentTables[0].ColumnHeaders[i];
                textColumn.Binding = binding;
                ColumnCollection.Add(textColumn);
            }

            canDemo = true;
        }
        private bool canDemo = true;
        private bool CanExecuteGenerateDemoDataMethod(object parameter) { return canDemo; }
        public void ExecuteGenerateDemoDataMethod(object parameter) { GenerateDemoData(); }

        // Reconnect xy mmc command
        private object XyMmcLock = new object();
        public bool XyMmcReconnecting = false;
        public bool XyMmcConnected = false;
        private bool canReconnectXyMmc = true;
        private bool CanExecuteReconnectXyMmcMethod(object parameter) { return canReconnectXyMmc; }
        private void ExecuteReconnectXyMmcMethod(object parameter)
        {
            canReconnectXyMmc = false;
            lock (XyMmcLock)
            {
                try
                {
                    XyMmcReconnecting = true;
                    XyMmcConnected = false;
                    PortResponse pr = XyMmc.Connect();
                    if (pr.Code == PortResponse.SUCCESS)
                    {
                        XyMmcConnected = true;
                        XyMmcReconnecting = false;
                        MessageBox.Show("X-Y scanner mmc2 on " + XyMmc.SerialPortName + " reconnected.", "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        XyMmcConnected = true;
                        XyMmcReconnecting = false;
                        MessageBox.Show("Cannot reconnect X-Y scanner mmc2 on " + XyMmc.SerialPortName + ". " + pr.Message,
                        "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }
                catch (Exception ex)
                {
                    PortResponse pr = new PortResponse("RE", ex.Message, ex);
                    XyMmcConnected = false;
                    XyMmcReconnecting = false;
                    MessageBox.Show("Cannot reconnect X-Y scanner mmc2 on " + XyMmc.SerialPortName + ". " + ex.Message,
                        "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            canReconnectXyMmc = true;
        }

        // Reconnect z mmc command
        private object ZMmcLock = new object();
        public bool ZMmcReconnecting = false;
        public bool ZMmcConnected = false;
        private bool canReconnectZMmc = true;
        private bool CanExecuteReconnectZMmcMethod(object parameter) { return canReconnectZMmc; }
        private void ExecuteReconnectZMmcMethod(object parameter)
        {
            canReconnectZMmc = false;
            lock (ZMmcLock)
            {
                try
                {
                    ZMmcReconnecting = true;
                    ZMmcConnected = false;
                    PortResponse pr = ZMmc.Connect();
                    ZMmcConnected = true;
                    ZMmcReconnecting = false;
                }
                catch (Exception ex)
                {
                    PortResponse pr = new PortResponse("RE", ex.Message, ex);
                    ZMmcConnected = false;
                    ZMmcReconnecting = false;
                }
            }
            canReconnectZMmc = true;
        }
    }
}

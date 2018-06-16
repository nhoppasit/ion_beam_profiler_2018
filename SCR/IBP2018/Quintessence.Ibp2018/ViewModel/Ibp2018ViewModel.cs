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
                Properties.Settings.Default.Ammeter1GpibAddress = value; Properties.Settings.Default.Save();
                OnPropertyChanged("A1GpibAddress");
                OnPropertyChanged("A1VisaAddress");
            }
        }
        public string A1VisaAddressText
        {
            get { return _Ammeters[0].VisaAddress; }
            set
            {
                _Ammeters[0].VisaAddress = value;
                OnPropertyChanged("A1VisaAddressText");
            }
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
                Properties.Settings.Default.Ammeter1GpibAddress = value; Properties.Settings.Default.Save();
                OnPropertyChanged("A2GpibAddress");
                OnPropertyChanged("A2VisaAddress");
            }
        }
        public string A2VisaAddressText { get { return _Ammeters[1].VisaAddress; } }
        public double Current2 { get { return _Ammeters[1].Current; } set { _Ammeters[1].Current = value; OnPropertyChanged("Current2Text"); } }
        public string Current2Text { get { return _Ammeters[0].Current.ToString(); } set { _Ammeters[0].Current = Convert.ToSingle(value); OnPropertyChanged("Current1Text"); } }

        // Sensor interval
        private int _sensorInterval = 1;
        public int SensorInterval { get { return _sensorInterval; } set { _sensorInterval = value; OnPropertyChanged("SensorInterval"); } }

        /* ----------------------------------------------------------
         * X-Y scanner and Z axis object
         * ---------------------------------------------------------- */
        private MMC2Info _XyMmc;
        public MMC2Info XyMmc { get { return _XyMmc; } set { _XyMmc = value; } }
        public string XyMmcPortName { get { return _XyMmc.SerialPortName; } set { _XyMmc.SerialPortName = value; OnPropertyChanged("XyMmcPortName"); } }
        public string XLPosText { get { return _XyMmc.ActualX.ToString("F2"); } }
        public string YLPosText { get { return _XyMmc.ActualY.ToString("F2"); } }
        public string XScanStep
        {
            get { return _XyMmc.XScanStep.ToString(); }
            set
            {
                try { _XyMmc.XScanStep = Convert.ToDouble(value); }
                catch { _XyMmc.XScanStep = 0; }
                OnPropertyChanged("XScanStep");
            }
        }
        public string YScanStep
        {
            get { return _XyMmc.YScanStep.ToString(); }
            set
            {
                try { _XyMmc.YScanStep = Convert.ToDouble(value); }
                catch { _XyMmc.YScanStep = 0; }
                OnPropertyChanged("YScanStep");
            }
        }
        public string XScanStart
        {
            get { return _XyMmc.XScanStart.ToString(); }
            set
            {
                try { _XyMmc.XScanStart = Convert.ToDouble(value); }
                catch { _XyMmc.XScanStart = 0; }
                OnPropertyChanged("XScanStart");
            }
        }
        public string YScanStart
        {
            get { return _XyMmc.YScanStart.ToString(); }
            set
            {
                try { _XyMmc.YScanStart = Convert.ToDouble(value); }
                catch { _XyMmc.YScanStart = 0; }
                OnPropertyChanged("YScanStart");
            }
        }
        public string XScanEnd
        {
            get { return _XyMmc.XScanEnd.ToString(); }
            set
            {
                try { _XyMmc.XScanEnd = Convert.ToDouble(value); }
                catch { _XyMmc.XScanEnd = 0; }
                OnPropertyChanged("XScanEnd");
            }
        }
        public string YScanEnd
        {
            get { return _XyMmc.YScanEnd.ToString(); }
            set
            {
                try { _XyMmc.YScanEnd = Convert.ToDouble(value); }
                catch { _XyMmc.YScanEnd = 0; }
                OnPropertyChanged("YScanEnd");
            }
        }
        public string XFigtureMinimum
        {
            get { return _XyMmc.XFigtureMinimum.ToString(); }
            set
            {
                try { _XyMmc.XFigtureMinimum = Convert.ToDouble(value); }
                catch { _XyMmc.XFigtureMinimum = 0; }
                OnPropertyChanged("XFigtureMinimum");
            }
        }
        public string YFigtureMinimum
        {
            get { return _XyMmc.YFigtureMinimum.ToString(); }
            set
            {
                try { _XyMmc.YFigtureMinimum = Convert.ToDouble(value); }
                catch { _XyMmc.YFigtureMinimum = 0; }
                OnPropertyChanged("YFigtureMinimum");
            }
        }
        public string XFigtureMaximum
        {
            get { return _XyMmc.XFigtureMaximum.ToString(); }
            set
            {
                try { _XyMmc.XFigtureMaximum = Convert.ToDouble(value); }
                catch { _XyMmc.XFigtureMaximum = 0; }
                OnPropertyChanged("XFigtureMaximum");
            }
        }
        public string YFigtureMaximum
        {
            get { return _XyMmc.YFigtureMaximum.ToString(); }
            set
            {
                try { _XyMmc.YFigtureMaximum = Convert.ToDouble(value); }
                catch { _XyMmc.YFigtureMaximum = 0; }
                OnPropertyChanged("XFigtureMaximum");
            }
        }
        private MMC2Info _ZMmc;
        public MMC2Info ZMmc { get { return _ZMmc; } set { _ZMmc = value; } }
        public string ZMmcPortName { get { return _ZMmc.SerialPortName; } set { _ZMmc.SerialPortName = value; OnPropertyChanged("ZMmcPortName"); } }
        public string ZLPosText { get { return _ZMmc.ActualX.ToString("F2"); } }
        public string ZFigtureMinimum
        {
            get { return _ZMmc.XFigtureMinimum.ToString(); }
            set
            {
                try { _ZMmc.XFigtureMinimum = Convert.ToDouble(value); }
                catch { _ZMmc.XFigtureMinimum = 0; }
                OnPropertyChanged("ZFigtureMinimum");
            }
        }
        public string ZFigtureMaximum
        {
            get { return _ZMmc.XFigtureMaximum.ToString(); }
            set
            {
                try { _ZMmc.XFigtureMaximum = Convert.ToDouble(value); }
                catch { _ZMmc.XFigtureMaximum = 0; }
                OnPropertyChanged("ZFigtureMaximum");
            }
        }

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
        public ObservableCollection<DataGridColumn> Current1ColumnCollection
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

        #region Commands declaration
        // Commands for binding to buttons
        public ICommand InitializeMeterCommand { get; set; }
        public ICommand ConfigureMeterCommand { get; set; }
        public ICommand MeasureCurrentCommand { get; set; }
        public ICommand GenerateNewDemoDataCommand { get; set; }

        // Commands of xy and z mmc reconnect
        public ICommand ReconnectXyMmcCommand { get; set; }
        public ICommand ReconnectZMmcCommand { get; set; }
        public ICommand QueryScannerCommand { get; set; }
        public ICommand UnqueryScannerCommand { get; set; }

        // New measurement
        public ICommand NewMeasurementCommand { get; set; }

        // Re-connect meters
        public ICommand ReconnectMeter1Command { get; set; }
        #endregion

        // --------------------------------------- CONSTRUCTOR ------------------------------------------------
        public Ibp2018ViewModel()
        {
            // Create meters object
            _Ammeters = new List<Gpib34401aInfo>();

            // Meter 1
            Gpib34401aInfo a1 = new Gpib34401aInfo();
            a1.GpibBoardNumber = 0;
            a1.GpibAddress = 26;
            _Ammeters.Add(a1);

            // Meter 2
            Gpib34401aInfo a2 = new Gpib34401aInfo();
            a2.GpibBoardNumber = 0;
            a2.GpibAddress = 27;
            _Ammeters.Add(a2);

            // Create XY scanner object
            _XyMmc = new MMC2Info();
            _XyMmc.SerialPortName = "COM1";
            _XyMmc.XFigtureMinimum = -10;
            _XyMmc.XFigtureMaximum = 10;
            _XyMmc.YFigtureMinimum = -10;
            _XyMmc.YFigtureMaximum = 10;

            // Create Z axis object
            _ZMmc = new MMC2Info();
            _ZMmc.SerialPortName = "COM2";
            _ZMmc.XFigtureMinimum = -10;
            _ZMmc.XFigtureMaximum = 10;
            _ZMmc.YFigtureMinimum = -10;
            _ZMmc.YFigtureMaximum = 10;

            // Create data tables object
            _CurrentTables = new List<Ibp2018DataTableModel>();
            _CurrentTables.Add(new Ibp2018DataTableModel());
            _CurrentTables.Add(new Ibp2018DataTableModel());

            // Create commands
            InitializeMeterCommand = new RelayCommand(ExecuteInitializeMeterMethod, CanExecuteInitializeMeterMethod);
            ConfigureMeterCommand = new RelayCommand(ExecuteConfigureMeterMethod, CanExecuteConfigureMeterMethod);
            MeasureCurrentCommand = new RelayCommand(ExecuteMeasureMethod, CanExecuteMeasureMethod);
            GenerateNewDemoDataCommand = new RelayCommand(ExecuteGenerateDemoDataMethod, CanExecuteGenerateDemoDataMethod);

            // Serial port of xy and z mmc
            ReconnectXyMmcCommand = new RelayCommand(ExecuteReconnectXyMmcMethod, CanExecuteReconnectXyMmcMethod);
            ReconnectZMmcCommand = new RelayCommand(ExecuteReconnectZMmcMethod, CanExecuteReconnectZMmcMethod);

            // Query and unquery scanner 
            QueryScannerCommand = new RelayCommand(ExecuteQueryScannerMethod, CanExecuteQueryScannerMethod);
            UnqueryScannerCommand = new RelayCommand(ExecuteUnqueryScannerMethod, CanExecuteUnqueryScannerMethod);

            // New measurement
            NewMeasurementCommand = new RelayCommand(ExecuteNewMeasurementMethod, CanExecuteNewMeasurementMethod);

            #region Reconnect meters
            // Reconnect meter 1
            ReconnectMeter1Command = new RelayCommand(ExecuteReconnectMeter1Method, CanExecuteReconnectMeter1Method);
            #endregion

            // Reload settings
            ReloadSettings();
        }

        // Reload settings
        void ReloadSettings()
        {
            A1GpibAddress = Properties.Settings.Default.Ammeter1GpibAddress;
            A2GpibAddress = Properties.Settings.Default.Ammeter2GpibAddress;
        }

        // Reconnect meter 1
        private object Meter1Lock = new object();
        public bool Meter1Reconnecting = false;
        public bool Meter1Connected = false;
        private bool canReconnectMeter1 = true;
        private bool CanExecuteReconnectMeter1Method(object parameter) { return canReconnectMeter1; }
        private void ExecuteReconnectMeter1Method(object parameter)
        {
            canReconnectMeter1 = false;
            lock (Meter1Lock)
            {
                try
                {
                    Meter1Reconnecting = true;
                    Meter1Connected = false;

                    GpibResponse gr = _Ammeters[0].InitializeMeterForCurrent();
                    gr = _Ammeters[0].ConfigureMeterForCurrent();
                    gr = _Ammeters[0].MeasureCurrent();

                    if (gr.Code == GpibResponse.SUCCESS)
                    {
                        Meter1Connected = true;
                        Meter1Reconnecting = false;
                        MessageBox.Show("Meter 1 on " + _Ammeters[0].VisaAddress + " reconnected.", "Connect meters", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        Meter1Connected = true;
                        Meter1Reconnecting = false;
                        MessageBox.Show("Cannot reconnect meter 1 on " + _Ammeters[0].VisaAddress + ". " + gr.Message,
                        "Connect meters", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }
                catch (SystemException ex)
                {
                    GpibResponse gr = new GpibResponse("RE", ex.Message, ex);
                    Meter1Connected = false;
                    Meter1Reconnecting = false;
                    MessageBox.Show("Cannot reconnect meter 1 on " + _Ammeters[0].VisaAddress + ". " + ex.Message,
                        "Connect meters", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            canReconnectMeter1 = true;
        }


        // -------------------------------------------------------------------------------
        // CODE OF COMMNADS
        // -------------------------------------------------------------------------------
        // New measurement command by button
        private void NewMeasurement()
        {
            // Wait dialog can show in UI
            // AWait for ColumnsGenerating == false
            canNewMeasurement = false;
            ColumnsGenerating = true;

            // Current-1 table columns definetion
            _CurrentTables[0].GenerateNewDemoColumns(_XyMmc.XScanStep, _XyMmc.YScanStep, _XyMmc.XScanStart, _XyMmc.XScanEnd, _XyMmc.YScanStart, _XyMmc.YScanEnd);

            // Current-1 table columns definetion

            // Binding columns name and header
            Current1ColumnCollection.Clear();
            for (int i = 0; i < _CurrentTables[0].ColumnNames.Count; i++)
            {
                Binding binding = new Binding(_CurrentTables[0].ColumnNames[i]);
                DataGridTextColumn textColumn = new DataGridTextColumn();
                textColumn.Header = _CurrentTables[0].ColumnHeaders[i];
                textColumn.Binding = binding;
                Current1ColumnCollection.Add(textColumn);
            }

            ColumnsGenerating = false;
            canNewMeasurement = true;
        }
        private bool canNewMeasurement = true;
        public bool MustSavePreviousData = false;
        private bool CanExecuteNewMeasurementMethod(object parameter) { return canNewMeasurement && !MustSavePreviousData; }
        public void ExecuteNewMeasurementMethod(object parameter) { NewMeasurement(); }

        // Query and unquery scanner
        private bool canQueryScanner = true;
        private bool _continueQueryScanner = true;
        private void QueryScanner()
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
                            PortResponse pr = XyMmc.QueryPosition();
                            // Verify
                            if (pr.Code == PortResponse.SUCCESS)
                            {
                                OnPropertyChanged("XLPosText");
                                OnPropertyChanged("YLPosText");
                            }
                            else
                            {
                                _continueQueryScanner = false;
                                canQueryScanner = false;
                                MessageBox.Show(pr.Message, "Query X-Y Scanner", MessageBoxButton.OK, MessageBoxImage.Information);
                                break;
                            }
                            PortResponse prZ = ZMmc.QueryPosition();
                            // Verify
                            if (pr.Code == PortResponse.SUCCESS)
                            {
                                OnPropertyChanged("ZLPosText");
                            }
                            else
                            {
                                _continueQueryScanner = false;
                                canQueryScanner = false;
                                MessageBox.Show(pr.Message, "Query Z axis", MessageBoxButton.OK, MessageBoxImage.Information);
                                break;
                            }
                        }
                        catch (Exception ex) { }
                        Thread.Sleep(500);
                    }
                    //MessageBox.Show("Auto query is turned OFF.", "Query X-Y Scanner", MessageBoxButton.OK, MessageBoxImage.Information);
                });
        }
        private bool CanExecuteQueryScannerMethod(object parameter) { return canQueryScanner; }
        private void ExecuteQueryScannerMethod(object parameter) { QueryScanner(); }
        private bool CanExecuteUnqueryScannerMethod(object parameter) { return true; }
        private void ExecuteUnqueryScannerMethod(object parameter) { _continueQueryScanner = false; canQueryScanner = true; }

        // Initialize Meter
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
                            Thread.Sleep(1000);

                            //DataTable dt = CurrentTable1;
                            DataRow r = CurrentTable1.Rows[scanY_Idx];
                            r[scanX_Idx + 1] = Current1;

                        }
                        catch (Exception ex) { }
                    }
                });
        }

        // Configure meter
        private bool CanExecuteConfigureMeterMethod(object parameter) { return canPause; }
        private void ExecuteConfigureMeterMethod(object parameter)
        {
            _continueInitMeter = false;
            canMeasure = true;
            canPause = false;
            canStop = true;
        }

        // Measurement
        private bool CanExecuteMeasureMethod(object parameter) { return canStop; }
        private void ExecuteMeasureMethod(object parameter)
        {
            _continueInitMeter = false;
            MessageBox.Show("Measured");
            canMeasure = true;
            canPause = false;
            canStop = false;
        }

        // DEMO DATA
        private void GenerateDemoData()
        {
            // Wait dialog can show in view 
            // Wait for columns.count

            canDemo = false;
            ColumnsGenerating = true;

            // Current-1 data table
            _CurrentTables[0].GenerateNewDemoColumns(_XyMmc.XScanStep, _XyMmc.YScanStep, _XyMmc.XScanStart, _XyMmc.XScanEnd, _XyMmc.YScanStart, _XyMmc.YScanEnd);

            // Binding columns name and header
            Current1ColumnCollection.Clear();
            for (int i = 0; i < _CurrentTables[0].ColumnNames.Count; i++)
            {
                Binding binding = new Binding(_CurrentTables[0].ColumnNames[i]);
                DataGridTextColumn textColumn = new DataGridTextColumn();
                textColumn.Header = _CurrentTables[0].ColumnHeaders[i];
                textColumn.Binding = binding;
                Current1ColumnCollection.Add(textColumn);
            }



            ColumnsGenerating = false;
            canDemo = true;
        }
        private bool canDemo = true;
        public bool ColumnsGenerating = false;
        private bool CanExecuteGenerateDemoDataMethod(object parameter) { return canDemo; }
        public void ExecuteGenerateDemoDataMethod(object parameter) { GenerateDemoData(); }

        #region Connect scanner
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
                    if (pr.Code == PortResponse.SUCCESS)
                    {
                        ZMmcConnected = true;
                        ZMmcReconnecting = false;
                        MessageBox.Show("X-Y scanner mmc2 on " + ZMmc.SerialPortName + " reconnected.", "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        ZMmcConnected = true;
                        ZMmcReconnecting = false;
                        MessageBox.Show("Cannot reconnect X-Y scanner mmc2 on " + ZMmc.SerialPortName + ". " + pr.Message,
                        "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }
                catch (Exception ex)
                {
                    PortResponse pr = new PortResponse("RE", ex.Message, ex);
                    ZMmcConnected = false;
                    ZMmcReconnecting = false;
                    MessageBox.Show("Cannot reconnect X-Y scanner mmc2 on " + ZMmc.SerialPortName + ". " + ex.Message,
                        "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            canReconnectZMmc = true;
        }
        #endregion

        #region Jog method
        // Jogging method        
        public void XJog(int step)
        {
            _XyMmc.JogX(step);
            OnPropertyChanged("XLPosText");
        }
        #endregion
    }
}

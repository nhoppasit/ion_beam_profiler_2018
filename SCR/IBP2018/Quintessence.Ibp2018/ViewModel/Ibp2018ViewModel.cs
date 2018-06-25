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
        /* ----------------------------------------------------------
         * Ammeter-1 Properties
         * ---------------------------------------------------------- */
        private Gpib34401aInfo _Ammeter1;
        public Gpib34401aInfo Ammeter1 { get { return _Ammeter1; } set { _Ammeter1 = value; } }
        public int A1GpibAddress
        {
            get { return _Ammeter1.GpibAddress; }
            set
            {
                _Ammeter1.GpibAddress = value;
                Properties.Settings.Default.Ammeter1GpibAddress = value; Properties.Settings.Default.Save();
                OnPropertyChanged("A1GpibAddress");
                OnPropertyChanged("A1VisaAddress");
            }
        }
        public string A1VisaAddressText
        {
            get { return _Ammeter1.VisaAddress; }
            set
            {
                _Ammeter1.VisaAddress = value;
                OnPropertyChanged("A1VisaAddressText");
            }
        }
        public double Current1 { get { return _Ammeter1.Current; } set { _Ammeter1.Current = value; OnPropertyChanged("Current1TextuA"); } }
        public string Current1TextuA { get { return (_Ammeter1.Current * 1e6).ToString(); } set { _Ammeter1.Current = Convert.ToDouble(value) / 1e6; OnPropertyChanged("Current1TextuA"); } }

        /* ----------------------------------------------------------
         * Ammeter-2 Properties
         * ---------------------------------------------------------- */
        private Gpib34401aInfo _Ammeter2;
        public Gpib34401aInfo Ammeter2 { get { return _Ammeter2; } set { _Ammeter2 = value; } }
        public int A2GpibAddress
        {
            get { return _Ammeter2.GpibAddress; }
            set
            {
                _Ammeter2.GpibAddress = value;
                Properties.Settings.Default.Ammeter2GpibAddress = value; Properties.Settings.Default.Save();
                OnPropertyChanged("A2GpibAddress");
                OnPropertyChanged("A2VisaAddress");
            }
        }
        public string A2VisaAddressText { get { return _Ammeter2.VisaAddress; } }
        public double Current2 { get { return _Ammeter2.Current; } set { _Ammeter2.Current = value; OnPropertyChanged("Current2TextuA"); } }
        public string Current2TextuA { get { return (_Ammeter2.Current * 1e6).ToString(); } set { _Ammeter2.Current = Convert.ToDouble(value) / 1e6; OnPropertyChanged("Current2TextuA"); } }

        // Sensor interval
        private int _sensorInterval = 1;
        public int SensorInterval { get { return _sensorInterval; } set { _sensorInterval = value; OnPropertyChanged("SensorInterval"); } }

        // Meter demo mode
        private bool _IsMetersDemo = false;
        public bool IsMetersDemo { get { return _IsMetersDemo; } set { _IsMetersDemo = _Ammeter1.IsDemo = _Ammeter2.IsDemo = value; OnPropertyChanged("IsMetersDemo"); } }

        /* ----------------------------------------------------------
         * X-Y scanner and Z axis object
         * ---------------------------------------------------------- */
        private MMC2Info _XyMmc;
        public MMC2Info XyMmc { get { return _XyMmc; } set { _XyMmc = value; } }
        public string XyMmcPortName { get { return _XyMmc.SerialPortName; } set { _XyMmc.SerialPortName = value; OnPropertyChanged("XyMmcPortName"); } }
        public string XLPosText { get { return _XyMmc.ActualX.ToString("F2"); } set { _XyMmc.ActualX = Convert.ToDouble(value); OnPropertyChanged("XLPosText"); } }
        public string YLPosText { get { return _XyMmc.ActualY.ToString("F2"); } set { _XyMmc.ActualY = Convert.ToDouble(value); OnPropertyChanged("YLPosText"); } }
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
                OnPropertyChanged("XScanRangeList");
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
                OnPropertyChanged("YScanRangeList");
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
                OnPropertyChanged("XScanRangeList");
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
                OnPropertyChanged("YScanRangeList");
            }
        }
        public List<double> XScanRangeList { get { return _XyMmc.XScanRangeList; } }
        public List<double> YScanRangeList { get { return _XyMmc.YScanRangeList; } }
        private MMC2Info _ZMmc;
        public MMC2Info ZMmc { get { return _ZMmc; } set { _ZMmc = value; } }
        public string ZMmcPortName { get { return _ZMmc.SerialPortName; } set { _ZMmc.SerialPortName = value; OnPropertyChanged("ZMmcPortName"); } }
        public string ZLPosText { get { return _ZMmc.ActualX.ToString("F2"); } set { _ZMmc.ActualX = Convert.ToDouble(value); OnPropertyChanged("ZLPosText"); } }
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

        // Demo mode of scanner
        private bool _IsScannerDemo = false;
        public bool IsScannerDemo { get { return _IsScannerDemo; } set { _IsScannerDemo = _XyMmc.IsDemo = _ZMmc.IsDemo = value; OnPropertyChanged("IsScannerDemo"); } }

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

        // Commands of xy and z mmc reconnect

        // New measurement

        // Re-connect meters
        #endregion

        // --------------------------------------- CONSTRUCTOR ------------------------------------------------
        public Ibp2018ViewModel()
        {
            // Create meters object
            _Ammeter1 = new Gpib34401aInfo();
            _Ammeter2 = new Gpib34401aInfo();

            // Meter 1
            Gpib34401aInfo a1 = new Gpib34401aInfo
            {
                GpibBoardNumber = 0,
                GpibAddress = 26
            };

            // Meter 2
            Gpib34401aInfo a2 = new Gpib34401aInfo
            {
                GpibBoardNumber = 0,
                GpibAddress = 27
            };

            // Create XY scanner object
            _XyMmc = new MMC2Info
            {
                SerialPortName = "COM1",
                XFigtureMinimum = -10,
                XFigtureMaximum = 10,
                YFigtureMinimum = -10,
                YFigtureMaximum = 10
            };

            // Create Z axis object
            _ZMmc = new MMC2Info
            {
                SerialPortName = "COM2",
                XFigtureMinimum = -10,
                XFigtureMaximum = 10,
                YFigtureMinimum = -10,
                YFigtureMaximum = 10
            };

            // Create data tables object
            _CurrentTables = new List<Ibp2018DataTableModel>
            {
                new Ibp2018DataTableModel(),
                new Ibp2018DataTableModel()
            };

            // Create commands
            InitializeMeterCommand = new RelayCommand(ExecuteInitializeMeterMethod, CanExecuteInitializeMeterMethod);
            ConfigureMeterCommand = new RelayCommand(ExecuteConfigureMeterMethod, CanExecuteConfigureMeterMethod);
            RunCommand = new RelayCommand(ExecuteRunMethod, CanExecuteRunMethod);
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
            ReconnectMeter2Command = new RelayCommand(ExecuteReconnectMeter2Method, CanExecuteReconnectMeter2Method);
            #endregion

            // Reload settings
            ReloadSettings();
        }

        /// <summary>
        /// Reload settings from ROM
        /// </summary>
        void ReloadSettings()
        {
            A1GpibAddress = Properties.Settings.Default.Ammeter1GpibAddress;
            A2GpibAddress = Properties.Settings.Default.Ammeter2GpibAddress;
            XyMmcPortName = Properties.Settings.Default.XyMmcPortName;
            ZMmcPortName = Properties.Settings.Default.ZMmcPortName;
        }

        #region Command definetino

        // Reconnect meter 1
        public ICommand ReconnectMeter1Command { get; set; }
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

                    GpibResponse gr = _Ammeter1.InitializeMeterForCurrent();
                    gr = _Ammeter1.ConfigureMeterForCurrent();
                    gr = _Ammeter1.MeasureCurrent();

                    if (gr.Code == GpibResponse.SUCCESS)
                    {
                        Meter1Connected = true;
                        Meter1Reconnecting = false;
                        MessageBox.Show("Meter 1 on " + _Ammeter1.VisaAddress + " reconnected.", "Connect meters", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        Meter1Connected = true;
                        Meter1Reconnecting = false;
                        MessageBox.Show("Cannot reconnect meter 1 on " + _Ammeter1.VisaAddress + ". " + gr.Message,
                            "Connect meters", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }
                catch (SystemException ex)
                {
                    GpibResponse gr = new GpibResponse("RE", ex.Message, ex);
                    Meter1Connected = false;
                    Meter1Reconnecting = false;
                    MessageBox.Show("Cannot reconnect meter 1 on " + _Ammeter1.VisaAddress + ". " + ex.Message,
                        "Connect meters", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            canReconnectMeter1 = true;
        }

        // Reconnect meter 2
        public ICommand ReconnectMeter2Command { get; set; }
        private object Meter2Lock = new object();
        public bool Meter2Reconnecting = false;
        public bool Meter2Connected = false;
        private bool canReconnectMeter2 = true;
        private bool CanExecuteReconnectMeter2Method(object parameter) { return canReconnectMeter2; }
        private void ExecuteReconnectMeter2Method(object parameter)
        {
            canReconnectMeter2 = false;
            lock (Meter2Lock)
            {
                try
                {
                    Meter2Reconnecting = true;
                    Meter2Connected = false;

                    GpibResponse gr = _Ammeter2.InitializeMeterForCurrent();
                    gr = _Ammeter2.ConfigureMeterForCurrent();
                    gr = _Ammeter2.MeasureCurrent();

                    if (gr.Code == GpibResponse.SUCCESS)
                    {
                        Meter2Connected = true;
                        Meter2Reconnecting = false;
                        MessageBox.Show("Meter 2 on " + _Ammeter2.VisaAddress + " reconnected.", "Connect meters", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        Meter2Connected = true;
                        Meter2Reconnecting = false;
                        MessageBox.Show("Cannot reconnect meter 2 on " + _Ammeter2.VisaAddress + ". " + gr.Message,
                            "Connect meters", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }
                catch (SystemException ex)
                {
                    GpibResponse gr = new GpibResponse("RE", ex.Message, ex);
                    Meter2Connected = false;
                    Meter2Reconnecting = false;
                    MessageBox.Show("Cannot reconnect meter 1 on " + _Ammeter2.VisaAddress + ". " + ex.Message,
                        "Connect meters", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            canReconnectMeter2 = true;
        }

        /// <summary>
        /// Read Current
        /// </summary>
        /// <param name="parameter"></param>
        /// <returns></returns>
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

                    GpibResponse gr = _Ammeter1.InitializeMeterForCurrent();
                    gr = _Ammeter1.ConfigureMeterForCurrent();
                    gr = _Ammeter1.MeasureCurrent();

                    if (gr.Code == GpibResponse.SUCCESS)
                    {
                        Meter1Connected = true;
                        Meter1Reconnecting = false;
                        MessageBox.Show("Meter 1 on " + _Ammeter1.VisaAddress + " reconnected.", "Connect meters", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        Meter1Connected = true;
                        Meter1Reconnecting = false;
                        MessageBox.Show("Cannot reconnect meter 1 on " + _Ammeter1.VisaAddress + ". " + gr.Message,
                            "Connect meters", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }
                catch (SystemException ex)
                {
                    GpibResponse gr = new GpibResponse("RE", ex.Message, ex);
                    Meter1Connected = false;
                    Meter1Reconnecting = false;
                    MessageBox.Show("Cannot reconnect meter 1 on " + _Ammeter1.VisaAddress + ". " + ex.Message,
                        "Connect meters", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            canReconnectMeter1 = true;
        }

        // New measurement command by button
        public ICommand NewMeasurementCommand { get; set; }
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
                DataGridTextColumn textColumn = new DataGridTextColumn
                {
                    Header = _CurrentTables[0].ColumnHeaders[i],
                    Binding = binding
                };
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
        public ICommand QueryScannerCommand { get; set; }
        public ICommand UnqueryScannerCommand { get; set; }
        private bool canQueryScanner = true;
        private void QueryScanner()
        {
            canQueryScanner = false;
            try
            {
                PortResponse pr1 = XyMmc.QueryPosition();
                // Verify
                if (pr1.Code == PortResponse.SUCCESS)
                {
                    OnPropertyChanged("XLPosText");
                    OnPropertyChanged("YLPosText");
                }
                else
                {
                    MessageBox.Show(pr1.Message, "Query X-Y Scanner", MessageBoxButton.OK, MessageBoxImage.Information);
                    if (MessageBox.Show("Do you want to re-connect X-Y scanner MMC2 driver on " + XyMmcPortName + "?", "Re-connect X-Y scanner MMC2 driver", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        ReconnectXyMmcCommand.Execute(new object());
                    }
                    else
                    {
                        // Do nothing
                    }
                    canQueryScanner = true;
                    return;
                }
                PortResponse pr2 = ZMmc.QueryPosition();
                // Verify
                if (pr2.Code == PortResponse.SUCCESS)
                {
                    OnPropertyChanged("ZLPosText");
                }
                else
                {
                    MessageBox.Show(pr2.Message, "Query Z axis", MessageBoxButton.OK, MessageBoxImage.Information);
                    if (MessageBox.Show("Do you want to re-connect Z axis MMC2 driver on " + ZMmcPortName + "?", "Re-connect Z axis MMC2 driver", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        ReconnectZMmcCommand.Execute(new object());
                    }
                    else
                    {
                        // Do nothing
                    }
                    canQueryScanner = true;
                    return;
                }
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse(PortResponse.ERR_QUERY, ex.Message, ex);
                MessageBox.Show("Query MMC2 error! " + ex.Message, "Query X-Y Scanner", MessageBoxButton.OK, MessageBoxImage.Information);
                canQueryScanner = true;
                return;
            }
        }
        private bool CanExecuteQueryScannerMethod(object parameter) { return canQueryScanner; }
        private void ExecuteQueryScannerMethod(object parameter) { QueryScanner(); }
        private bool CanExecuteUnqueryScannerMethod(object parameter) { return true; }
        private void ExecuteUnqueryScannerMethod(object parameter) { canQueryScanner = true; }

        /// <summary>
        /// Initialize dmm 34401a command
        /// </summary>
        public ICommand InitializeMeterCommand { get; set; }
        private bool canMeasure = true, canPause = false, canStop = false;
        private bool _continueInitMeter = true;
        private bool CanExecuteInitializeMeterMethod(object parameter) { return canMeasure; }
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
                            //DataRow r = CurrentTable1.Rows[scanY_Idx];
                            //r[scanX_Idx + 1] = Current1;

                        }
                        catch (Exception ex) { }
                    }
                });
        }

        // Configure meter
        public ICommand ConfigureMeterCommand { get; set; }
        private bool CanExecuteConfigureMeterMethod(object parameter) { return canPause; }
        private void ExecuteConfigureMeterMethod(object parameter)
        {
            _continueInitMeter = false;
            canMeasure = true;
            canPause = false;
            canStop = true;
        }

        // Measurement
        public ICommand RunCommand { get; set; }
        private bool CanExecuteRunMethod(object parameter) { return canStop; }
        private void ExecuteRunMethod(object parameter)
        {
            _continueInitMeter = false;
            MessageBox.Show("Measured");
            canMeasure = true;
            canPause = false;
            canStop = false;
        }

        // DEMO DATA
        public ICommand GenerateNewDemoDataCommand { get; set; }
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

        // Reconnect xy mmc command
        public ICommand ReconnectXyMmcCommand { get; set; }
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
        public ICommand ReconnectZMmcCommand { get; set; }
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

        // X goto zero command
        public ICommand GotoXZeroCommand { get; set; }
        private bool canGotoXZeroCommand = false;
        private bool CanExecuteGotoXZeroMethod(object param) { return canGotoXZeroCommand; }
        private void ExecuteGotoXZeroMethod(object param)
        {
            canGotoXZeroCommand = false;
            try
            {
                PortResponse pr = XyMmc.GotoZeroX();
                if (pr.Code == PortResponse.SUCCESS)
                {
                    MessageBox.Show("Done X axis goto zero on " + XyMmc.SerialPortName + ".", "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("Error when X axis goto to zero on " + XyMmc.SerialPortName + ". " + pr.Message,
                    "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse(PortResponse.ERR_ZERO, ex.Message, ex);
                MessageBox.Show("Error when X axis goto to zero on " + XyMmc.SerialPortName + ". " + ex.Message,
                    "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            canGotoXZeroCommand = true;
        }

        // Y goto zero command
        public ICommand GotoYZeroCommand { get; set; }
        private bool canGotoYZeroCommand = false;
        private bool CanExecuteGotoYZeroMethod(object param) { return canGotoYZeroCommand; }
        private void ExecuteGotoYZeroMethod(object param)
        {
            canGotoYZeroCommand = false;
            try
            {
                PortResponse pr = XyMmc.GotoZeroY();
                if (pr.Code == PortResponse.SUCCESS)
                {
                    MessageBox.Show("Done Y axis goto zero on " + XyMmc.SerialPortName + ".", "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("Error when Y axis goto to zero on " + XyMmc.SerialPortName + ". " + pr.Message,
                    "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse(PortResponse.ERR_ZERO, ex.Message, ex);
                MessageBox.Show("Error when Y axis goto to zero on " + XyMmc.SerialPortName + ". " + ex.Message,
                    "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            canGotoYZeroCommand = true;
        }

        // Z goto zero command
        public ICommand GotoZZeroCommand { get; set; }
        private bool canGotoZZeroCommand = false;
        private bool CanExecuteGotoZZeroMethod(object param) { return canGotoZZeroCommand; }
        private void ExecuteGotoZZeroMethod(object param)
        {
            canGotoZZeroCommand = false;
            try
            {
                PortResponse pr = ZMmc.GotoZeroX();
                if (pr.Code == PortResponse.SUCCESS)
                {
                    MessageBox.Show("Done Z axis goto zero on " + ZMmc.SerialPortName + ".", "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("Error when Z axis goto to zero on " + ZMmc.SerialPortName + ". " + pr.Message,
                    "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse(PortResponse.ERR_ZERO, ex.Message, ex);
                MessageBox.Show("Error when Z axis goto to zero on " + ZMmc.SerialPortName + ". " + ex.Message,
                    "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            canGotoZZeroCommand = true;
        }

        // X set zero command
        public ICommand SetXZeroCommand { get; set; }
        private bool canSetXZeroCommand = false;
        private bool CanExecuteSetXZeroMethod(object param) { return canSetXZeroCommand; }
        private void ExecuteSetXZeroMethod(object param)
        {
            canSetXZeroCommand = false;
            try
            {
                PortResponse pr = XyMmc.SetZeroX();
                if (pr.Code == PortResponse.SUCCESS)
                {
                    MessageBox.Show("Done set X axis as zero on " + XyMmc.SerialPortName + ".", "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("Error when set X axis as zero on " + XyMmc.SerialPortName + ". " + pr.Message,
                    "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse(PortResponse.ERR_ZERO, ex.Message, ex);
                MessageBox.Show("Error when set X axis as zero on " + XyMmc.SerialPortName + ". " + ex.Message,
                    "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            canSetXZeroCommand = true;
        }

        // Y set zero command
        public ICommand SetYZeroCommand { get; set; }
        private bool canSetYZeroCommand = false;
        private bool CanExecuteSetYZeroMethod(object param) { return canSetYZeroCommand; }
        private void ExecuteSetYZeroMethod(object param)
        {
            canSetYZeroCommand = false;
            try
            {
                PortResponse pr = XyMmc.SetZeroY();
                if (pr.Code == PortResponse.SUCCESS)
                {
                    MessageBox.Show("Done set Y axis as zero on " + XyMmc.SerialPortName + ".", "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("Error when set Y axis as zero on " + XyMmc.SerialPortName + ". " + pr.Message,
                    "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse(PortResponse.ERR_ZERO, ex.Message, ex);
                MessageBox.Show("Error when set Y axis as zero on " + XyMmc.SerialPortName + ". " + ex.Message,
                    "X-Y scanner", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            canSetYZeroCommand = true;
        }

        // Z set zero command
        public ICommand SetZZeroCommand { get; set; }
        private bool canSetZZeroCommand = false;
        private bool CanExecuteSetZZeroMethod(object param) { return canSetZZeroCommand; }
        private void ExecuteSetZZeroMethod(object param)
        {
            canSetZZeroCommand = false;
            try
            {
                PortResponse pr = ZMmc.SetZeroX();
                if (pr.Code == PortResponse.SUCCESS)
                {
                    MessageBox.Show("Done set Z axis as zero on " + ZMmc.SerialPortName + ".", "Z axis", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("Error when set Y axis as zero on " + ZMmc.SerialPortName + ". " + pr.Message,
                    "Z axis", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse(PortResponse.ERR_ZERO, ex.Message, ex);
                MessageBox.Show("Error when set Y axis as zero on " + ZMmc.SerialPortName + ". " + ex.Message,
                    "Z axis", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            canSetZZeroCommand = true;
        }

        #endregion

        #region Jog method
        // X Jogging method        
        public PortResponse XJog(int step)
        {
            PortResponse pr = _XyMmc.JogX(step);
            OnPropertyChanged("XLPosText");
            return pr;
        }
        // Y Jogging method        
        public PortResponse YJog(int step)
        {
            PortResponse pr = _XyMmc.JogY(step);
            OnPropertyChanged("YLPosText");
            return pr;
        }
        // Y Jogging method        
        public PortResponse ZJog(int step)
        {
            PortResponse pr = _ZMmc.JogX(step);
            //_ZMmc.JogY(step);
            OnPropertyChanged("ZLPosText");
            return pr;
        }
        #endregion
    }
}

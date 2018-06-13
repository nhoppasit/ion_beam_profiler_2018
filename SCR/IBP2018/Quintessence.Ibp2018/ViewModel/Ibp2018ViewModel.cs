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
        private IList<MMC2Info> _MMC2s;
        public IList<MMC2Info> MMC2s { get { return _MMC2s; } set { _MMC2s = value; } }
        private IList<Ibp2018DataTableModel> _CurrentTables;
        public IList<Ibp2018DataTableModel> CurrentTables { get { return _CurrentTables; } set { _CurrentTables = value; } }

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
        private bool isBusyAmmeter2 = false;


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
         * Commands for binding to buttons
         * ---------------------------------------------------------- */
        public ICommand InitializeMeter { get; set; }
        public ICommand ConfigureMeter { get; set; }
        public ICommand Measure { get; set; }
        public ICommand GenerateNewDemoData { get; set; }

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
            bool _continueProcessing=true;
            ThreadPool.QueueUserWorkItem(
                o =>
                {
                    while (_continueProcessing)
                    {
                        Current1 += 1;
                        Thread.Sleep(200);
                    }
                });

            
            //canDemo = false;
            // Current-1 data table
            _CurrentTables[0].GenerateNewDemoColumns(0.1, 0.1, 0, 20, 0, 20, true);

            // Binding columns name and header
            for (int i = 0; i < _CurrentTables[0].ColumnNames.Count; i++)
            {
                Binding binding = new Binding(_CurrentTables[0].ColumnNames[i]);
                DataGridTextColumn textColumn = new DataGridTextColumn();
                textColumn.Header = _CurrentTables[0].ColumnHeaders[i];
                textColumn.Binding = binding;
                ColumnCollection.Add(textColumn);
            }
            _continueProcessing = false;
            canDemo = true;
        }
        private bool canDemo = true;
        private bool CanExecuteGenerateDemoDataMethod(object parameter) { return canDemo; }
        public void ExecuteGenerateDemoDataMethod(object parameter) { GenerateDemoData(); }
    }
}

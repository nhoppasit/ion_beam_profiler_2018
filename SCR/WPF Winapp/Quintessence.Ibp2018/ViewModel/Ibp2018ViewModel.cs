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
        private bool isBusyAmmeter1 = false;

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
        private DataTable _datatable = new DataTable();
        public DataTable Datatable
        {
            get
            {
                return _datatable;
            }
            set
            {
                if (_datatable != value)
                {
                    _datatable = value;
                }
                OnPropertyChanged("Datatable");
            }
        }

        /* ----------------------------------------------------------
         * Commands for binding to buttons
         * ---------------------------------------------------------- */
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
            Current2Text = "200.02";

            InitializeMeter = new DelegateCommand(ExecuteInitializeMeterMethod, CanExecuteInitializeMeterMethod);
            ConfigureMeter = new RelayCommand(ExecuteConfigureMeterMethod, CanExecuteConfigureMeterMethod);
            Measure = new RelayCommand(ExecuteMeasureMethod, CanExecuteMeasureMethod);

            Datatable.Columns.Add("Name", typeof(string));
            Datatable.Columns.Add("Color", typeof(string));
            Datatable.Columns.Add("Phone", typeof(string));

            Datatable.Rows.Add("Vinoth", "#00FF00", "456345654");
            Datatable.Rows.Add("lkjasdgl", "Blue", "45654");
            Datatable.Rows.Add("Vinoth", "#FF0000", "456456");

            System.Windows.Data.Binding bindings = new System.Windows.Data.Binding("Name");
            System.Windows.Data.Binding bindings1 = new System.Windows.Data.Binding("Phone");
            System.Windows.Data.Binding bindings2 = new System.Windows.Data.Binding("Color");
            DataGridTextColumn s = new DataGridTextColumn();
            s.Header = "Name";
            s.Binding = bindings;
            DataGridTextColumn s1 = new DataGridTextColumn();
            s1.Header = "Phone";
            s1.Binding = bindings1;
            DataGridTextColumn s2 = new DataGridTextColumn();
            s2.Header = "Color";
            s2.Binding = bindings2;

            FrameworkElementFactory textblock = new FrameworkElementFactory(typeof(TextBlock));
            textblock.Name = "text";
            System.Windows.Data.Binding prodID = new System.Windows.Data.Binding("Name");
            System.Windows.Data.Binding color = new System.Windows.Data.Binding("Color");
            textblock.SetBinding(TextBlock.TextProperty, prodID);
            textblock.SetValue(TextBlock.TextWrappingProperty, TextWrapping.Wrap);
            //textblock.SetValue(TextBlock.BackgroundProperty, color);
            textblock.SetValue(TextBlock.NameProperty, "textblock");
            //FrameworkElementFactory border = new FrameworkElementFactory(typeof(Border));
            //border.SetValue(Border.NameProperty, "border");
            //border.AppendChild(textblock);
            //DataTrigger t = new DataTrigger();
            //t.Binding = new System.Windows.Data.Binding { Path = new PropertyPath("Name"), Converter = new EnableConverter(), ConverterParameter = "Phone" };
            //t.Value = 1;
            //t.Setters.Add(new Setter(TextBlock.BackgroundProperty, Brushes.LightGreen, textblock.Name));
            //t.Setters.Add(new Setter(TextBlock.ToolTipProperty, bindings, textblock.Name));
            //DataTrigger t1 = new DataTrigger();
            //t1.Binding = new System.Windows.Data.Binding { Path = new PropertyPath("Name"), Converter = new EnableConverter(), ConverterParameter = "Phone" };
            //t1.Value = 2;
            //t1.Setters.Add(new Setter(TextBlock.BackgroundProperty, Brushes.LightYellow, textblock.Name));
            //t1.Setters.Add(new Setter(TextBlock.ToolTipProperty, bindings, textblock.Name));

            DataTemplate d = new DataTemplate();
            d.VisualTree = textblock;
            //d.Triggers.Add(t);
            //d.Triggers.Add(t1);

            DataGridTemplateColumn s3 = new DataGridTemplateColumn();
            s3.Header = "Name 1";
            s3.CellTemplate = d;
            s3.Width = 140;

            ColumnCollection.Add(s);
            ColumnCollection.Add(s1);
            ColumnCollection.Add(s2);
            ColumnCollection.Add(s3);
        }

        private bool CanExecuteInitializeMeterMethod(object parameter)
        {
            return !isBusyAmmeter1;
        }

        private void ExecuteInitializeMeterMethod(object parameter)
        {
            ThreadPool.QueueUserWorkItem(
                o =>
                {
                    while (true)
                    {
                        isBusyAmmeter1 = true;
                        Current1 += 1;
                        //isBusyAmmeter1 = false;
                        System.Threading.Thread.Sleep(1000);
                        GpibResponse gr = _Ammeters[0].CreateIO488Object();
                        //isBusyAmmeter1 = true;
                        Current1 += 1;
                        isBusyAmmeter1 = false;
                    }
                });
        }

        private bool CanExecuteConfigureMeterMethod(object parameter) { return !isBusyAmmeter1; }

        private void ExecuteConfigureMeterMethod(object parameter)
        {
            MessageBox.Show("Configured");
        }

        private bool CanExecuteMeasureMethod(object parameter) { return !isBusyAmmeter1; }

        private void ExecuteMeasureMethod(object parameter)
        {
            Current1 += 1;
            //Current1Text = "111.11";
            //Current2 = "222.22";
            //MessageBox.Show("Measured");
        }

    }
}

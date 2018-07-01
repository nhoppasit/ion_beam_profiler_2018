using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Ribbon;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using System.Threading;
using System.ComponentModel;
using Quintessence.Ibp2018.ViewModel;
using Quintessence.MotionControl;
using Quintessence.MotionControl.MMC2;
using System.Reflection;
using Quintessence.Ibp2018.Model;
using IBP2018.View;
using System.Collections.ObjectModel;

namespace IBP2018
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : RibbonWindow
    {
        // ------------------------------- CONSTRUCTOR ---------------------------------
        public MainWindow()
        {
            // Initialize
            InitializeComponent();

            // Combobox
            InitializeRibbonComboboxMember();

            // Version
            this.Title = "Ion Beam Profiler Version: " + Assembly.GetExecutingAssembly().GetName().Version.ToString();

            #region JOGGING

            #region X
            // X jog background worker
            bwXJog.WorkerReportsProgress = true;
            bwXJog.DoWork += BwXJog_DoWork;
            bwXJog.RunWorkerCompleted += BwXJog_RunWorkerCompleted;

            // X negative jog
            mnuXJogNegative.PreviewMouseDown += MnuXJogNegative_PreviewMouseDown;
            mnuXJogNegative.PreviewMouseUp += MnuXJogNegative_PreviewMouseUp;

            // X positive jog
            mnuXJogPositive.PreviewMouseDown += MnuXJogPositive_PreviewMouseDown;
            mnuXJogPositive.PreviewMouseUp += MnuXJogPositive_PreviewMouseUp;
            #endregion

            #region Y
            // Y Jog background worker
            bwYJog.WorkerReportsProgress = true;
            bwYJog.DoWork += BwYJog_DoWork;
            bwYJog.RunWorkerCompleted += BwYJog_RunWorkerCompleted;

            // Y negative jog
            mnuYJogNegative.PreviewMouseDown += MnuYJogNegative_PreviewMouseDown;
            mnuYJogNegative.PreviewMouseUp += MnuYJogNegative_PreviewMouseUp;

            // Y positive jog
            mnuYJogPositive.PreviewMouseDown += MnuYJogPositive_PreviewMouseDown;
            mnuYJogPositive.PreviewMouseUp += MnuYJogPositive_PreviewMouseUp;

            #endregion

            #region Z

            // Z Jog background worker
            bwZJog.WorkerReportsProgress = true;
            bwZJog.DoWork += BwZJog_DoWork;
            bwZJog.RunWorkerCompleted += BwZJog_RunWorkerCompleted;

            // Y negative jog
            mnuZJogNegative.PreviewMouseDown += MnuZJogNegative_PreviewMouseDown;
            mnuZJogNegative.PreviewMouseUp += MnuZJogNegative_PreviewMouseUp;

            // Y positive jog
            mnuZJogPositive.PreviewMouseDown += MnuZJogPositive_PreviewMouseDown;
            mnuZJogPositive.PreviewMouseUp += MnuZJogPositive_PreviewMouseUp;

            #endregion

            #endregion

            #region New Measurement
            bwNewMeasurement.WorkerReportsProgress = true;
            bwNewMeasurement.DoWork += BwNewMeasurement_DoWork;
            bwNewMeasurement.RunWorkerCompleted += BwNewMeasurement_RunWorkerCompleted;
            mnuNew.Click += MnuNew_Click;
            #endregion

            #region Exit application
            mnuExit.Click += MnuExit_Click;
            this.Closing += MainWindow_Closing;
            #endregion
        }

        #region Close application

        private void MainWindow_Closing(object sender, CancelEventArgs e) { Ibp2018ViewModel vm = this.DataContext as Ibp2018ViewModel; vm.FinalizeForClose(); e.Cancel = !vm.Finalized; }

        private void MnuExit_Click(object sender, RoutedEventArgs e) { this.Close(); }

        #endregion

        #region Define combobox items

        void InitializeRibbonComboboxMember()
        {
            for (int i = 1; i <= 25; i++)
            {
                // Mininum resolution of 0.02 millimeters in X and Y axis
                catXStep.Items.Add((i * 0.02).ToString("F2"));
                catYStep.Items.Add((i * 0.02).ToString("F2"));
            }
            cboXStep.SelectedItem = catXStep.Items[0].ToString();
            cboYStep.SelectedItem = catYStep.Items[0].ToString();

            txtXMin.PreviewKeyUp += TxtXMinMax_PreviewKeyUp;
            txtXMax.PreviewKeyUp += TxtXMinMax_PreviewKeyUp;
            UpdateXScanRangeList();
            cboXStart.SelectedValue = catXStart.Items[0].ToString();
            cboXEnd.SelectedValue = catXEnd.Items[catXEnd.Items.Count - 1].ToString();
            txtYMin.PreviewKeyUp += TxtYMinMax_PreviewKeyUp;
            txtYMax.PreviewKeyUp += TxtYMinMax_PreviewKeyUp;
            UpdateYScanRangeList();
            cboYStart.SelectedValue = catYStart.Items[0].ToString();
            cboYEnd.SelectedValue = catYEnd.Items[catYEnd.Items.Count - 1].ToString();

            for (int i = 1; i <= 10; i++)
                catSensorInterval.Items.Add((i).ToString());
            cboSensorInterval.SelectedItem = catSensorInterval.Items[0].ToString();

            for (int i = 1; i <= 4; i++)
                catAveragingNumber.Items.Add((i).ToString());
            cboAveragingNumber.SelectedItem = catAveragingNumber.Items[0].ToString();
        }
        private void TxtXMinMax_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                this.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate () { UpdateXScanRangeList(); }));
            }
        }
        private void TxtYMinMax_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                this.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate () { UpdateYScanRangeList(); }));
            }
        }
        void UpdateXScanRangeList()
        {
            Ibp2018ViewModel vm = this.DataContext as Ibp2018ViewModel;
            catXStart.Items.Clear();
            catXEnd.Items.Clear();
            for (int i = 0; i < vm.XScanRangeList.Count; i++)
            {
                catXStart.Items.Add(vm.XScanRangeList[i].ToString());
                catXEnd.Items.Add(vm.XScanRangeList[i].ToString());
            }
            double pv;
            pv = Convert.ToDouble(cboXStart.SelectedValue);
            if (pv < vm.XScanRangeList[0]) cboXStart.SelectedValue = catXStart.Items[0];
            pv = Convert.ToDouble(cboXEnd.SelectedValue);
            if (vm.XScanRangeList[vm.XScanRangeList.Count - 1] < pv) cboXEnd.SelectedValue = catXEnd.Items[catXEnd.Items.Count - 1];
        }
        void UpdateYScanRangeList()
        {
            Ibp2018ViewModel vm = this.DataContext as Ibp2018ViewModel;
            catYStart.Items.Clear();
            catYEnd.Items.Clear();
            for (int i = 0; i < vm.YScanRangeList.Count; i++)
            {
                catYStart.Items.Add(vm.YScanRangeList[i].ToString());
                catYEnd.Items.Add(vm.YScanRangeList[i].ToString());
            }
            double pv;
            pv = Convert.ToDouble(cboYStart.SelectedValue);
            if (pv < vm.YScanRangeList[0]) cboYStart.SelectedValue = catYStart.Items[0];
            pv = Convert.ToDouble(cboYEnd.SelectedValue);
            if (vm.YScanRangeList[vm.YScanRangeList.Count - 1] < pv) cboYEnd.SelectedValue = catYEnd.Items[catYEnd.Items.Count - 1];
        }

        #endregion

        /// <summary>
        /// Jogging flag
        /// </summary>
        private volatile bool canJog = true;

        #region X Jog by button's mouse down and up

        #region X Jog background worker
        BackgroundWorker bwXJog = new BackgroundWorker();
        private void BwXJog_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) { canJog = false; }
        private void BwXJog_DoWork(object sender, DoWorkEventArgs e)
        {
            canJog = true;
            while (canJog)
            {
                this.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    Ibp2018ViewModel vm = this.DataContext as Ibp2018ViewModel;
                    PortResponse pr = vm.XJog(xJogValue);
                    if (pr.Code != PortResponse.SUCCESS && pr.Code != PortResponse.DEMO)
                    {
                        MessageBox.Show(pr.Message, "X Jog", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        canJog = false;
                        if (MessageBox.Show("Do you want to re-connect X-Y scanner MMC2 driver on " + vm.XyMmcPortName + "?", "Re-connect X-Y scanner MMC2 driver", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                        {
                            vm.ReconnectXyMmcCommand.Execute(new object());
                        }
                        else
                        {
                            // Do nothing
                        }
                    }
                }));
                Thread.Sleep(100);
            }
        }
        #endregion

        private int xJogValue = 50;
        private void MnuXJogPositive_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            canJog = false; //bwXJogPositive.CancelAsync();
        }
        private void MnuXJogPositive_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (bwXJog.IsBusy != true)
            {
                xJogValue = (int)(Convert.ToDouble(cboXStep.SelectedItem.ToString()) * MMC2Info.STEPPERMILLIMETER);
                bwXJog.RunWorkerAsync();
            }
        }
        private void MnuXJogNegative_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            canJog = false; //bwXJogPositive.CancelAsync();
        }
        private void MnuXJogNegative_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (bwXJog.IsBusy != true)
            {
                xJogValue = -(int)(Convert.ToDouble(cboXStep.SelectedItem.ToString()) * MMC2Info.STEPPERMILLIMETER);
                bwXJog.RunWorkerAsync();
            }
        }
        #endregion

        #region Y Jog by button's mouse down and up 

        #region Y Jog background worker
        BackgroundWorker bwYJog = new BackgroundWorker();
        private void BwYJog_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) { canJog = false; }
        private void BwYJog_DoWork(object sender, DoWorkEventArgs e)
        {
            canJog = true;
            while (canJog)
            {
                this.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    Ibp2018ViewModel vm = this.DataContext as Ibp2018ViewModel;
                    PortResponse pr = vm.YJog(yJogValue);
                    if (pr.Code != PortResponse.SUCCESS && pr.Code != PortResponse.DEMO)
                    {
                        MessageBox.Show(pr.Message, "Y Jog", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        canJog = false;
                        if (MessageBox.Show("Do you want to re-connect X-Y scanner MMC2 driver on " + vm.XyMmcPortName + "?", "Re-connect X-Y scanner MMC2 driver", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                        {
                            vm.ReconnectXyMmcCommand.Execute(new object());
                        }
                        else
                        {
                            // Do nothing
                        }
                    }
                }));
                Thread.Sleep(100);
            }
        }
        #endregion

        private int yJogValue = 50;
        private void MnuYJogPositive_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            canJog = false; //bwYJogPositive.CancelAsync();
        }
        private void MnuYJogPositive_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (bwYJog.IsBusy != true)
            {
                yJogValue = (int)(Convert.ToDouble(cboXStep.SelectedItem.ToString()) * MMC2Info.STEPPERMILLIMETER);
                bwYJog.RunWorkerAsync();
            }
        }
        private void MnuYJogNegative_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            canJog = false; //bwYJogPositive.CancelAsync();
        }
        private void MnuYJogNegative_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (bwYJog.IsBusy != true)
            {
                yJogValue = -(int)(Convert.ToDouble(cboYStep.SelectedItem.ToString()) * MMC2Info.STEPPERMILLIMETER);
                bwYJog.RunWorkerAsync();
            }
        }
        #endregion

        #region Z Jog by button's mouse down and up

        #region Z Jog background worker
        BackgroundWorker bwZJog = new BackgroundWorker();
        private void BwZJog_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) { canJog = false; }
        private void BwZJog_DoWork(object sender, DoWorkEventArgs e)
        {
            canJog = true;
            while (canJog)
            {
                this.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    Ibp2018ViewModel vm = this.DataContext as Ibp2018ViewModel;
                    PortResponse pr = vm.ZJog(zJogValue);
                    if (pr.Code != PortResponse.SUCCESS && pr.Code != PortResponse.DEMO)
                    {
                        MessageBox.Show(pr.Message, "Z Jog", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        canJog = false;
                        if (MessageBox.Show("Do you want to re-connect Z axis MMC2 driver on " + vm.ZMmcPortName + "?", "Re-connect Z axis MMC2 driver", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                        {
                            vm.ReconnectZMmcCommand.Execute(new object());
                        }
                        else
                        {
                            // Do nothing
                        }
                    }
                }));
                Thread.Sleep(100);
            }
        }
        #endregion

        private int zJogValue = 50;
        private void MnuZJogPositive_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            canJog = false; //bwZJogPositive.CancelAsync();
        }
        private void MnuZJogPositive_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (bwZJog.IsBusy != true)
            {
                zJogValue = (int)(Convert.ToDouble(cboXStep.SelectedItem.ToString()) * MMC2Info.STEPPERMILLIMETER);
                bwZJog.RunWorkerAsync();
            }
        }
        private void MnuZJogNegative_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            canJog = false; //bwZJogPositive.CancelAsync();
        }
        private void MnuZJogNegative_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (bwZJog.IsBusy != true)
            {
                zJogValue = -(int)(Convert.ToDouble(cboXStep.SelectedItem.ToString()) * MMC2Info.STEPPERMILLIMETER);
                bwZJog.RunWorkerAsync();
            }
        }
        #endregion

        #region New measurement

        #region Background worker of new measurement 
        BackgroundWorker bwNewMeasurement = new BackgroundWorker();
        WaitForNewMeasurementDialog popWaitNewMeasurement;
        private void BwNewMeasurement_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) { mnuNew.IsEnabled = true; popWaitNewMeasurement.Close(); }
        private void BwNewMeasurement_DoWork(object sender, DoWorkEventArgs e)
        {
            Thread.Sleep(10);
            mnuNew.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate () { mnuNew.IsEnabled = false; }));
            Ibp2018ViewModel vm = null;
            Ibp2018DataTableModel dt1 = null;
            Ibp2018DataTableModel dt2 = null;
            MMC2Info scn = null;
            bool running = true;
            this.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
            {
                vm = this.DataContext as Ibp2018ViewModel;
                dt1 = vm.CurrentDataTables[0];
                dt2 = vm.CurrentDataTables[1];
                scn = vm.XyMmc;
                double xStep = scn.XScanStep, yStep = scn.YScanStep;
                double xMin = Math.Min(scn.XScanStart, scn.XScanEnd), xMax = Math.Max(scn.XScanStart, scn.XScanEnd);
                double yMin = Math.Min(scn.YScanStart, scn.YScanEnd), yMax = Math.Max(scn.YScanStart, scn.YScanEnd);
                int colCount = (int)((xMax - xMin) / xStep) + 1;
                int rowCount = (int)((yMax - yMin) / yStep) + 1;
                dt1.ColumnNames = new List<string>();
                dt2.ColumnNames = new List<string>();
                dt1.ColumnHeaders = new List<string>();
                dt2.ColumnHeaders = new List<string>();
                dt1.Datatable.Rows.Clear();
                dt2.Datatable.Rows.Clear();
                dt1.Datatable.Columns.Clear();
                dt2.Datatable.Columns.Clear();
                dt1.Datatable.Columns.Add("Y_Step", typeof(string));
                dt2.Datatable.Columns.Add("Y_Step", typeof(string));
                dt1.ColumnNames.Add("Y_Step");
                dt2.ColumnNames.Add("Y_Step");
                dt1.ColumnHeaders.Add("Y Step");
                dt2.ColumnHeaders.Add("Y Step");
                for (int i = 0; i < colCount; i++)
                {
                    dt1.Datatable.Columns.Add("X_" + i.ToString(), typeof(string));
                    dt1.ColumnNames.Add("X_" + i.ToString());
                    dt1.ColumnHeaders.Add("X=" + (i * xStep).ToString("F2"));
                }
                running = false;
            }));
            while (running)
            {
                Thread.Sleep(10);
            }
            running = true;
            this.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
            {
                // Binding columns name and header
                vm.Current1ColumnCollection.Clear();
                for (int i = 0; i < dt1.ColumnNames.Count; i++)
                {
                    Binding binding = new Binding(dt1.ColumnNames[i]);
                    DataGridTextColumn textColumn = new DataGridTextColumn
                    {
                        Header = dt1.ColumnHeaders[i],
                        Binding = binding
                    };
                    vm.Current1ColumnCollection.Add(textColumn);
                }
                running = false;
            }));
            while (running)
            {
                Thread.Sleep(10);
            }

            vm.CurrentDataTables[0].NeedSave = true;
            vm.CurrentDataTables[1].NeedSave = true;

            System.Diagnostics.Trace.WriteLine(">>> " + DateTime.Now.ToString() + " End new measurement worker.");
        }
        #endregion

        /// <summary>
        /// New icon on ribbon button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MnuNew_Click(object sender, RoutedEventArgs e)
        {
            popWaitNewMeasurement = new WaitForNewMeasurementDialog
            {
                Topmost = true
            };
            popWaitNewMeasurement.Show();
            bwNewMeasurement.RunWorkerAsync();
        }

        #endregion

    }
}

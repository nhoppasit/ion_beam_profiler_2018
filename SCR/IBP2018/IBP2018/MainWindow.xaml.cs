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
using System.Data;

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
            // ----------------------------------------------------------
            // Initialize
            // ----------------------------------------------------------
            InitializeComponent();

            // ----------------------------------------------------------
            // current grid models
            // ----------------------------------------------------------
            var vm = this.DataContext as Ibp2018ViewModel;
            dgvCurrent1.Model = vm.CurrentGrid[0];
            dgvCurrent1.MinColumnWidthOverride = 10;
            dgvCurrent1.HeaderWidth = 60;

            // ----------------------------------------------------------
            // current grid models
            // ----------------------------------------------------------


            // ----------------------------------------------------------
            // Comboboxs
            // ----------------------------------------------------------
            InitializeRibbonComboboxMember();

            // ----------------------------------------------------------
            // Version
            // ----------------------------------------------------------
            this.Title = "Ion Beam Profiler Version: " + Assembly.GetExecutingAssembly().GetName().Version.ToString();

            // ----------------------------------------------------------
            // X jog background worker
            // ----------------------------------------------------------
            bwXJog.WorkerReportsProgress = true;
            bwXJog.DoWork += BwXJog_DoWork;
            bwXJog.RunWorkerCompleted += BwXJog_RunWorkerCompleted;

            // ----------------------------------------------------------
            // X negative jog
            // ----------------------------------------------------------
            mnuXJogNegative.PreviewMouseDown += MnuXJogNegative_PreviewMouseDown;
            mnuXJogNegative.PreviewMouseUp += MnuXJogNegative_PreviewMouseUp;

            // ----------------------------------------------------------
            // X positive jog
            // ----------------------------------------------------------
            mnuXJogPositive.PreviewMouseDown += MnuXJogPositive_PreviewMouseDown;
            mnuXJogPositive.PreviewMouseUp += MnuXJogPositive_PreviewMouseUp;

            // ----------------------------------------------------------
            // Y Jog background worker
            // ----------------------------------------------------------
            bwYJog.WorkerReportsProgress = true;
            bwYJog.DoWork += BwYJog_DoWork;
            bwYJog.RunWorkerCompleted += BwYJog_RunWorkerCompleted;

            // ----------------------------------------------------------
            // Y negative jog
            // ----------------------------------------------------------
            mnuYJogNegative.PreviewMouseDown += MnuYJogNegative_PreviewMouseDown;
            mnuYJogNegative.PreviewMouseUp += MnuYJogNegative_PreviewMouseUp;

            // Y positive jog
            mnuYJogPositive.PreviewMouseDown += MnuYJogPositive_PreviewMouseDown;
            mnuYJogPositive.PreviewMouseUp += MnuYJogPositive_PreviewMouseUp;

            // ----------------------------------------------------------
            // Z Jog background worker
            // ----------------------------------------------------------
            bwZJog.WorkerReportsProgress = true;
            bwZJog.DoWork += BwZJog_DoWork;
            bwZJog.RunWorkerCompleted += BwZJog_RunWorkerCompleted;

            // ----------------------------------------------------------
            // Y negative jog
            // ----------------------------------------------------------
            mnuZJogNegative.PreviewMouseDown += MnuZJogNegative_PreviewMouseDown;
            mnuZJogNegative.PreviewMouseUp += MnuZJogNegative_PreviewMouseUp;

            // ----------------------------------------------------------
            // Y positive jog
            // ----------------------------------------------------------
            mnuZJogPositive.PreviewMouseDown += MnuZJogPositive_PreviewMouseDown;
            mnuZJogPositive.PreviewMouseUp += MnuZJogPositive_PreviewMouseUp;

            // ----------------------------------------------------------
            // Exit application
            // ----------------------------------------------------------
            mnuExit.Click += MnuExit_Click;
            this.Closing += MainWindow_Closing;

            // ----------------------------------------------------------
            // Background for paste csv to cells
            // ----------------------------------------------------------
            bwPasteCSV.DoWork += BwPasteCSV_DoWork;
            bwPasteCSV.RunWorkerCompleted += BwPasteCSV_RunWorkerCompleted;
            bwPasteCSV.ProgressChanged += BwPasteCSV_ProgressChanged;
            bwPasteCSV.WorkerReportsProgress = true;
            bwPasteCSV.WorkerSupportsCancellation = true;

            // ----------------------------------------------------------
            // DATA GRID CURRENT1: Delete cell's content
            // ----------------------------------------------------------
            dgvCurrent1.KeyUp += DgvCurrent1_KeyUp;

        }

        // ------------------------------- CONSTRUCTOR ---------------------------------

        #region Datagrid cerrent1: keyup

        private void DgvCurrent1_KeyUp(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Delete:
                    DeleteSelectedCellContent();
                    break;
                default:
                    break;
            }
        }

        void DeleteSelectedCellContent()
        {
            MessageBoxResult mbr = MessageBox.Show("DATA WILL BE PERMANENT DELETED. Press [OK] to delete or [CANCEL] to cancel.", "Delete cell content", MessageBoxButton.OKCancel, MessageBoxImage.Question);
            if (mbr == MessageBoxResult.OK)
            {
                int row = (int)dgvCurrent1.CurrentRow;
                int col = (int)dgvCurrent1.CurrentColumn;
                var model = (this.DataContext as Ibp2018ViewModel).CurrentGrid[0] as CurrentGridModel;
                var sc = dgvCurrent1.GetSelectedModelCells();
                foreach (FastWpfGrid.FastGridCellAddress ce in sc)
                {
                    var key = Tuple.Create((int)ce.Row, (int)ce.Column);
                    if (model.EditedCells.ContainsKey(key)) model.EditedCells.Remove(key);
                }
                model.InvalidateAll();
            }
            else
            {
                /*DONOTHING*/
            }
        }

        #endregion

        #region Paste clipboard to cells ---------------------------------------------------------------------

        BackgroundWorker bwPasteCSV = new BackgroundWorker();
        private void BwPasteCSV_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            lblStatus.Content = string.Format("Paste from clipboard: {0}...", e.ProgressPercentage);
            pgStatus.Value = e.ProgressPercentage;
        }
        private void BwPasteCSV_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            var model = (this.DataContext as Ibp2018ViewModel).CurrentGrid[0] as CurrentGridModel;
            model.InvalidateAll();
            lblStatus.Content = string.Empty;
            pgStatus.Value = 0;
            pgStatus.Visibility = Visibility.Hidden;
        }
        private void BwPasteCSV_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = (BackgroundWorker)sender;
            object[] varargin = (object[])e.Argument;
            string text = (string)varargin[0];
            var model = varargin[1] as CurrentGridModel;

            int row = (int)dgvCurrent1.CurrentRow;
            int col = (int)dgvCurrent1.CurrentColumn;

            bool asked = false, replaceFlag = false;
            int r = 0, c = 0;
            string[] rText = text.Replace("\r\n", ";").Split(';');
            foreach (string s in rText)
            {
                string[] cText = s.Split(',');
                c = 0;
                foreach (string t in cText)
                {
                    try
                    {
                        if (model.EditedCells[Tuple.Create(row + r, col + c)] != null)
                        {
                            if (!asked)
                            {
                                if (MessageBox.Show("Do you want to replace the non-empty cells?", "Paste", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                                    replaceFlag = true;
                                else
                                    replaceFlag = false;
                                asked = true;
                            }
                            if (replaceFlag)
                                if (double.TryParse(t, out double dbl)) model.EditedCells[Tuple.Create(row + r, col + c)] = dbl;
                        }
                    }
                    catch
                    {
                        if (double.TryParse(t, out double dbl)) model.EditedCells[Tuple.Create(row + r, col + c)] = dbl;
                    }
                    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(delegate ()
                    {
                        model.InvalidateCell(row + r, col + c);
                        //model.InvalidateAll();
                    })); // ไม่ค่อยเห็นว่า invalidate                    
                    c++;
                    if (model.ColumnCount < col + c) break;
                }
                r++;
                if (model.RowCount < row + r) break;
                worker.ReportProgress((int)((float)r / (float)rText.Length * 100));
            }
            model.InvalidateAll();
        }
        private void MnuC1Paste_from_clipboard_Click(object sender, RoutedEventArgs e)
        {
            //var model = (this.DataContext as Ibp2018ViewModel).CurrentGrid[0] as CurrentGridModel;
            string text = Clipboard.GetText(TextDataFormat.CommaSeparatedValue);
            if (text == "" || text == null)
            {
                return;
            }
            else
            {
                if (!bwPasteCSV.IsBusy)
                {
                    pgStatus.Visibility = Visibility.Visible;
                    bwPasteCSV.RunWorkerAsync(new object[] { text, (this.DataContext as Ibp2018ViewModel).CurrentGrid[0] as CurrentGridModel });
                }
            }
        }

        #endregion

        #region Close application ------------------------------------------------------------
        private void MainWindow_Closing(object sender, CancelEventArgs e) { Ibp2018ViewModel vm = this.DataContext as Ibp2018ViewModel; vm.FinalizeForClose(); e.Cancel = !vm.Finalized; }
        private void MnuExit_Click(object sender, RoutedEventArgs e) { this.Close(); }
        #endregion

        #region Define combobox items ------------------------------------------------------------
        void InitializeRibbonComboboxMember()
        {
            // ---------------------------------------------------------
            // X-Y Scanner resolution
            // ---------------------------------------------------------
            for (int i = 1; i <= 25; i++)
            {
                // Mininum resolution of 0.02 millimeters in X and Y axis
                catXStep.Items.Add((i * 0.02).ToString("F2"));
                catYStep.Items.Add((i * 0.02).ToString("F2"));
            }
            cboXStep.SelectedItem = catXStep.Items[0].ToString();
            cboYStep.SelectedItem = catYStep.Items[0].ToString();

            // ---------------------------------------------------------
            // Scan region of X-Y scanner
            // ---------------------------------------------------------
            txtXMin.PreviewKeyUp += TxtXMinMax_PreviewKeyUp; // re-assign items to combobox both of min and max
            txtXMax.PreviewKeyUp += TxtXMinMax_PreviewKeyUp;
            UpdateXScanRangeList();
            cboXStart.SelectedValue = catXStart.Items[0].ToString();
            cboXEnd.SelectedValue = catXEnd.Items[catXEnd.Items.Count - 1].ToString();
            //
            txtYMin.PreviewKeyUp += TxtYMinMax_PreviewKeyUp; // re-assign items to combobox both of min and max
            txtYMax.PreviewKeyUp += TxtYMinMax_PreviewKeyUp;
            UpdateYScanRangeList();
            cboYStart.SelectedValue = catYStart.Items[0].ToString();
            cboYEnd.SelectedValue = catYEnd.Items[catYEnd.Items.Count - 1].ToString();

            // ---------------------------------------------------------
            // DMMs meas interval
            // ---------------------------------------------------------
            for (int i = 1; i <= 10; i++)
                catSensorInterval.Items.Add((i).ToString());
            cboSensorInterval.SelectedItem = catSensorInterval.Items[0].ToString();

            // ---------------------------------------------------------
            // DMMs meas average
            // ---------------------------------------------------------
            for (int i = 1; i <= 4; i++)
                catAveragingNumber.Items.Add((i).ToString());
            cboAveragingNumber.SelectedItem = catAveragingNumber.Items[0].ToString();
        }
        private void TxtXMinMax_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                this.Dispatcher.Invoke(DispatcherPriority.ApplicationIdle, new Action(delegate () { UpdateXScanRangeList(); }));
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

        #region X Jog by button's mouse down and up ------------------------------------------------------------
        private volatile bool canJog = true; // Jogging flag
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

        #region Y Jog by button's mouse down and up ------------------------------------------------------------
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

        #region Z Jog by button's mouse down and up ------------------------------------------------------------
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

    }
}

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

namespace IBP2018
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : RibbonWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            InitializeRibbonComboboxMember();
            SetStatesText();

            // "Version: " + Assembly.GetExecutingAssembly().GetName().Version.ToString();

            mnuDemoData.Click += MnuDemoData_Click;
            chkAutoQueryScanner.Checked += ChkAutoQueryScanner_Checked;
            chkAutoQueryScanner.Unchecked += ChkAutoQueryScanner_Unchecked;
        }

        private void ChkAutoQueryScanner_Unchecked(object sender, RoutedEventArgs e)
        {
            Ibp2018ViewModel vm = this.mainGrid.DataContext as Ibp2018ViewModel;
            vm.UnqueryScanner.Execute(new object());
        }

        private void ChkAutoQueryScanner_Checked(object sender, RoutedEventArgs e)
        {
            Ibp2018ViewModel vm = this.mainGrid.DataContext as Ibp2018ViewModel;
            vm.QueryScanner.Execute(new object());
        }

        BackgroundWorker bwDefineTableColumns = new BackgroundWorker();
        private void MnuDemoData_Click(object sender, RoutedEventArgs e)
        {
            Ibp2018ViewModel vm = this.mainGrid.DataContext as Ibp2018ViewModel;
            WaitDialog dlg = new WaitDialog();
            dlg.Title = "Wait";
            dlg.Topmost = true;
            bwDefineTableColumns.DoWork += (o, ea) =>
            {
                Thread.Sleep(100);
                dlg.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    try
                    {
                        dlg.Visibility = Visibility.Visible;
                        dlg.Show();
                    }
                    catch { }
                    Thread.Sleep(100);
                }));
                Stopwatch sw = new Stopwatch();
                sw.Start();
                bool _continue = true;
                while (_continue)
                {
                    // Timeout
                    if (sw.ElapsedMilliseconds > 1000 * 10)
                    {
                        sw.Stop();
                        _continue = false;
                    }
                    // Break by columns count
                    if (!vm.ColumnsGenerating) _continue = false;
                    Thread.Sleep(100);
                }
            };
            bwDefineTableColumns.RunWorkerCompleted += (o, ea) =>
            {
                dlg.Close();
            };
            dlg.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
            {
                dlg.Visibility = Visibility.Visible;
                dlg.Show();
                Thread.Sleep(100);
            }));
            bwDefineTableColumns.RunWorkerAsync();
        }


        // เพิ่มค่า Step
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

            for (int i = 0; i <= 20; i += 5)
            {
                catXStart.Items.Add((i).ToString());
                catXEnd.Items.Add((i).ToString());
                catYStart.Items.Add((i).ToString());
                catYEnd.Items.Add((i).ToString());
            }
            cboXStart.SelectedItem = catXStart.Items[0].ToString();
            cboXEnd.SelectedItem = catXEnd.Items[3].ToString();
            cboYStart.SelectedItem = catYStart.Items[0].ToString();
            cboYEnd.SelectedItem = catYEnd.Items[3].ToString();

            for (int i = 1; i <= 10; i++)
                catSensorInterval.Items.Add((i).ToString());
            cboSensorInterval.SelectedItem = catSensorInterval.Items[0].ToString();
        }

        void SetStatesText()
        {
            lbScannerSetting1.Dispatcher.Invoke(DispatcherPriority.Normal, (Action)delegate
             {
                 StringBuilder sb = new StringBuilder();
                 sb.Append("Step = ( "); sb.Append(cboXStep.SelectedItem.ToString()); sb.Append(", "); sb.Append(cboYStep.SelectedItem.ToString()); sb.Append(" ) mm");
                 sb.Append(Environment.NewLine);
                 sb.Append("X Range = ( "); sb.Append(cboXStart.SelectedItem.ToString()); sb.Append(", "); sb.Append(cboXEnd.SelectedItem.ToString()); sb.Append(" ) mm");
                 lbScannerSetting1.Content = sb.ToString();

                 sb.Clear();
                 sb.Append("Sampling rate = "); sb.Append(cboSensorInterval.SelectedItem.ToString()); sb.Append(" sec/Point");
                 sb.Append(Environment.NewLine);
                 sb.Append("Y Range = ( "); sb.Append(cboYStart.SelectedItem.ToString()); sb.Append(", "); sb.Append(cboYEnd.SelectedItem.ToString()); sb.Append(" ) mm");
                 lbScannerSetting2.Content = sb.ToString();

             });
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace IBP2018.View
{
    /// <summary>
    /// Interaction logic for WaitForNewMeasurementDialog.xaml
    /// </summary>
    public partial class WaitForNewMeasurementDialog : Window
    {
        public WaitForNewMeasurementDialog()
        {
            InitializeComponent();
            bw.DoWork += Bw_DoWork;
            bw.RunWorkerCompleted += Bw_RunWorkerCompleted;
            this.Loaded += WaitForNewMeasurementDialog_Loaded;
            pgWait.Value = 0;
        }

        #region Auto progress trigger
        private void WaitForNewMeasurementDialog_Loaded(object sender, RoutedEventArgs e)
        {
            bw.RunWorkerAsync();
        }
        BackgroundWorker bw = new BackgroundWorker();
        private bool stopFlag = false;
        private void Bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // Donothing
        }
        private void Bw_DoWork(object sender, DoWorkEventArgs e)
        {
            double v = 0;
            while (!stopFlag)
            {
                v += 5;
                if (100 < v) v = 0;
                SetProgressValue(v);
                System.Threading.Thread.Sleep(70);
            }
        }
        public void Stop() { stopFlag = true; }
        #endregion

        /// <summary>
        /// Set progress bar value delegate
        /// </summary>
        /// <param name="e">Value is 0.0-1.0</param>
        public void SetProgressValue(double e)
        {
            this.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
            {
                pgWait.Value = e;
                pgWait.InvalidateVisual();
            }));
        }
    }
}

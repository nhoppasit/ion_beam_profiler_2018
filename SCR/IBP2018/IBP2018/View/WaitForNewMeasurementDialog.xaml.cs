using System;
using System.Collections.Generic;
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
        }

        public double ProgressValue
        {
            get { return pgWait.Value; }
            set
            {
                SetProgressValue(value);
            }
        }

        /// <summary>
        /// Set progress bar value delegate
        /// </summary>
        /// <param name="e">Value is 0.0-1.0</param>
        public void SetProgressValue(double e)
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
            {
                pgWait.Value = e;
            }));
        }
    }
}

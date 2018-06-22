﻿using System;
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
using Quintessence.MotionControl.MMC2;
using System.Reflection;

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

            // Auto query
            chkAutoQueryScanner.Checked += ChkAutoQueryScanner_Checked;
            chkAutoQueryScanner.Unchecked += ChkAutoQueryScanner_Unchecked;

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


        }

        #region Initialize application
        // Initialize combobox
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
        #endregion

        private volatile bool canJog = true;

        #region X Jog

        #region X Jog background worker
        BackgroundWorker bwXJog = new BackgroundWorker();
        private void BwXJog_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            canJog = false;
        }
        private void BwXJog_DoWork(object sender, DoWorkEventArgs e)
        {
            canJog = true;
            while (canJog)
            {
                this.mainGrid.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    Ibp2018ViewModel vm = this.mainGrid.DataContext as Ibp2018ViewModel;
                    vm.XJog(xJogValue);
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

        #region Y Jog

        #region Y Jog background worker
        BackgroundWorker bwYJog = new BackgroundWorker();
        private void BwYJog_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            canJog = false;
        }
        private void BwYJog_DoWork(object sender, DoWorkEventArgs e)
        {
            canJog = true;
            while (canJog)
            {
                this.mainGrid.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    Ibp2018ViewModel vm = this.mainGrid.DataContext as Ibp2018ViewModel;
                    vm.YJog(xJogValue);
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
                yJogValue = -(int)(Convert.ToDouble(cboXStep.SelectedItem.ToString()) * MMC2Info.STEPPERMILLIMETER);
                bwYJog.RunWorkerAsync();
            }
        }
        #endregion

        #region Z Jog

        #region Z Jog background worker
        BackgroundWorker bwZJog = new BackgroundWorker();
        private void BwZJog_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            canJog = false;
        }
        private void BwZJog_DoWork(object sender, DoWorkEventArgs e)
        {
            canJog = true;
            while (canJog)
            {
                this.mainGrid.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    Ibp2018ViewModel vm = this.mainGrid.DataContext as Ibp2018ViewModel;
                    vm.ZJog(xJogValue);
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


        #region Auto query scanner
        private void ChkAutoQueryScanner_Unchecked(object sender, RoutedEventArgs e)
        {
            Ibp2018ViewModel vm = this.mainGrid.DataContext as Ibp2018ViewModel;
            vm.UnqueryScannerCommand.Execute(new object());
        }
        private void ChkAutoQueryScanner_Checked(object sender, RoutedEventArgs e)
        {
            Ibp2018ViewModel vm = this.mainGrid.DataContext as Ibp2018ViewModel;
            vm.QueryScannerCommand.Execute(new object());
        }
        #endregion               
    }
}

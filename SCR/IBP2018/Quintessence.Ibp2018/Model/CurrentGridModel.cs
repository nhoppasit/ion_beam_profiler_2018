using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using FastWpfGrid;

namespace Quintessence.Ibp2018.Model
{
    public class CurrentGridModel : FastGridModelBase
    {
        #region CONSTRUCTOR
        public CurrentGridModel()
        {
            bwPasteCSV.DoWork += BwPasteCSV_DoWork;
            bwPasteCSV.ProgressChanged += BwPasteCSV_ProgressChanged;
            bwPasteCSV.RunWorkerCompleted += BwPasteCSV_RunWorkerCompleted;
            bwPasteCSV.WorkerReportsProgress = true;
            bwPasteCSV.WorkerSupportsCancellation = true;
        }
        #endregion


        #region Variables -----------------------------------------------------------------------------------
        public Dictionary<Tuple<int, int>, double> EditedCells = new Dictionary<Tuple<int, int>, double>();
        #endregion

        #region Row and column count ------------------------------------------------------------------------
        public override int ColumnCount
        {
            get { return 70; }
        }
        public override int RowCount
        {
            get { return 1000; }
        }
        #endregion

        #region Cell editing --------------------------------------------------------------------------------
        public override string GetCellText(int row, int column)
        {
            var key = Tuple.Create(row, column);
            if (EditedCells.ContainsKey(key)) return String.Format("{0,8:F2}", EditedCells[key]);
            return "";
        }
        public override void SetCellText(int row, int column, string value)
        {
            var key = Tuple.Create(row, column);
            if (double.TryParse(value, out double dbl))
                EditedCells[key] = dbl;
        }
        #endregion

        #region Cell editing --------------------------------------------------------------------------------
        public override void HandleSelectionCommand(IFastGridView view, string command)
        {
            MessageBox.Show(command);// command ของ cell selection
        }
        #endregion

        #region Header of row and columns -------------------------------------------------------------------
        public override IFastGridCell GetColumnHeader(IFastGridView view, int column)
        {
            var res = new FastGridCellImpl();
            res.Blocks.Add(new FastGridBlockImpl
            {
                IsBold = false,
                TextData = string.Format("X={0:0.00}", column * 0.02),
            });
            return res;
        }
        public override IFastGridCell GetRowHeader(IFastGridView view, int row)
        {
            var res = new FastGridCellImpl();
            res.Blocks.Add(new FastGridBlockImpl
            {
                IsBold = false,
                TextData = String.Format("Y={0:0.00}", row * 0.02)
            });
            return res;
        }
        #endregion

        #region Paste thread -------------------------------------------------------------------------------
        ProgressBar pgStatus;
        Label lblStatus;
        BackgroundWorker bwPasteCSV = new BackgroundWorker();
        private void BwPasteCSV_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            lblStatus.Content = string.Format("Paste from clipboard: {0}...", e.ProgressPercentage);
            pgStatus.Value = e.ProgressPercentage;
        }
        private void BwPasteCSV_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.InvalidateAll();
            lblStatus.Content = String.Empty;
            pgStatus.Value = 0;
            pgStatus.Visibility = Visibility.Hidden;
        }
        private void BwPasteCSV_DoWork(object sender, DoWorkEventArgs e)
        {
            // sneder and parameters
            BackgroundWorker worker = (BackgroundWorker)sender;
            object[] varargin = (object[])e.Argument;
            string text = (string)varargin[0];
            int row = (int)varargin[1];            
            int col = (int)varargin[2];

            // start pasting
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
                        if (this.EditedCells[Tuple.Create(row + r, col + c)] != null)
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
                                if (double.TryParse(t, out double dbl)) this.EditedCells[Tuple.Create(row + r, col + c)] = dbl;
                        }
                    }
                    catch
                    {
                        if (double.TryParse(t, out double dbl)) this.EditedCells[Tuple.Create(row + r, col + c)] = dbl;
                    }                    
                    c++;
                    if (this.ColumnCount < col + c) break;
                }
                r++;
                if (this.RowCount < row + r) break;
                worker.ReportProgress((int)((float)r / (float)rText.Length * 100));
            }
            this.InvalidateAll();
        }
        /// <summary>
        /// เริ่มต้นธีด สำหรับวางข้อความจาก clipboard
        /// </summary>
        /// <param name="objArray">1. Status Label Control,  2. Progress Bar Control,  3. Row index,  4. Column index</param>
        public void StartPastingFromClipboardThread(object[] objArray)
        {
            lblStatus = (Label)objArray[0];
            pgStatus = (ProgressBar)objArray[1];
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
                    bwPasteCSV.RunWorkerAsync(new object[] { text/*from clipboard*/, (int)objArray[2]/*number of row*/, (int)objArray[3] /*number of column*/ });
                }
            }
        }
        #endregion
    }
}

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
        #region CONSTRUCTOR ---------------------------------------------------------------------------------
        public CurrentGridModel()
        {
            // --------------------------------------------------------
            // Paste from clipboard
            // --------------------------------------------------------
            bgwPasteCSV.DoWork += BgwPasteCSV_DoWork;
            bgwPasteCSV.ProgressChanged += BgwPasteCSV_ProgressChanged;
            bgwPasteCSV.RunWorkerCompleted += BgwPasteCSV_RunWorkerCompleted;
            bgwPasteCSV.WorkerReportsProgress = true;
            bgwPasteCSV.WorkerSupportsCancellation = true;

            // --------------------------------------------------------
            // Copy to clipboard
            // --------------------------------------------------------
            bgwCopyCSV.DoWork += BgwCopyCSV_DoWork;
            bgwCopyCSV.ProgressChanged += BgwCopyCSV_ProgressChanged;
            bgwCopyCSV.RunWorkerCompleted += BgwCopyCSV_RunWorkerCompleted;
            bgwCopyCSV.WorkerReportsProgress = true;
            bgwCopyCSV.WorkerSupportsCancellation = true;
        }
        #endregion

        #region Variables -----------------------------------------------------------------------------------
        ProgressBar pgStatus;
        Label lblStatus;
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
        BackgroundWorker bgwPasteCSV = new BackgroundWorker();
        private void BgwPasteCSV_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            lblStatus.Content = string.Format("Paste from clipboard...{0:#}%", e.ProgressPercentage);
            pgStatus.Value = e.ProgressPercentage;
        }
        private void BgwPasteCSV_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.InvalidateAll();
            lblStatus.Content = String.Empty;
            pgStatus.Value = 0;
            pgStatus.Visibility = Visibility.Hidden;
        }
        private void BgwPasteCSV_DoWork(object sender, DoWorkEventArgs e)
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
                if (!bgwPasteCSV.IsBusy)
                {
                    pgStatus.Visibility = Visibility.Visible;
                    bgwPasteCSV.RunWorkerAsync(new object[] { text/*from clipboard*/, (int)objArray[2]/*number of row*/, (int)objArray[3] /*number of column*/ });
                }
            }
        }
        #endregion

        #region Paste thread -------------------------------------------------------------------------------
        StringBuilder sbCopy = new StringBuilder();
        BackgroundWorker bgwCopyCSV = new BackgroundWorker();
        private void BgwCopyCSV_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            lblStatus.Content = string.Format("Copying...{0:#}%", e.ProgressPercentage);
            pgStatus.Value = e.ProgressPercentage;
        }
        private void BgwCopyCSV_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        { 
            Clipboard.SetText(sbCopy.ToString(), TextDataFormat.CommaSeparatedValue);
            this.InvalidateAll();
            lblStatus.Content = String.Empty;
            pgStatus.Value = 0;
            pgStatus.Visibility = Visibility.Hidden;
        }
        private void BgwCopyCSV_DoWork(object sender, DoWorkEventArgs e)
        {
            // sneder and parameters
            BackgroundWorker worker = (BackgroundWorker)sender;
            object[] varargin = (object[])e.Argument;
            var selectedCellAddresses = (HashSet<FastGridCellAddress>)varargin[0];
            Dictionary<Tuple<int, int>, double?> selectedCells = new Dictionary<Tuple<int, int>, double?>();
            int r0 = int.MaxValue, rm = int.MinValue, c0 = int.MaxValue, cm = int.MinValue;
            int p = 0;
            foreach (FastWpfGrid.FastGridCellAddress cell in selectedCellAddresses)
            {
                // update range
                r0 = Math.Min(r0, (int)cell.Row);
                rm = Math.Max(rm, (int)cell.Row);
                c0 = Math.Min(c0, (int)cell.Column);
                cm = Math.Max(cm, (int)cell.Column);

                // delete and save to temp dict
                var key = Tuple.Create((int)cell.Row, (int)cell.Column);
                if (this.EditedCells.ContainsKey(key))
                {
                    selectedCells[key] = this.EditedCells[key];
                }
                else
                {
                    selectedCells[key] = null;
                }
                bgwCopyCSV.ReportProgress(p / selectedCellAddresses.Count * 50);
            }
            this.InvalidateAll();
            bgwCopyCSV.ReportProgress(50);

            // Set clipboard text as selected
            p = 0;
            sbCopy.Clear();
            for (int r = r0; r <= rm; r++)
            {
                for (int c = c0; c <= cm; c++)
                {
                    var key = Tuple.Create(r, c);
                    sbCopy.Append(selectedCells[key].ToString());
                    if (c < cm) sbCopy.Append(",");
                }
                if (r < rm) sbCopy.Append("\r\n");
                bgwCopyCSV.ReportProgress(p / (rm * cm) * 100); 
            }            
        }
        /// <summary>
        /// เริ่มต้นธีด สำหรับสำเนาข้อความไปไว้ที่ clipboard
        /// </summary>
        /// <param name="objArray">1. Status Label Control,  2. Progress Bar Control,  3. HashSet<FastGridCellAddress> obj เป็นรายการจาก FastWpfGrid.GetSelectedModelCells()</param>
        public void StartCopyngToClipboardThread(object[] objArray)
        {
            lblStatus = (Label)objArray[0];
            pgStatus = (ProgressBar)objArray[1];
            lblStatus.Content = "Copying...";
            var selectedCellAddresses = (HashSet<FastGridCellAddress>)objArray[2];
            if (!bgwCopyCSV.IsBusy)
            {
                pgStatus.Visibility = Visibility.Visible;
                bgwCopyCSV.RunWorkerAsync(new object[] { selectedCellAddresses });
            }
        }
        #endregion
    }
}

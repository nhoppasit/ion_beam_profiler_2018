using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
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

            // --------------------------------------------------------
            // cut to clipboard
            // --------------------------------------------------------
            bgwDeleteContents.DoWork += BgwDeleteContents_DoWork;
            bgwDeleteContents.ProgressChanged += BgwDeleteContents_ProgressChanged;
            bgwDeleteContents.RunWorkerCompleted += BgwDeleteContents_RunWorkerCompleted;
            bgwDeleteContents.WorkerReportsProgress = true;
            bgwDeleteContents.WorkerSupportsCancellation = true;

            // --------------------------------------------------------
            // delete
            // --------------------------------------------------------
            bgwCutContents.DoWork += BgwCutContents_DoWork;
            bgwCutContents.ProgressChanged += BgwCutContents_ProgressChanged;
            bgwCutContents.RunWorkerCompleted += BgwCutContents_RunWorkerCompleted;
            bgwCutContents.WorkerReportsProgress = true;
            bgwCutContents.WorkerSupportsCancellation = true;
        }
        #endregion

        #region Variables -----------------------------------------------------------------------------------
        StringBuilder sbCopy = new StringBuilder();
        ProgressBar pgStatus;
        Label lblStatus;
        public Dictionary<Tuple<int, int>, double> EditedCells = new Dictionary<Tuple<int, int>, double>();
        #endregion

        #region Row and column count ------------------------------------------------------------------------
        private int _columnCount = 30;
        public int ColumnSize { set { _columnCount = value; } }
        public override int ColumnCount
        {
            get { return 10; }
        }
        private int _rowCount = 30;
        public int RowSize { set { _rowCount = value; } }
        public override int RowCount
        {
            get { return 10; }
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
            if (lblStatus != null) lblStatus.Content = string.Format("Paste from clipboard...{0:#}%", e.ProgressPercentage);
            if (pgStatus != null) pgStatus.Value = e.ProgressPercentage;
        }
        private void BgwPasteCSV_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.InvalidateAll();
            if (lblStatus != null) lblStatus.Content = String.Empty;
            if (pgStatus != null)
            {
                pgStatus.Value = 0;
                pgStatus.Visibility = Visibility.Hidden;
            }
        }
        private void BgwPasteCSV_DoWork(object sender, DoWorkEventArgs e)
        {
            // sneder and parameters
            BackgroundWorker worker = (BackgroundWorker)sender;
            var varargin = (PasteToGridContentObject)e.Argument;
            string text = varargin.Text;
            int row = varargin.Row;
            int col = varargin.Column;

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
        public void StartPastingFromClipboardThread(PasteToGridContentObject e)
        {
            try { lblStatus = e.StatusLabel; } catch { lblStatus = null; }
            try { pgStatus = e.StatusProgressBar; } catch { pgStatus = null; }
            e.Text = Clipboard.GetText(TextDataFormat.CommaSeparatedValue);
            if (e.Text == "" || e.Text == null)
            {
                return;
            }
            else
            {
                if (!bgwPasteCSV.IsBusy)
                {
                    pgStatus.Visibility = Visibility.Visible;
                    bgwPasteCSV.RunWorkerAsync(e);
                }
            }
        }
        #endregion

        #region Paste thread -------------------------------------------------------------------------------
        BackgroundWorker bgwCopyCSV = new BackgroundWorker();
        private void BgwCopyCSV_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage < 100)
            {
                if (lblStatus != null) lblStatus.Content = string.Format("Copying...{0:#}%", e.ProgressPercentage);
                if (pgStatus != null) pgStatus.Value = e.ProgressPercentage;
            }
            else
            {
                if (lblStatus != null) lblStatus.Content = "Done copying to clipboard";
                if (pgStatus != null) pgStatus.Visibility = Visibility.Hidden;
            }

        }
        private void BgwCopyCSV_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Clipboard.SetText(sbCopy.ToString(), TextDataFormat.CommaSeparatedValue);
            this.InvalidateAll();
            if (lblStatus != null) lblStatus.Content = string.Empty;
            if (pgStatus != null)
            {
                pgStatus.Value = 0;
                pgStatus.Visibility = Visibility.Hidden;
            }
        }
        private void BgwCopyCSV_DoWork(object sender, DoWorkEventArgs e)
        {
            // sneder and parameters
            BackgroundWorker worker = (BackgroundWorker)sender;
            var varargin = (SelectedGridContentsObject)e.Argument;

            // Start
            Dictionary<Tuple<int, int>, double?> selectedCells = new Dictionary<Tuple<int, int>, double?>();
            int r0 = int.MaxValue, rm = int.MinValue, c0 = int.MaxValue, cm = int.MinValue;
            double p = 0;
            foreach (FastWpfGrid.FastGridCellAddress cell in varargin.SelectedCellAddresses)
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
                bgwCopyCSV.ReportProgress((int)(p++ / varargin.SelectedCellAddresses.Count * 50));
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
                    bgwCopyCSV.ReportProgress((int)(p++ / ((rm - r0 + 1) * (cm - c0 + 1)) * 100));
                }
                if (r < rm) sbCopy.Append("\r\n");
            }
            bgwCopyCSV.ReportProgress(100);
            Thread.Sleep(1500);
        }
        /// <summary>
        /// เริ่มต้นธีด สำหรับสำเนาข้อความไปไว้ที่ clipboard
        /// </summary>
        /// <param name="objArray">1. Status Label Control,  2. Progress Bar Control,  3. HashSet<FastGridCellAddress> obj เป็นรายการจาก FastWpfGrid.GetSelectedModelCells()</param>
        public void StartCopyingToClipboardThread(SelectedGridContentsObject e)
        {
            try { lblStatus = e.StatusLabel; } catch { lblStatus = null; }
            try { pgStatus = e.StatusProgressBar; } catch { pgStatus = null; }
            if (lblStatus != null) lblStatus.Content = "Copying...";
            if (!bgwCopyCSV.IsBusy)
            {
                pgStatus.Visibility = Visibility.Visible;
                bgwCopyCSV.RunWorkerAsync(e);
            }
        }
        #endregion

        #region Delete contents thread -------------------------------------------------------------------------------
        BackgroundWorker bgwDeleteContents = new BackgroundWorker();
        private void BgwDeleteContents_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (lblStatus != null) lblStatus.Content = string.Format("Deleting...{0:#}%", e.ProgressPercentage);
            if (pgStatus != null) pgStatus.Value = e.ProgressPercentage;
        }
        private void BgwDeleteContents_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.InvalidateAll();
            if (lblStatus != null) lblStatus.Content = String.Empty;
            if (pgStatus != null)
            {
                pgStatus.Value = 0;
                pgStatus.Visibility = Visibility.Hidden;
            }
        }
        private void BgwDeleteContents_DoWork(object sender, DoWorkEventArgs e)
        {
            // sneder and parameters
            BackgroundWorker worker = (BackgroundWorker)sender;
            var varargin = (SelectedGridContentsObject)e.Argument;

            // Start
            foreach (FastWpfGrid.FastGridCellAddress ce in varargin.SelectedCellAddresses)
            {
                var key = Tuple.Create((int)ce.Row, (int)ce.Column);
                if (this.EditedCells.ContainsKey(key)) this.EditedCells.Remove(key);
            }
        }
        /// <summary>
        /// เริ่มต้นธีด ลบ ข้อมูลในเซลล์ที่เลือกไว้
        /// </summary>
        /// <param name="objArray">1. Status Label Control,  2. Progress Bar Control,  3. HashSet<FastGridCellAddress> obj เป็นรายการจาก FastWpfGrid.GetSelectedModelCells()</param>
        public void StartDeleteContentsThread(SelectedGridContentsObject e)
        {
            MessageBoxResult mbr = MessageBox.Show("DATA WILL BE PERMANENT DELETED. Press [OK] to delete or [CANCEL] to cancel.", "Delete cell content", MessageBoxButton.OKCancel, MessageBoxImage.Question);
            if (mbr == MessageBoxResult.OK)
            {
                try { lblStatus = e.StatusLabel; } catch { lblStatus = null; }
                try { pgStatus = e.StatusProgressBar; } catch { pgStatus = null; }
                if (lblStatus != null) lblStatus.Content = "Deleting...";
                if (!bgwDeleteContents.IsBusy)
                {
                    pgStatus.Visibility = Visibility.Visible;
                    bgwDeleteContents.RunWorkerAsync(e);
                }
            }
        }
        #endregion

        #region Cut content thread -------------------------------------------------------------------------------
        BackgroundWorker bgwCutContents = new BackgroundWorker();
        private void BgwCutContents_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage < 100)
            {
                if (lblStatus != null) lblStatus.Content = string.Format("Cutting...{0:#}%", e.ProgressPercentage);
                if (pgStatus != null) pgStatus.Value = e.ProgressPercentage;
            }
            else
            {
                if (lblStatus != null) lblStatus.Content = "Done copying to clipboard";
                if (pgStatus != null) pgStatus.Visibility = Visibility.Hidden;
            }

        }
        private void BgwCutContents_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Clipboard.SetText(sbCopy.ToString(), TextDataFormat.CommaSeparatedValue);
            this.InvalidateAll();
            if (lblStatus != null) lblStatus.Content = string.Empty;
            if (pgStatus != null)
            {
                pgStatus.Value = 0;
                pgStatus.Visibility = Visibility.Hidden;
            }
        }
        private void BgwCutContents_DoWork(object sender, DoWorkEventArgs e)
        {
            // sneder and parameters
            BackgroundWorker worker = (BackgroundWorker)sender;
            var varargin = (SelectedGridContentsObject)e.Argument;

            // start
            Dictionary<Tuple<int, int>, double?> selectedCells = new Dictionary<Tuple<int, int>, double?>();
            int r0 = int.MaxValue, rm = int.MinValue, c0 = int.MaxValue, cm = int.MinValue;
            double p = 0;
            foreach (FastWpfGrid.FastGridCellAddress cell in varargin.SelectedCellAddresses)
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
                    this.EditedCells.Remove(key);
                }
                else
                {
                    selectedCells[key] = null;
                }
                bgwCutContents.ReportProgress((int)(p++ / varargin.SelectedCellAddresses.Count * 50));
            }
            this.InvalidateAll();
            bgwCutContents.ReportProgress(50);

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
                    bgwCutContents.ReportProgress((int)(p++ / ((rm - r0 + 1) * (cm - c0 + 1)) * 100));
                }
                if (r < rm) sbCopy.Append("\r\n");
            }
            bgwCutContents.ReportProgress(100);
            Thread.Sleep(500);
        }
        /// <summary>
        /// เริ่มต้นธีด ทำสำเนาข้อมูลไปไว้ที่ clipboard และลบออกจากเซลล์
        /// </summary>
        /// <param name="objArray">1. Status Label Control,  2. Progress Bar Control,  3. HashSet<FastGridCellAddress> obj เป็นรายการจาก FastWpfGrid.GetSelectedModelCells()</param>
        public void StartCutContentToClipboardThread(SelectedGridContentsObject e)
        {
            try { lblStatus = e.StatusLabel; } catch { lblStatus = null; }
            try { pgStatus = e.StatusProgressBar; } catch { pgStatus = null; }
            if (lblStatus != null) lblStatus.Content = "Cutting...";
            if (!bgwCutContents.IsBusy)
            {
                pgStatus.Visibility = Visibility.Visible;
                bgwCutContents.RunWorkerAsync(e);
            }
        }
        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using FastWpfGrid;

namespace Quintessence.Ibp2018.Model
{
    public class CurrentGridModel : FastGridModelBase
    {
        /// <summary>
        /// Headers ทีเก็บหัวข้อ
        /// </summary>
        public Dictionary<Tuple<int>, string> Headers = new Dictionary<Tuple<int>, string>();

        public void DefineNewHeaders()
        {
            int i = 0;
            Headers.Clear();
            Headers[Tuple.Create(i++)] = " Student ID ";
            Headers[Tuple.Create(i++)] = "           Full Name                  ";
            Headers[Tuple.Create(i++)] = "   A1   ";
            Headers[Tuple.Create(i++)] = "   A2   ";
            Headers[Tuple.Create(i++)] = "   A3   ";
            Headers[Tuple.Create(i++)] = "   A4   ";
            Headers[Tuple.Create(i++)] = "   A1   ";
            Headers[Tuple.Create(i++)] = "   A2   ";
            Headers[Tuple.Create(i++)] = "   A3   ";
            Headers[Tuple.Create(i++)] = "   A4   ";
            Headers[Tuple.Create(i++)] = "   A1   ";
            Headers[Tuple.Create(i++)] = "   A2   ";
            Headers[Tuple.Create(i++)] = "   A3   ";
            Headers[Tuple.Create(i++)] = "   A4   ";
            Headers[Tuple.Create(i++)] = "   A1   ";
            Headers[Tuple.Create(i++)] = "   A2   ";
            Headers[Tuple.Create(i++)] = "   A3   ";
            Headers[Tuple.Create(i++)] = "   A4   ";
            Headers[Tuple.Create(i++)] = "   A1   ";
            Headers[Tuple.Create(i++)] = "   A2   ";
            Headers[Tuple.Create(i++)] = "   A3   ";
            Headers[Tuple.Create(i++)] = "   A4   ";
            Headers[Tuple.Create(i++)] = "   A1   ";
            Headers[Tuple.Create(i++)] = "   A2   ";
            Headers[Tuple.Create(i++)] = "   A3   ";
            Headers[Tuple.Create(i++)] = "   A4   ";
            Headers[Tuple.Create(i++)] = "   A1   ";
            Headers[Tuple.Create(i++)] = "   A2   ";
            Headers[Tuple.Create(i++)] = "   A3   ";
            Headers[Tuple.Create(i++)] = "   A4   ";
            Headers[Tuple.Create(i++)] = "   A1   ";
            Headers[Tuple.Create(i++)] = "   A2   ";
            Headers[Tuple.Create(i++)] = "   A3   ";
            Headers[Tuple.Create(i++)] = "   A4   ";
            _columnCount = Headers.Count;
        }

        public Dictionary<Tuple<int, int>, string> EditedCells = new Dictionary<Tuple<int, int>, string>();
 
        public double XStep { get; set; }        

        public double YStep { get; set; }

        public int XCount { get { return _columnCount; } set { _columnCount = value; } }
        private int _columnCount = 70;
        public override int ColumnCount
        {
            get { return _columnCount; }
        }

        private int _rowCount = 10;
        public override int RowCount
        {
            get { return _rowCount; }
        }
        
        public override string GetCellText(int row, int column)
        {
            var key = Tuple.Create(row, column);
            if (EditedCells.ContainsKey(key)) return EditedCells[key];

            return "";
        }

        public override void SetCellText(int row, int column, string value)
        {
            var key = Tuple.Create(row, column);
            EditedCells[key] = value;
        }

        public override void HandleSelectionCommand(IFastGridView view, string command)
        {
            MessageBox.Show(command);
        }

        public override IFastGridCell GetGridHeader(IFastGridView view)
        {
            var res = new FastGridCellImpl();
            res.Blocks.Add(new FastGridBlockImpl
            {
                IsBold = true,
                TextData = "  No.",
            });
            return res;
        }

        public override IFastGridCell GetColumnHeader(IFastGridView view, int column)
        {
            var res = new FastGridCellImpl();
            var key = Tuple.Create(column);
            res.Blocks.Add(new FastGridBlockImpl
            {
                IsBold = true,
                //TextData = string.Format("X={0:0.00}", column * 0.02),
                TextData = Headers[key],
            });
            return res;
        }

        public override IFastGridCell GetRowHeader(IFastGridView view, int row)
        {
            var res = new FastGridCellImpl();
            res.Blocks.Add(new FastGridBlockImpl
            {
                IsBold = false,
                TextData = String.Format("{0,6}", row*50+1)
            });
            return res;
        }


    }
}

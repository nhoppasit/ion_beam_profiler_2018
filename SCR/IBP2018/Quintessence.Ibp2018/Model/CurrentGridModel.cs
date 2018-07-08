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
        public Dictionary<Tuple<int, int>, string> EditedCells = new Dictionary<Tuple<int, int>, string>();

        public override int ColumnCount
        {
            get { return 70; }
        }

        public override int RowCount
        {
            get { return 1000; }
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

        public override IFastGridCell GetColumnHeader(IFastGridView view, int column)
        {
            var res = new FastGridCellImpl();
            res.Blocks.Add(new FastGridBlockImpl
            {
                IsBold = true,
                TextData = string.Format("X={0:0.00}", column * 0.02),
            });
            return res;
        }

        public override IFastGridCell GetRowHeader(IFastGridView view, int row)
        {
            var res = new FastGridCellImpl();
            res.Blocks.Add(new FastGridBlockImpl
            {
                IsBold = true,
                TextData = String.Format("Y={0:0.00}", row * 0.02)
            });
            return res;
        }

    }
}

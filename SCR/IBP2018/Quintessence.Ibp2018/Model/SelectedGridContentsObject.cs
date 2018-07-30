using FastWpfGrid;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Quintessence.Ibp2018.Model
{
    public class SelectedGridContentsObject
    {
        public Label StatusLabel { get; set; }
        public ProgressBar StatusProgressBar { get; set; }
        public HashSet<FastGridCellAddress> SelectedCellAddresses { get; set; }
    }
}

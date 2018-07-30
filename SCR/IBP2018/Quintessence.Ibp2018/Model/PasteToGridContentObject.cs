using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Quintessence.Ibp2018.Model
{
    public class PasteToGridContentObject 
    {
        public string Text { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }
        public Label StatusLabel { get; set; }
        public ProgressBar StatusProgressBar { get; set; }
    }
}

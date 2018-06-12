using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace Quintessence.Ibp2018.Model
{
    public class Ibp2018FileModel : INotifyPropertyChanged
    {
        #region INotifyPropertyChanged Implementation
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion

        /* -----------------------------------------------------
         * Properties of machine
         *    1. Motion control
         *    2. Meters
         * ----------------------------------------------------- */
        // Motion control
        private string _FilePath;
        private string _FileName;
        private string _XyMmc2PortName;
        private string _ZMmc2PortName;
        private string _XScanStep;
        private string _YScanStep;
        private string _XScanStart;
        private string _XScanEnd;
        private string _YScanStart;
        private string _YScanEnd;
        private string _ZScanPos;
        private string _XMin;
        private string _XMax;
        private string _YMin;
        private string _YMax;
        private string _ZMin;
        private string _ZMax;
        // Meters
        private string _Dmm1VisaAddress;
        private string _Dmm2VisaAddress;
        private string _ReadInterval;
        private string _AveragingNumber;

        /* -----------------------------------------------------
         * Properties of file
         *    1. Table dimension
         *    2. 3D surface configurations
         * ----------------------------------------------------- */
        private int _RowsCount;
        private int _ColumnsCount;      
    }

    public class Ibp2018RowModel : INotifyPropertyChanged
    {
        #region INotifyPropertyChanged Implementation
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion

        /* -----------------------------------------------------
         * Properties of file
         *    - Rows of current
         *    Need dynamic columns code!
         *    
         * ----------------------------------------------------- */
         
    }
}

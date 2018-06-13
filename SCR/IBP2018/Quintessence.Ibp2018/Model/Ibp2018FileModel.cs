using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Data;
using Quintessence.Meter.Gpib34401a;

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

    public class Ibp2018DataTableModel : INotifyPropertyChanged
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
         *    Need dynamic columns code in view-model code!
         *    
         * ----------------------------------------------------- */
        private DataTable dataTable = new DataTable();
        public DataTable Datatable
        {
            get
            {
                return dataTable;
            }
            set
            {
                if (dataTable != value)
                {
                    dataTable = value;
                }
                OnPropertyChanged("Datatable");
            }
        }

        public IList<string> ColumnNames;
        public IList<string> ColumnHeaders;

        public void GenerateNewDemoColumns(double xStep, double yStep, double xMin, double xMax, double yMin, double yMax)
        {
            // Number of columns
            int colCount = (int)((xMax - xMin) / xStep) + 1;
            int rowCount = (int)((yMax - yMin) / yStep) + 1;

            // Colums definetion
            ColumnNames = new List<string>();
            ColumnHeaders = new List<string>();
            dataTable.Rows.Clear();
            dataTable.Columns.Clear();
            dataTable.Columns.Add("Y_Step", typeof(string));
            ColumnNames.Add("Y_Step");
            ColumnHeaders.Add("Y Step");
            for (int i = 0; i < colCount; i++)
            {
                dataTable.Columns.Add("X_" + i.ToString(), typeof(string));
                ColumnNames.Add("X_" + i.ToString());
                ColumnHeaders.Add("X=" + (i * xStep).ToString("F2"));
            }
        }
    }
}

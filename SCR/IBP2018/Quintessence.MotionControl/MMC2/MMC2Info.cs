using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.Threading;
using System.ComponentModel;

namespace Quintessence.MotionControl.MMC2
{
    public class MMC2Info : INotifyPropertyChanged
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

        /* ----------------------------------------------------------  
         * กายภาพของมอเตอร์ 5-phase
         * ----------------------------------------------------------  */
        public const double MillimeterPerStep = 1f / 500.0;
        public const double StepPerMillimeter = 500;
        public const double StepPerDegree = 500 / 360;
        public const double DegreePerStep = 360 / 500;

        /* ----------------------------------------------------------  
         * พอร์ต
         * ----------------------------------------------------------  */
        private string _SerialPortName; public string SerialPortName { get { return _SerialPortName; } set { _SerialPortName = value; } }

        /* ----------------------------------------------------------  
         * ข้อมูลตำแหน่ง
         * ----------------------------------------------------------  */
        private int _ActualXStep; public int ActualXStep { get { return _ActualXStep; } }
        public double ActualX { get { return (double)_ActualXStep * MillimeterPerStep; } }
        private int _ActualYStep; public int ActualYStep { get { return _ActualYStep; } }
        public double ActualY { get { return (double)_ActualYStep * MillimeterPerStep; } }
        private string _SensorX; public string SensorX { get { return _SensorX; } }
        private string _SensorY; public string SensorY { get { return _SensorY; } }
        private bool _IsReady; public bool IsReady { get { return _IsReady; } }

        /* ----------------------------------------------------------  
         * การเคลื่อนที่แบบสัมพัทธ์
         * ----------------------------------------------------------  */
        private int _XRelativeStep; public int XRelativeStep { get { return _XRelativeStep; } set { _XRelativeStep = value; } }
        private int _YRelativeStep; public int YRelativeStep { get { return _YRelativeStep; } set { _YRelativeStep = value; } }

        /* ----------------------------------------------------------  
         * Resolution
         * ----------------------------------------------------------  */
        private double _XScanResolution; public double XScanStep { get { return _XScanResolution; } set { _XScanResolution = value; } }
        private double _YScanResolution; public double YScanStep { get { return _YScanResolution; } set { _YScanResolution = value; } }

        /* ----------------------------------------------------------  
         * Scan Area
         * ----------------------------------------------------------  */
        private double _XScanMinimum; public double XScanStart { get { return _XScanMinimum; } set { _XScanMinimum = value; } }
        private double _YScanMinimum; public double YScanStart { get { return _YScanMinimum; } set { _YScanMinimum = value; } }
        private double _XScanMaximum; public double XScanEnd { get { return _XScanMaximum; } set { _XScanMaximum = value; } }
        private double _YScanMaximum; public double YScanEnd { get { return _YScanMaximum; } set { _YScanMaximum = value; } }

        /* ----------------------------------------------------------  
         * Figture Range
         * ----------------------------------------------------------  */
        private double _XFigtureMinimum; public double XFigtureMinimum { get { return _XFigtureMinimum; } set { _XFigtureMinimum = value; } }
        private double _YFigtureMinimum; public double YFigtureMinimum { get { return _YFigtureMinimum; } set { _YFigtureMinimum = value; } }
        private double _XFigtureMaximum; public double XFigtureMaximum { get { return _XFigtureMaximum; } set { _XFigtureMaximum = value; } }
        private double _YFigtureMaximum; public double YFigtureMaximum { get { return _YFigtureMaximum; } set { _YFigtureMaximum = value; } }

        /* ----------------------------------------------------------  
         * Initialize serial port
         * ----------------------------------------------------------  */
        public SerialPort Port = new SerialPort();
        public PortResponse Connect()
        {
            try
            {
                if (_SerialPortName == null) return new PortResponse(PortResponse.ERR_PORTNAME, "Port name cannot be null.", null);
                if (_SerialPortName.Equals("")) return new PortResponse(PortResponse.ERR_PORTNAME, "Port name must not be empty.", null);
                if (Port == null)
                {
                    Port = new SerialPort();
                    Port.PortName = _SerialPortName;
                }
                if (Port.IsOpen) Port.Close();
                Port.PortName = _SerialPortName;
                Port.BaudRate = 9600;
                Port.DataBits = 8;
                Port.StopBits = StopBits.One;
                Port.Handshake = Handshake.None;
                Port.Parity = Parity.None;
                Port.Encoding = System.Text.Encoding.Default;
                Port.ReadTimeout = 70;
                Port.RtsEnable = false;
                Port.DtrEnable = false;
                Port.NewLine = "\n";
                Port.Open();
                Thread.Sleep(100);
                PortResponse pr = new PortResponse(PortResponse.SUCCESS, "Connected on " + _SerialPortName + " successful. ", null);
                return pr;
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse(PortResponse.ERR_OPEN, "Connect on " + _SerialPortName + " error. " + ex.Message, ex);
                return pr;
            }
        }
        public PortResponse Disconnect()
        {
            try
            {
                Port.Close();
                Thread.Sleep(100);
                PortResponse pr = new PortResponse("00", "Disconnected", null);
                return pr;
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse("DE", "Disconnected " + _SerialPortName + " error. " + ex.Message, ex);
                return pr;
            }
        }

        /* ----------------------------------------------------------  
         * Read current position and sensors
         * ----------------------------------------------------------  */
        public PortResponse QueryPosition(bool demo = false)
        {
            try
            {
                if (demo)
                {
                    _ActualXStep += 10;
                    _ActualYStep += 10;
                    _SensorX = "K";
                    _SensorY = "K";
                    _IsReady = true;
                    PortResponse pr = new PortResponse(PortResponse.DEMO, "Demo mode.", null);
                    return pr;
                }

                Port.Write("Q:\n");
                string Incomming = string.Empty;
                Incomming = Port.ReadLine();
                string[] aStr = Incomming.Split(',');
                int val;
                if (aStr == null)
                {
                    PortResponse pr = new PortResponse(PortResponse.ERR_QUERY, "Invalid existing buffer!", null);
                    return pr;
                }
                else
                {
                    if (1 <= aStr.Length)
                    {
                        if (int.TryParse(aStr[0], out val)) _ActualXStep = val;
                        else _ActualXStep = 0;
                    }
                    if (2 <= aStr.Length)
                    {
                        if (int.TryParse(aStr[1], out val)) _ActualYStep = val;
                        else _ActualYStep = 0;
                    }
                    if (3 <= aStr.Length)
                    {
                        _SensorX = aStr[2];
                    }
                    if (4 <= aStr.Length)
                    {
                        _SensorY = aStr[3];
                    }
                    if (5 <= aStr.Length)
                    {
                        if (aStr[4] == "R") _IsReady = true;
                        else _IsReady = false;
                    }
                    PortResponse pr = new PortResponse(PortResponse.SUCCESS, "", null);
                    return pr;
                }
            }
            catch (Exception ex)
            {
                _ActualXStep = _ActualYStep = 0;
                _SensorX = _SensorY = "";
                _IsReady = true;
                PortResponse pr = new PortResponse(PortResponse.ERR_QUERY, "Query on " + _SerialPortName + " error. " + ex.Message, ex);
                return pr;
            }
        }
    }
}

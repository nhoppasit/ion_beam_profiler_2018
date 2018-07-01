﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.Threading;
using System.ComponentModel;
using System.Collections.ObjectModel;

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
        public const double MILLIMETERPERSTEP = 1f / 500.0;
        public const double STEPPERMILLIMETER = 500;
        public const double STEPPERDEGREE = 500 / 360;
        public const double DEGREEPERSTEP = 360 / 500;

        /* ----------------------------------------------------------  
         * พอร์ต
         * ----------------------------------------------------------  */
        private string _SerialPortName; public string SerialPortName { get { return _SerialPortName; } set { _SerialPortName = value; } }

        private bool _IsDemo = false; public bool IsDemo { get { return _IsDemo; } set { _IsDemo = value; } }

        /* ----------------------------------------------------------  
         * ข้อมูลตำแหน่ง
         * ----------------------------------------------------------  */
        private int _ActualXStep; public int ActualXStep { get { return _ActualXStep; } }
        public double ActualX { get { return (double)_ActualXStep * MILLIMETERPERSTEP; } set { _ActualXStep = (int)(value * STEPPERMILLIMETER); } }
        private int _ActualYStep; public int ActualYStep { get { return _ActualYStep; } }
        public double ActualY { get { return (double)_ActualYStep * MILLIMETERPERSTEP; } set { _ActualYStep = (int)(value * STEPPERMILLIMETER); } }
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
        private double _XScanStep; public double XScanStep { get { return _XScanStep; } set { _XScanStep = value; } }
        private double _YScanStep; public double YScanStep { get { return _YScanStep; } set { _YScanStep = value; } }

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
        private double _XFigtureMinimum; public double XFixtureMinimum { get { return _XFigtureMinimum; } set { _XFigtureMinimum = value; } }
        private double _YFigtureMinimum; public double YFixtureMinimum { get { return _YFigtureMinimum; } set { _YFigtureMinimum = value; } }
        private double _XFigtureMaximum; public double XFixtureMaximum { get { return _XFigtureMaximum; } set { _XFigtureMaximum = value; } }
        private double _YFigtureMaximum; public double YFixtureMaximum { get { return _YFigtureMaximum; } set { _YFigtureMaximum = value; } }

        /// <summary>
        /// รายการสำหรับเลือก X scan range
        /// </summary>        
        private ObservableCollection<double> _XScanRangeList;
        public ObservableCollection<double> XScanRangeList
        {
            get
            {
                if (_XScanRangeList == null) _XScanRangeList = new ObservableCollection<double>();
                if (_XScanRangeList.Count > 0) _XScanRangeList.Clear();
                for (int i = 0; i <= (int)(_XFigtureMaximum - _XFigtureMinimum); i++)
                    _XScanRangeList.Add(i + _XFigtureMinimum);
                return _XScanRangeList;
            }
        }

        /// <summary>
        /// รายการสำหรับเลือก Y scan range
        /// </summary>
        private ObservableCollection<double> _YScanRangeList;
        public ObservableCollection<double> YScanRangeList
        {
            get
            {
                if (_YScanRangeList == null) _YScanRangeList = new ObservableCollection<double>();
                if (_YScanRangeList.Count > 0) _YScanRangeList.Clear();
                for (int i = (int)_YFigtureMinimum; i <= (int)_YFigtureMaximum; i++)
                    _YScanRangeList.Add(i);
                return _YScanRangeList;
            }
        }


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

        // Read current position and sensors
        public PortResponse QueryPosition()
        {
            try
            {
                if (_IsDemo)
                {
                    PortResponse pr = new PortResponse(PortResponse.DEMO, "Demo mode.", null);
                    return pr;
                }

                Port.DiscardOutBuffer();
                Port.DiscardInBuffer();
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
                        if (aStr[4].Contains("R")) _IsReady = true;
                        else _IsReady = false;
                    }
                    PortResponse pr = new PortResponse(PortResponse.SUCCESS, Incomming, null);
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

        // X jog
        public PortResponse JogX(int step)
        {
            try
            {
                if (_IsDemo)
                {
                    _ActualXStep += step;
                    _SensorX = "K";
                    _SensorY = "K";
                    _IsReady = true;
                    PortResponse pr0 = new PortResponse(PortResponse.DEMO, "Demo mode.", null);
                    return pr0;
                }

                // Validateion
                PortResponse pr = QueryPosition();
                if (pr.Code != PortResponse.SUCCESS) { return pr; }
                double tAPos = step * MILLIMETERPERSTEP + ActualX;
                if (tAPos < XFixtureMinimum) step = (int)((XFixtureMinimum - ActualX) * STEPPERMILLIMETER);
                if (XFixtureMaximum < tAPos) step = (int)((XFixtureMaximum - ActualX) * STEPPERMILLIMETER);

                // Communication
                string stepMessage = "M:XP" + step.ToString() + "\n";
                string goMessage = "G:\n";
                pr = new PortResponse("00", stepMessage, null);
                pr = new PortResponse("00", goMessage, null);
                string Incomming = string.Empty;
                Port.DiscardOutBuffer();
                Port.DiscardInBuffer();
                Port.Write(stepMessage);
                Incomming = Port.ReadLine();
                Port.Write(goMessage);
                Incomming = Port.ReadLine();
                for (int i = 0; i < 3; i++) if (!IsReady) { pr = QueryPosition(); Thread.Sleep(70); } else break;
                return pr;
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse(PortResponse.ERR_JOG, "Jog X on " + _SerialPortName + " error. " + ex.Message, ex);
                return pr;
            }
        }

        // Y jog
        public PortResponse JogY(int step)
        {
            try
            {
                if (_IsDemo)
                {
                    _ActualYStep += step;
                    _SensorX = "K";
                    _SensorY = "K";
                    _IsReady = true;
                    PortResponse pr0 = new PortResponse(PortResponse.DEMO, "Demo mode.", null);
                    return pr0;
                }

                // Validateion
                PortResponse pr = QueryPosition();
                if (pr.Code != PortResponse.SUCCESS) { return pr; }
                double tAPos = step * MILLIMETERPERSTEP + ActualX;
                if (tAPos < YFixtureMinimum) step = (int)((YFixtureMinimum - ActualY) * STEPPERMILLIMETER);
                if (YFixtureMaximum < tAPos) step = (int)((YFixtureMaximum - ActualY) * STEPPERMILLIMETER);

                // Communication
                string stepMessage = "M:YP" + step.ToString() + "\n";
                string goMessage = "G:\n";
                pr = new PortResponse("00", stepMessage, null);
                pr = new PortResponse("00", goMessage, null);
                string Incomming = string.Empty;
                Port.DiscardOutBuffer();
                Port.DiscardInBuffer();
                Port.Write(stepMessage);
                Incomming = Port.ReadLine();
                Port.Write(goMessage);
                Incomming = Port.ReadLine();
                for (int i = 0; i < 3; i++) if (!IsReady) { pr = QueryPosition(); Thread.Sleep(70); } else break;
                return pr;
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse(PortResponse.ERR_JOG, "Jog Y on " + _SerialPortName + " error. " + ex.Message, ex);
                return pr;
            }
        }

        // Absolute move of X axis
        public PortResponse AbsoluteMoveX()
        {
            try
            {
                if (_IsDemo)
                {
                    _SensorX = "K";
                    _SensorY = "K";
                    _IsReady = true;
                    PortResponse pr0 = new PortResponse(PortResponse.DEMO, "Demo mode.", null);
                    return pr0;
                }

                // Validateion with member
                int step = (int)(ActualX * STEPPERMILLIMETER);
                if (ActualX < XFixtureMinimum) step = (int)((XFixtureMinimum) * STEPPERMILLIMETER);
                if (XFixtureMaximum < ActualX) step = (int)((XFixtureMaximum) * STEPPERMILLIMETER);

                // Communication                
                string stepMessage = "A:XP" + step.ToString() + "\n";
                PortResponse pr = new PortResponse("00", stepMessage, null);
                string Incomming = string.Empty;
                Port.DiscardOutBuffer();
                Port.DiscardInBuffer();
                Port.Write(stepMessage);
                Incomming = Port.ReadLine();
                for (int i = 0; i < 3; i++) if (!IsReady) { pr = QueryPosition(); Thread.Sleep(70); } else break;
                return pr;
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse(PortResponse.ERR_ABS, "Absolute move of X on " + _SerialPortName + " error. " + ex.Message, ex);
                return pr;
            }
        }

        // Absolute move of Y axis
        public PortResponse AbsoluteMoveY()
        {
            try
            {
                if (_IsDemo)
                {
                    _SensorX = "K";
                    _SensorY = "K";
                    _IsReady = true;
                    PortResponse pr0 = new PortResponse(PortResponse.DEMO, "Demo mode.", null);
                    return pr0;
                }

                // Validateion with member
                int step = (int)(ActualY * STEPPERMILLIMETER);
                if (ActualY < YFixtureMinimum) step = (int)((YFixtureMinimum) * STEPPERMILLIMETER);
                if (YFixtureMaximum < ActualY) step = (int)((YFixtureMaximum) * STEPPERMILLIMETER);

                // Communication                
                string stepMessage = "A:YP" + step.ToString() + "\n";
                PortResponse pr = new PortResponse("00", stepMessage, null);
                string Incomming = string.Empty;
                Port.DiscardOutBuffer();
                Port.DiscardInBuffer();
                Port.Write(stepMessage);
                Incomming = Port.ReadLine();
                for (int i = 0; i < 3; i++) if (!IsReady) { pr = QueryPosition(); Thread.Sleep(70); } else break;
                return pr;
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse(PortResponse.ERR_ABS, "Absolute move of Y on " + _SerialPortName + " error. " + ex.Message, ex);
                return pr;
            }
        }

        // Goto Zero reference of X axis
        public PortResponse GotoZeroX()
        {
            try
            {
                if (_IsDemo)
                {
                    ActualX = 0;
                    _SensorX = "K";
                    _SensorY = "K";
                    _IsReady = true;
                    PortResponse pr0 = new PortResponse(PortResponse.DEMO, "Demo mode.", null);
                    return pr0;
                }

                // Communication                
                string stepMessage = "A:XP0\n";
                PortResponse pr = new PortResponse("00", stepMessage, null);
                string Incomming = string.Empty;
                Port.DiscardOutBuffer();
                Port.DiscardInBuffer();
                Port.Write(stepMessage);
                Incomming = Port.ReadLine();
                for (int i = 0; i < 3; i++) if (!IsReady) { pr = QueryPosition(); Thread.Sleep(70); } else break;
                return pr;
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse(PortResponse.ERR_ZERO, "Zero reference of X on " + _SerialPortName + " error. " + ex.Message, ex);
                return pr;
            }
        }

        // Goto Zero reference of Y axis
        public PortResponse GotoZeroY()
        {
            try
            {
                if (_IsDemo)
                {
                    ActualY = 0;
                    _SensorX = "K";
                    _SensorY = "K";
                    _IsReady = true;
                    PortResponse pr0 = new PortResponse(PortResponse.DEMO, "Demo mode.", null);
                    return pr0;
                }

                // Communication                
                string stepMessage = "A:YP0\n";
                PortResponse pr = new PortResponse("00", stepMessage, null);
                string Incomming = string.Empty;
                Port.DiscardOutBuffer();
                Port.DiscardInBuffer();
                Port.Write(stepMessage);
                Incomming = Port.ReadLine();
                for (int i = 0; i < 3; i++) if (!IsReady) { pr = QueryPosition(); Thread.Sleep(70); } else break;
                return pr;
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse(PortResponse.ERR_ZERO, "Zero reference of Y on " + _SerialPortName + " error. " + ex.Message, ex);
                return pr;
            }
        }

        /// <summary>
        /// Set as zero reference of X axis
        /// </summary>
        /// <returns>Port Response class</returns>
        public PortResponse SetZeroX()
        {
            try
            {
                if (_IsDemo)
                {
                    ActualY = 0;
                    _SensorX = "K";
                    _SensorY = "K";
                    _IsReady = true;
                    PortResponse pr0 = new PortResponse(PortResponse.DEMO, "Demo mode.", null);
                    return pr0;
                }

                // Communication                
                string stepMessage = "P:14P0\n";
                PortResponse pr = new PortResponse("00", stepMessage, null);
                string Incomming = string.Empty;
                Port.DiscardOutBuffer();
                Port.DiscardInBuffer();
                Port.Write(stepMessage);
                Incomming = Port.ReadLine();
                for (int i = 0; i < 3; i++) if (!IsReady) { pr = QueryPosition(); Thread.Sleep(70); } else break;
                return pr;
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse(PortResponse.ERR_REFZ, "Set as zero reference of X on " + _SerialPortName + " error. " + ex.Message, ex);
                return pr;
            }
        }

        /// <summary>
        /// Set as zero reference of X axis 
        /// </summary>
        /// <returns>Port Response class</returns>
        public PortResponse SetZeroY()
        {
            try
            {
                if (_IsDemo)
                {
                    ActualY = 0;
                    _SensorX = "K";
                    _SensorY = "K";
                    _IsReady = true;
                    PortResponse pr0 = new PortResponse(PortResponse.DEMO, "Demo mode.", null);
                    return pr0;
                }

                // Communication                
                string stepMessage = "P:15P0\n";
                PortResponse pr = new PortResponse("00", stepMessage, null);
                string Incomming = string.Empty;
                Port.DiscardOutBuffer();
                Port.DiscardInBuffer();
                Port.Write(stepMessage);
                Incomming = Port.ReadLine();
                for (int i = 0; i < 3; i++) if (!IsReady) { pr = QueryPosition(); Thread.Sleep(70); } else break;
                return pr;
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse(PortResponse.ERR_REFZ, "Set as zero reference of Y on " + _SerialPortName + " error. " + ex.Message, ex);
                return pr;
            }
        }



    }
}

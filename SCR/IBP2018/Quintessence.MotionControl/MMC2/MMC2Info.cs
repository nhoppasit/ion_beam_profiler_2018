using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.Threading;

namespace Quintessence.MotionControl.MMC2
{
    public class MMC2Info
    {
        /* ----------------------------------------------------------  
         * กายภาพของมอเตอร์ 5-phase
         * ----------------------------------------------------------  */
        public const float MillimeterPerStep = 1f / 500.0f;
        public const float StepPerMillimeter = 500f;
        public const float StepPerDegree = 500f / 360f;
        public const float DegreePerStep = 360f / 500f;

        /* ----------------------------------------------------------  
         * พอร์ต
         * ----------------------------------------------------------  */
        private string _SerialPortName; public string SerialPortName { get { return _SerialPortName; } set { _SerialPortName = value; } }

        /* ----------------------------------------------------------  
         * ข้อมูลตำแหน่ง
         * ----------------------------------------------------------  */
        private int _ActualXStep; public int ActualXStep { get { return _ActualXStep; } set { _ActualXStep = value; } }
        private int _ActualYStep; public int ActualYStep { get { return _ActualYStep; } set { _ActualYStep = value; } }
        private bool _IsReady; public bool IsReady { get { return _IsReady; } set { _IsReady = value; } }

        /* ----------------------------------------------------------  
         * การเคลื่อนที่แบบสัมพัทธ์
         * ----------------------------------------------------------  */
        private int _XRelativeStep; public int XRelativeStep { get { return _XRelativeStep; } set { _XRelativeStep = value; } }
        private int _YRelativeStep; public int YRelativeStep { get { return _YRelativeStep; } set { _YRelativeStep = value; } }

        /* ----------------------------------------------------------  
         * Resolution
         * ----------------------------------------------------------  */
        private float _XScanResolution; public float XScanResolution { get { return _XScanResolution; } set { _XScanResolution = value; } }
        private float _YScanResolution; public float YScanResolution { get { return _YScanResolution; } set { _YScanResolution = value; } }

        /* ----------------------------------------------------------  
         * Scan Area
         * ----------------------------------------------------------  */
        private float _XScanMinimum; public float XScanMinimum { get { return _XScanMinimum; } set { _XScanMinimum = value; } }
        private float _YScanMinimum; public float YScanMinimum { get { return _YScanMinimum; } set { _YScanMinimum = value; } }
        private float _XScanMaximum; public float XScanMaximum { get { return _XScanMaximum; } set { _XScanMaximum = value; } }
        private float _YScanMaximum; public float YScanMaximum { get { return _YScanMaximum; } set { _YScanMaximum = value; } }

        /* ----------------------------------------------------------  
         * Figture Range
         * ----------------------------------------------------------  */
        private float _XFigtureMinimum; public float XFigtureMinimum { get { return _XFigtureMinimum; } set { _XFigtureMinimum = value; } }
        private float _YFigtureMinimum; public float YFigtureMinimum { get { return _YFigtureMinimum; } set { _YFigtureMinimum = value; } }
        private float _XFigtureMaximum; public float XFigtureMaximum { get { return _XFigtureMaximum; } set { _XFigtureMaximum = value; } }
        private float _YFigtureMaximum; public float YFigtureMaximum { get { return _YFigtureMaximum; } set { _YFigtureMaximum = value; } }

        /* ----------------------------------------------------------  
         * Initialize serial port
         * ----------------------------------------------------------  */
        public SerialPort Port = new SerialPort();
        public PortResponse Connect()
        {
            try
            {
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
                PortResponse pr = new PortResponse("00", "Connected", null);
                return pr;
            }
            catch (Exception ex)
            {
                PortResponse pr = new PortResponse("CE", "Connected " + _SerialPortName + " error. " + ex.Message, ex);
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
    }
}

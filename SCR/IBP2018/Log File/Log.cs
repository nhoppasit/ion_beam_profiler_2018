﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LogFile
{
    public class Log
    {
        private static object ThisLock = new object();

        private string _filePath = @"C:\Log";
        private string _fileName = "Log";
        public Log(string filePath = @"C:\Log", string fileName = "Log")
        {
            if (filePath != null) _filePath = filePath;
            if (fileName != null) _fileName = fileName;
        }

        private string FullFileName()
        {
            return _filePath + @"\" + _fileName + " " + DateTime.Now.ToString("yyyyMMdd") + ".txt";
        }
        private void CheckPath()
        {
            lock (ThisLock)
            {
                try
                {
                    if (!System.IO.Directory.Exists(_filePath))
                    {
                        System.IO.Directory.CreateDirectory(_filePath);
                    }
                    if (!System.IO.File.Exists(FullFileName()))
                    {
                        System.IO.File.Create(FullFileName()).Close();
                        int loop = 0;
                        while (!System.IO.File.Exists(FullFileName()) & loop < 30) { loop++; }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
            }
        }
        public void AppendText(string text)
        {
            lock (ThisLock)
            {
                try
                {
                    CheckPath();
                    using (System.IO.StreamWriter w = new System.IO.StreamWriter(FullFileName(), true))
                    {
                        string str;
                        str = DateTime.Now.ToString("yyyy/MM/dd, HH:mm:ss") + " : " + text;
                        w.WriteLine(str);
                        w.Close();
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
            }
        }

        public string GetAllTextFromCurrentFile()
        {
            return System.IO.File.ReadAllText(FullFileName());
        }
    }
}

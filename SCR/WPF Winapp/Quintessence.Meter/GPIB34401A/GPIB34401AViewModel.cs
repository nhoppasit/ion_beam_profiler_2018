using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Collections.ObjectModel;

namespace Quintessence.Meter.Gpib34401a
{
    public class Gpib34401aViewModel 
    {        
        // Model
        Gpib34401aInfo _Gpig34401a;
        public Gpib34401aInfo Gpig34401a { get { return _Gpig34401a; } set { _Gpig34401a = value; } }
        
        public Gpib34401aViewModel()
        {
            _Gpig34401a = new Gpib34401aInfo();
        }

    }
}

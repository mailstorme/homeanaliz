using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Подсчет_начислений
{
    public class tarifInfo
    {
        public string tarif;
        public int count;
        public int maximalka;
        private int _goodCount;
        public int goodCount
        {
            get { return _goodCount; }
            set { _goodCount = value; count++; }
        }

        public tarifInfo(string tariff)
        {
            maximalka = 0;
            tarif = tariff;
            this.count = 0;
            this._goodCount = 0;

        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Подсчет_начислений
{
    class TariffList
    {
        public string[] tariffs;
        int between;

        public TariffList(string[] ab,string[] reg)
        {
            tariffs = ab.Union(reg).ToArray();
            between = ab.Length;
        }
    }
}

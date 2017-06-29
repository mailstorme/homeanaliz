using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Подсчет_начислений
{
    class UserCount
    {
        public int first;
        public int second;
        public int third;
        public int m46;
        public int m79;

        public int gfirst;
        public int gsecond;
        public int gthird;
        public int gm46;
        public int gm79;

        public void AddBadUser(int n)
        {
            if (n < 4)
                switch (n)
                {
                    case 1:
                        first++;
                        break;
                    case 2:
                        second++;
                        break;
                    case 3:
                        third++;
                        break;
                }
            if (n > 3 && n < 7)
                m46++;
            if (n > 6)
                m79++;
        }

        public void AddGoodUser(int n)
        {
            if (n < 4)
                switch (n)
                {
                    case 1:
                        gfirst++;
                        first++;
                        break;
                    case 2:
                        gsecond++;
                        second++;
                        break;
                    case 3:
                        gthird++;
                        third++;
                        break;
                }
            if (n > 3 && n < 7)
            {
                gm46++;
                m46++;
            }
            if (n > 6)
            {
                gm79++;
                m79++;
            }
        }

        public int AllUsers(int n)
        {
            if (n < 4)
                switch (n)
                {
                    case 1:
                        return first;
                    case 2:
                        return second;
                    case 3:
                        return third;
                }
            if (n > 3 && n < 7)
                return m46;
            if (n > 6)
                return m79;
            return 0;
        }

        public int AllGoodUsers(int n)
        {
            if (n < 4)
                switch (n)
                {
                    case 1:
                        return gfirst;
                    case 2:
                        return gsecond;
                    case 3:
                        return gthird;
                }
            if (n > 3 && n < 7)
                return gm46;
            if (n > 6)
                return gm79;
            return 0;
        }
    }
}

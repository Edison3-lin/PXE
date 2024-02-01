using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaptainWin.CommonAPI
{
    public class MousePosition
    {
        public int X { get; set; }
        public int Y { get; set; }

        public void adjustByRatio(float xRatio, float yRatio)
        {
            X = Convert.ToInt32(X / xRatio);
            Y = Convert.ToInt32(Y / yRatio);
        }

    }
}

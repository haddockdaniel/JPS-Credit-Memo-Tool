using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JurisUtilityBase
{
    public class CreditMemo
    {
        public int inv { get; set; }
        public int mat { get; set; }
        public double fees { get; set; }
        public double cashexp { get; set; }
        public double noncashexp { get; set; }
        public int LHID { get; set; }
        public int BatchNumber { get; set; }
    }
}

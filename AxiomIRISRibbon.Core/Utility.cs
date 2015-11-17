using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AxiomIRISRibbon.Core
{
    public class Utility
    {
       
        public static bool IsNull(DateTime d)
        {
            return (d == DateTime.MinValue || d == Constants.SQL_NULL_DT    );
        }
    }
}

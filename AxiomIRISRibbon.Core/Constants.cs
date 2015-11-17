using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AxiomIRISRibbon.Core
{
    public class Constants
    {
        public static readonly string LOG_PATH = AppDomain.CurrentDomain.BaseDirectory + ConfigurationSettings.AppSettings["LOG_FILE_PATH"];
        public static readonly DateTime SQL_NULL_DT = new DateTime(1900, 01, 01, 00, 00, 00);
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace El2Core.Constants
{
    /// <summary>
    /// Represents the different types of shifts available in the system.
    /// </summary>
    public class ShiftTypes
    {
        public ShiftTypes() { }
        public enum ShiftType : int
        {
            [Description("Start So 1")]
            StartSun1 = 1,
            [Description("Start So 2")]
            StartSun2 = 2,
            [Description("Start So 3")]
            StartSun3 = 3,
            [Description("Start Mo 1")]
            StartMon1 = 11,
            [Description("Start Mo 2")]
            StartMon2 = 12,
            [Description("Start Mo 3")]
            StartMon3 = 13,
            [Description("Ausfall")]
            DropOut = 100
        }
        public static IEnumerable<ShiftType> GetEnumTypes => Enum.GetValues(typeof(ShiftType)).Cast<ShiftType>();
    }
    
}

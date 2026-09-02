using System.ComponentModel;

namespace El2Core.Constants
{
    public class ProjectTypes
    {
        /// <summary>
        /// Defines the types of projects available in the system.
        /// </summary>
        public enum ProjectType
        {
            [Browsable(true)]
            [Description("kein Projekttyp")]
            None = 0,
            [Browsable(true)]
            [Description("Entwicklungsmuster")]
            DevelopeSpecimen = 1,
            [Browsable(true)]
            [Description("Verkaufsmuster")]
            SaleSpecimen = 2,
            [Browsable(true)]
            [Description("Versuchsproject")]
            TestOrder = 3,
            [Browsable (false)]
            [Description("_keine")]
            Null = int.MaxValue
        }
    }
}

using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations.Schema;

namespace El2Core.Constants
{
    public class ProjectTypes
    {
        public enum ProjectType
        {
            [Description("kein Projekttyp")]
            None = 0,
            [Description("Entwicklungsmuster")]
            DevelopeSpecimen = 1,
            [Description("Verkaufsmuster")]
            SaleSpecimen = 2,
            [Description("Versuchsproject")]
            TestOrder = 3,
            [NotMapped]
            [Description("_keine")]
            Null = int.MaxValue
        }
    }
}

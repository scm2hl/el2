using Microsoft.VisualStudio.TestTools.UnitTesting;
using ModulePlanning.Specials;
using System.Linq;
using System.Runtime.Versioning;

namespace LieferlisteUnitTest
{
    [TestClass]
    [SupportedOSPlatform("windows10.0.19041.0")]
    public class PlanningColumnNamesTests
    {
        [TestMethod]
        public void ColumnNames_ContainsExpectedPlanningColumns()
        {
            var columns = Constances.TLColumn.ColumnNames;

            Assert.AreEqual(5, columns.Count);
            CollectionAssert.AreEquivalent(
                new[] { "Auftragsnummer", "Vorgang", "TypTeileNummer", "Materialbezeichnung", "Kurztext" },
                columns.Keys.ToArray());
        }

        [TestMethod]
        public void ColumnNames_MapsEachColumnToItsDisplayAndPropertyNames()
        {
            var columns = Constances.TLColumn.ColumnNames;

            Assert.AreEqual(("Auftrag", "Aid", "", ""), columns["Auftragsnummer"]);
            Assert.AreEqual(("Vorg", "Vnr", "", "{0:d4}"), columns["Vorgang"]);
            Assert.AreEqual(("Material", "AidNavigation.Material", "El2Core.Converters.TTNR_Converter, El2Core, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null", ""), columns["TypTeileNummer"]);
            Assert.AreEqual(("Bezeichnung", "AidNavigation.MaterialNavigation.Bezeichng", "", ""), columns["Materialbezeichnung"]);
            Assert.AreEqual(("Kurztext", "Text", "", ""), columns["Kurztext"]);
        }
    }
}

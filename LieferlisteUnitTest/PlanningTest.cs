using El2Core.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ModulePlanning.Planning;
using ModulePlanning.Specials;
using System.Linq;
using System.Runtime.Versioning;

namespace LieferlisteUnitTest
{
    [TestClass]
    [SupportedOSPlatform("windows10.0.19041.0")]
    public class PlanningTest
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
        [TestMethod]
        public void GetFastCopyValues_ReturnsTrimmedValuesInClipboardOrder()
        {
            var vorgang = new Vorgang
            {
                Aid = "  AID-100  ",
                Vnr = 7,
                AidNavigation = new OrderRb
                {
                    Quantity = 12,
                    Material = " MAT-200 ",
                    MaterialNavigation = new TblMaterial
                    {
                        Bezeichng = " Part description "
                    }
                }
            };

            var values = PlanMachine.GetFastCopyValues(vorgang);

            CollectionAssert.AreEqual(
                new[] { "12", "Part description", "MAT-200", "7", "AID-100" },
                values);
        }

        [TestMethod]
        public void GetFastCopyValues_UsesAidAndDummyFallbacksWhenMaterialDataIsMissing()
        {
            var vorgang = new Vorgang
            {
                Aid = " AID-300 ",
                Vnr = 3,
                AidNavigation = new OrderRb()
            };

            var values = PlanMachine.GetFastCopyValues(vorgang);

            CollectionAssert.AreEqual(
                new[] { "", "AID-300", "DUMMY", "3", "AID-300" },
                values);
        }
    }
}

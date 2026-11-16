using El2Core.Constants;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ModuleDeliverList.ViewModels;
using System;
using System.Runtime.Versioning;

namespace LieferlisteUnitTest
{
    [TestClass]
    [SupportedOSPlatform("windows10.0.19041.0")]
    public class ProjectStructTests
    {
        [TestMethod]
        public void CompareTo_ComparesProjectPspInAscendingOrder()
        {
            var first = new ProjectStruct("A-100", ProjectTypes.ProjectType.Null, "first");
            var second = new ProjectStruct("B-100", ProjectTypes.ProjectType.Null, "second");

            Assert.IsTrue(first.CompareTo(second) < 0);
            Assert.IsTrue(second.CompareTo(first) > 0);
        }

        [TestMethod]
        public void CompareTo_ReturnsZeroWhenProjectPspMatches()
        {
            var first = new ProjectStruct("A-100", ProjectTypes.ProjectType.Null, "first");
            var second = new ProjectStruct("A-100", ProjectTypes.ProjectType.SaleSpecimen, "second");

            Assert.AreEqual(0, first.CompareTo(second));
        }

        [TestMethod]
        public void CompareTo_ReturnsOneWhenOtherObjectIsNull()
        {
            var project = new ProjectStruct("A-100", ProjectTypes.ProjectType.Null, null);

            Assert.AreEqual(1, project.CompareTo(null));
        }

        [TestMethod]
        public void CompareTo_ThrowsForObjectOfAnotherType()
        {
            var project = new ProjectStruct("A-100", ProjectTypes.ProjectType.Null, null);

            Assert.ThrowsExactly<ArgumentException>(() => project.CompareTo("A-100"));
        }
    }
}

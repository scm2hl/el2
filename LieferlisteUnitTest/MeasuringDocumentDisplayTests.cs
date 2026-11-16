using Microsoft.VisualStudio.TestTools.UnitTesting;
using ModuleMeasuring.ViewModels;
using System.Runtime.Versioning;

namespace LieferlisteUnitTest
{
    [TestClass]
    [SupportedOSPlatform("windows10.0.19041.0")]
    public class MeasuringDocumentDisplayTests
    {
        [TestMethod]
        public void Constructor_SetsFullNameAndDisplay()
        {
            var document = new MeasuringDocumentsViewModel.DocumentDisplay(
                @"C:\Messungen\report.pdf",
                "Messbericht.pdf");

            Assert.AreEqual(@"C:\Messungen\report.pdf", document.FullName);
            Assert.AreEqual("Messbericht.pdf", document.Display);
        }

        [TestMethod]
        public void Constructor_ClassifiesDisplayContainingGutteil()
        {
            var document = new MeasuringDocumentsViewModel.DocumentDisplay(
                @"C:\Messungen\Gutteil.pdf",
                "Messung Gutteil.pdf");

            Assert.AreEqual(1, document.DocumentType);
        }

        [TestMethod]
        public void Constructor_UsesDefaultTypeWhenDisplayDoesNotContainGutteil()
        {
            var document = new MeasuringDocumentsViewModel.DocumentDisplay(
                @"C:\Messungen\Ausschuss.pdf",
                "Ausschuss.pdf");

            Assert.AreEqual(0, document.DocumentType);
        }
    }
}

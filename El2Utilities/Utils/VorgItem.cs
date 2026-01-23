using El2Core.Models;
using System.Windows.Media.Media3D;

namespace El2Core.Utils
{
    public class VorgItem
    {
        public string Auftrag { get; private set; }
        public string Vorgang { get; private set; }
        public string? Kurztext { get; private set; }
        public string? Material { get; private set; }
        public string? Bezeichnung { get; private set; }
        public string VorgangId { get; private set; }
        public Vorgang? SourceVorgang { get; private set; }
        public VorgItem(Vorgang vorgang)
        {
            Auftrag = vorgang.Aid;
            Vorgang = vorgang.Vnr.ToString("D4");
            Kurztext = vorgang.Text;
            Material = vorgang.AidNavigation.Material ?? vorgang.AidNavigation.DummyMat;
            Bezeichnung = vorgang.AidNavigation.MaterialNavigation?.Bezeichng ??
                vorgang.AidNavigation.DummyMatNavigation?.Mattext;
            VorgangId = vorgang.VorgangId;
            SourceVorgang = vorgang;
        }
        public VorgItem(ViewVorgangClosedDate view)
        {
            Auftrag = view.Aid;
            Vorgang = view.Vnr.ToString("D4");
            Kurztext = view.Text;
            Material= view.Material ?? view.DummyMat;
            Bezeichnung = view.Bezeichng ?? view.Mattext;
            VorgangId= view.VorgangId;
        }
    }

}

using System.Collections.Generic;
using System.Collections.Immutable;

namespace ModulePlanning.Specials
{
    public static class Constances
    {
        public static class TLColumn
        {
            public static ImmutableDictionary<string, (string, string, string, string)> ColumnNames { get; } =
            new Dictionary<string, (string, string, string, string)>
            {
                { "Auftragsnummer", ("Auftrag", "Aid","","") },
                { "Vorgang", ("Vorg", "Vnr","","{0:d4}") },
                { "TypTeileNummer", ("Material", "AidNavigation.Material","El2Core.Converters.TTNR_Converter, El2Core, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null","") },
                { "Materialbezeichnung", ("Bezeichnung", "AidNavigation.MaterialNavigation.Bezeichng","","") },
                { "Kurztext", ("Kurztext", "Text","","") }
            }.ToImmutableDictionary();


        }
        

    }
    
}

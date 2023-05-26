using System;
using System.Collections.Generic;
using System.Text;

namespace ImportPlatnosci
{
    public class ListXML
    {
        public string ID { get; set; }
        public string Data { get; set; }
        public string Kwota { get; set; }
        public string Prowizja { get; set; }
        public string Wyplata { get; set; }
        public string Opis { get; set; }
        public string Kupujacy { get; set; }
        public string Numer_zamowienia { get; set; }

        public ListXML(string id, string data, string kwota, string prowizja, string wyplata, string opis, string kupujacy, string numer_zamowienia)
        {
            this.ID = id;
            this.Data = data;
            this.Kwota = kwota;
            this.Prowizja = prowizja;
            this.Wyplata = wyplata;
            this.Opis = opis;
            this.Kupujacy = kupujacy;
            this.Numer_zamowienia = numer_zamowienia;
        }
    }
}

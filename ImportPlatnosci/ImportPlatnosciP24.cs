using Soneta.Business;
using Soneta.Business.UI;
using Soneta.CRM;
using Soneta.Handel;
using Soneta.Kasa;
using Soneta.Magazyny;
using Soneta.Types;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

[assembly: Worker(typeof(ImportPlatnosci.ImportPlatnosciP24), typeof(Soneta.Kasa.RaportESP))]

namespace ImportPlatnosci
{
    public class ImportPlatnosciP24
    {
        [Context, Required]
        public NamedStream XMLFileName { get; set; }

        RaportESP raport;
        [Context]
        public RaportESP Raport
        {
            get { return raport; }
            set { raport = value; }
        }

        [Action(
            "Import transakcji Przelewy24",
            Description = "Funkcja importuje transakcje Przelewy24",
            Priority = 1000,
            Icon = ActionIcon.Wizard,
            Mode = ActionMode.SingleSession,
            Target = ActionTarget.Menu)]

        public MessageBoxInformation Funkcja()
        {
            ArrayList tablica = new ArrayList();
            string numer = "";
            string operacja = "";
            string data = "";
            string kwota = "";
            string prowizja = "";
            string id = "";
            string opis = "";
            string kupujacy = "";

            List<ListXML> lxml = new List<ListXML>();
            IList<OError> errorsList = new List<OError>();

            StreamReader objReader = new StreamReader(XMLFileName.FileName);
            string linia = "";
            int temp = 0;
            while (linia != null)
            {
                linia = objReader.ReadLine();

                if (linia != null && temp != 0 && linia != "")
                {
                    if (linia.Contains("26\"\",28")) //obsługa nadmiarowych przecinków
                        linia = linia.Replace("26\"\",28", "26\"\"-28");

                    var kawalki = Regex.Split(linia, "\",");
                    id = kawalki[2].ToString().Replace("\"", "");
                    data = kawalki[0].Replace("\"", "");
                    kwota = kawalki[8].Replace("\"", "").Split('z')[0].Replace(".", ",").Trim();
                    operacja = kawalki[3].Replace("\"", "");
                    kupujacy = kawalki[5].Replace("\"", "");
                    lxml.Add(new ListXML(id, data, kwota, "", "", operacja, kupujacy, ""));


                }

                temp++;
            }
            objReader.Close();


            //  string test = "";
            //  for (int i = 1; i < wyplata.Count; i++) test = test + wyplata[i] + "|";
            //  MessageBox.Show(test);


            using (Session session = raport.Session.Login.CreateSession(false, false))
            {
                KasaModule km = KasaModule.GetInstance(session);
                CRMModule cm = CRMModule.GetInstance(session);
                HandelModule hm = HandelModule.GetInstance(session);
                MagazynyModule mm = MagazynyModule.GetInstance(session);

                EwidencjaSP kasa = raport.Kasa;
                FromTo okres = raport.Okres;
                Date data1 = raport.Data;
                RaportESP rap = km.RaportyESP[raport.ID];

                Magazyn magazyn_wysylkowy = mm.Magazyny.WgSymbol["MG"];

                DefDokHandlowego def = hm.DefDokHandlowych.WgSymbolu["FF"];
                DefDokHandlowego def2 = hm.DefDokHandlowych.WgSymbolu["KS"];


                using (ITransaction t = session.Logout(true))
                {

                    foreach (ListXML lx in lxml)
                    {
                        var view = hm.DokHandlowe.WgMagazyn[magazyn_wysylkowy];
                        RowCondition condition = new FieldCondition.In("Features.ID_Zamowienia", lx.ID);
                        view = view[condition];

                        if (lx.Opis == "wpłata")
                        {
                            // raport validation
                            bool skipPayment = false;
                            foreach (var wr in raport.Zaplaty)
                            {
                                DokumentHandlowy temp_dok = hm.DokHandlowe.WgNumer[wr.NumeryDokumentow];

                                if (temp_dok != null)
                                    if (wr.Kwota.Value.ToString() == lx.Kwota && temp_dok.Features["ID_Zamowienia"].ToString() == lx.ID)
                                        skipPayment = true;
                            }
                            if (skipPayment)
                                continue;


                            WplataRaport dok = new WplataRaport(rap);
                            km.Zaplaty.AddRow(dok);
                            dok.Podmiot = cm.Kontrahenci.WgKodu["P24"];
                            dok.Kwota = new Currency(System.Convert.ToDouble(lx.Kwota), "PLN");
                            dok.Opis = lx.Kupujacy;
                            string kontrahent = lx.Kupujacy.Split(';')[1];
                            string nick_allegro = lx.Kupujacy.Split(';')[0];
                            string kontrahentTemp = string.Empty;
                            if (kontrahent.Contains(" "))
                            {
                                kontrahentTemp = kontrahent.Split(' ')[1] + " " + kontrahent.Split(' ')[0];
                            }
                            else
                            {
                                kontrahentTemp = kontrahent;
                            }

                            DokumentHandlowy commercialDocument = null;
                            foreach (DokumentHandlowy dok_fak in view)
                            {
                                if (dok_fak.Features["ID_Zamowienia"].ToString() == lx.ID && dok_fak.Definicja == def)
                                {
                                    dok.NumeryDokumentow = dok_fak.Numer.NumerPelny;
                                    commercialDocument = dok_fak;
                                    break;
                                }
                            }

                            if (commercialDocument == null)
                            {
                                errorsList.Add(new OError(lx.ID, "Brak dokumentu", ""));
                                continue;
                            }

                            try
                            {
                                // rozliczenie
                                SubTable st = km.RozrachunkiIdx.WgPodmiot[commercialDocument.Platnosci.GetFirst().Podmiot, Date.MaxValue];

                                Wplata paymentInRaport = null;
                                Naleznosc due = null;
                                foreach (RozrachunekIdx idx in st)
                                {
                                    if (idx.Typ == TypRozrachunku.Wpłata && paymentInRaport == null)
                                        paymentInRaport = (Wplata)idx.Dokument;
                                    if (idx.Typ == TypRozrachunku.Należność && due == null && !idx.Dokument.Bufor && idx.Numer == commercialDocument.Numer.Pelny)
                                        due = (Naleznosc)idx.Dokument;
                                    if (paymentInRaport != null && due != null)
                                        break;
                                }


                                paymentInRaport.DataDokumentu = rap.Okres.From;
                                using (ITransaction trans = session.Logout(true))
                                {
                                    RozliczenieSP settlement = new RozliczenieSP(due, (Wplata)dok);
                                    km.RozliczeniaSP.AddRow(settlement);
                                    trans.Commit();
                                }
                            }
                            catch { }

                        }
                        else if (lx.Opis == "zwrot")
                        {
                            // raport validation
                            bool skipPayment = false;
                            foreach (var wr in raport.Zaplaty)
                            {
                                DokumentHandlowy temp_dok = hm.DokHandlowe.WgNumer[wr.NumeryDokumentow];

                                if (temp_dok != null)
                                    if (((-1) * wr.Kwota.Value).ToString() == lx.Kwota && temp_dok.Features["ID_Zamowienia"].ToString() == lx.ID)
                                        skipPayment = true;
                            }
                            if (skipPayment)
                                continue;

                            WyplataRaport dokWyplata = new WyplataRaport(rap);
                            km.Zaplaty.AddRow(dokWyplata);
                            dokWyplata.Podmiot = cm.Kontrahenci.WgKodu["P24"];
                            dokWyplata.Kwota = new Currency(System.Convert.ToDouble(lx.Kwota), "PLN") * (-1);
                            dokWyplata.Opis = "Zwrot Allegro " + lx.Kupujacy;
                            dokWyplata.SposobZaplaty = km.SposobyZaplaty.WgNazwy["Przelew"];

                            DokumentHandlowy commercialDocument = null;

                            foreach (DokumentHandlowy dok_fak in view)
                            {
                                if (dok_fak.Features["ID_Zamowienia"].ToString() == lx.ID && dok_fak.BruttoCy.Value.ToString() == lx.Kwota && dok_fak.Definicja == def2)
                                {
                                    dokWyplata.NumeryDokumentow = dok_fak.Numer.NumerPelny;
                                    commercialDocument = dok_fak;
                                    break;
                                }
                            }

                            if (commercialDocument == null)
                            {
                                errorsList.Add(new OError(lx.ID, "Brak dokumentu", ""));
                                continue;
                            }

                            SubTable st = km.RozrachunkiIdx.WgPodmiot[commercialDocument.Platnosci.GetFirst().Podmiot, Date.MaxValue];
                            try
                            {
                                Wyplata paymentInRaport = null;
                                Zobowiazanie due = null;
                                foreach (RozrachunekIdx idx in st)
                                {
                                    if (idx.Typ == TypRozrachunku.Wypłata && paymentInRaport == null)
                                        paymentInRaport = (Wyplata)idx.Dokument;
                                    if (idx.Typ == TypRozrachunku.Zobowiązanie && due == null && !idx.Dokument.Bufor && idx.Numer == commercialDocument.Numer.Pelny)
                                        due = (Zobowiazanie)idx.Dokument;
                                    if (paymentInRaport != null && due != null)
                                        break;
                                }
                                if (paymentInRaport == null) continue;
                                paymentInRaport.DataDokumentu = rap.Okres.From;
                                using (ITransaction trans = session.Logout(true))
                                {
                                    RozliczenieSP settlement = new RozliczenieSP(due, (Wyplata)dokWyplata);
                                    km.RozliczeniaSP.AddRow(settlement);
                                    trans.Commit();
                                }
                            }
                            catch { }
                        }
                    }
                    t.Commit();
                }

                session.Save();

            }

            return new MessageBoxInformation("Przelewy24")
            {
                Type = MessageBoxInformationType.Information,
                Text = "Zakończono proces importowania płatności Przelewy24." + Environment.NewLine + "Odśwież listę lub naciśnij klawisz F5.",
                OKHandler = () => null
            };
        }
    }
}

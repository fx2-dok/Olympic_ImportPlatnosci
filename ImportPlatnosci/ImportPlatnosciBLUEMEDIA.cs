using Soneta.Business;
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

[assembly: Worker(typeof(ImportPlatnosci.ImportPlatnosciBLUEMEDIA), typeof(Soneta.Kasa.RaportESP))]

namespace ImportPlatnosci
{
    public class ImportPlatnosciBLUEMEDIA
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
            "Import transakcji Bluemedia",
            Description = "Funkcja importuje transakcje Bluemedia",
            Priority = 1000,
            Icon = ActionIcon.Wizard,
            Mode = ActionMode.SingleSession,
            Target = ActionTarget.Menu)]

        public void Funkcja()
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

            StreamReader objReader = new StreamReader(XMLFileName.FileName);
            string linia = "";
            int temp = 0;
            while (linia != null)
            {
                linia = objReader.ReadLine();

                if (linia != null && temp != 0)
                {
                    var kawalki = Regex.Split(linia, ";");
                    id = kawalki[2].ToString();
                    kwota = kawalki[5].Replace(".", ",").Trim();
                    kupujacy = kawalki[4].Replace("\"", "");
                    lxml.Add(new ListXML(id, data, kwota, "", "", operacja, kupujacy, ""));
                }
                temp++;
            }
            objReader.Close();

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


                View view = hm.DokHandlowe.CreateView();
                view.Condition &= new FieldCondition.Equal("Definicja", def);
                view.Condition &= new FieldCondition.Equal("Magazyn", magazyn_wysylkowy);
                view.Condition &= new FieldCondition.GreaterEqual("Data", "2022-03-01");

                using (ITransaction t = session.Logout(true))
                {

                    foreach (ListXML lx in lxml)
                    {
                        WplataRaport dok = new WplataRaport(rap);
                        km.Zaplaty.AddRow(dok);
                        dok.Podmiot = cm.Kontrahenci.WgKodu["PayNow"];
                        dok.Kwota = new Currency(System.Convert.ToDouble(lx.Kwota), "PLN");
                        dok.Opis = lx.Kupujacy;
                        string kontrahent = lx.Kupujacy;
                        string numery_dokumentow = "";

                        foreach (DokumentHandlowy dok_fak in view)
                        {

                            if (dok_fak.Obcy.Numer.ToString() == lx.ID)
                                dok.NumeryDokumentow = dok_fak.Numer.NumerPelny;
                        }
                    }



                    t.Commit();
                }
                session.Save();
            }
        }
        public static bool IsVisibleFunkcja(RaportESP raport)
        {
            return true;
        }

        public static bool IsEnabledFunkcja(RaportESP raport)
        {
            return raport.Zamknięty == false;
        }
    }
}

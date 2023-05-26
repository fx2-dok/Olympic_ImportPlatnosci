using Soneta.Business;
using Soneta.Business.UI;
using Soneta.CRM;
using Soneta.Handel;
using Soneta.Kasa;
using Soneta.Types;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

[assembly: Worker(typeof(ImportPlatnosci.ImportPlatnosciDPD), typeof(Soneta.Kasa.RaportESP))]

namespace ImportPlatnosci
{
    public class ImportPlatnosciDOTPAY
    {
        [Context, Required]
        public NamedStream CSVFileName { get; set; }

        RaportESP raport;
        [Context]
        public RaportESP Raport
        {
            get { return raport; }
            set { raport = value; }
        }

        [Action(
            "Import Transakcji DotPay",
            Description = "Funkcja importuje transakcje DotPay",
            Priority = 1000,
            Icon = ActionIcon.Wizard,
            Mode = ActionMode.SingleSession,
            Target = ActionTarget.Menu)]


        public MessageBoxInformation Funkcja()
        {
            ArrayList tablica = new ArrayList();

            var numer = new List<string>();
            var kanal = new List<string>();
            var data = new List<string>();
            var kwota = new List<string>();
            var prowizja = new List<string>();
            var wyplata = new List<string>();
            var opis = new List<string>();
            var control = new List<string>();



            StreamReader objReader = new StreamReader(CSVFileName.FileName);
            string linia = "";
            while (linia != null)
            {
                linia = objReader.ReadLine();
                if (linia != null)
                {
                    var kawalki = linia.Split(',');
                    numer.Add(kawalki[1].Replace("\"", ""));
                    kanal.Add(kawalki[2].Replace("\"", ""));
                    data.Add(kawalki[3].Replace("\"", ""));
                    kwota.Add(kawalki[4].Replace("\"", "").Replace('.', ','));
                    prowizja.Add(kawalki[5].Replace("\"", "").Replace('.', ','));
                    wyplata.Add(kawalki[6].Replace("\"", "").Replace('.', ','));
                    opis.Add(kawalki[7].Replace("\"", ""));
                    control.Add(kawalki[8].Replace("\"", ""));
                }
            }
            objReader.Close();


            using (Session session = raport.Session.Login.CreateSession(false, false))
            {
                KasaModule km = KasaModule.GetInstance(session);
                CRMModule cm = CRMModule.GetInstance(session);
                HandelModule hm = HandelModule.GetInstance(session);

                EwidencjaSP kasa = raport.Kasa;
                FromTo okres = raport.Okres;
                Date data1 = raport.Data;
                RaportESP rap = km.RaportyESP[raport.ID];

                using (ITransaction t = session.Logout(true))
                {

                    for (int i = 1; i < numer.Count; i++)
                    {
                        WplataRaport dok = new WplataRaport(rap);
                        km.Zaplaty.AddRow(dok);
                        dok.Podmiot = cm.Kontrahenci.WgKodu["DOT PAY"];
                        dok.Kwota = new Currency(System.Convert.ToDouble(kwota[i]), "PLN");

                        string dok_obcy_numer = "";

                        foreach (DokumentHandlowy dok_fak in hm.DokHandlowe)
                        {
                            if (control[i] != "brak")
                            {
                                if (dok_fak.Obcy.Numer == control[i]) dok_obcy_numer = dok_fak.Numer.ToString();
                            }

                        }

                        dok.NumeryDokumentow = dok_obcy_numer;

                        if (control[i] == "") control[i] = "brak";
                        dok.Opis = control[i];

                    }

                    WyplataRaport prowizja_paypal = new WyplataRaport(rap);
                    km.Zaplaty.AddRow(prowizja_paypal);
                    prowizja_paypal.Podmiot = cm.Kontrahenci.WgKodu["PayPal"];
                    double kwota_prowizji = 0;
                    for (int i = 1; i < prowizja.Count; i++) { kwota_prowizji = kwota_prowizji + System.Convert.ToDouble(prowizja[i]); }
                    prowizja_paypal.Kwota = new Currency(System.Convert.ToDouble(kwota_prowizji), "PLN");
                    prowizja_paypal.Opis = Path.GetFileName(CSVFileName.FileName).Split('.')[0];


                    t.Commit();
                }
                session.Save();


                return new MessageBoxInformation("Przelewy24", "Zakończono proces importowania płatności Przelewy24." + Environment.NewLine + "Odśwież listę lub naciśnij klawisz F5.");

            }
        }
    }
}

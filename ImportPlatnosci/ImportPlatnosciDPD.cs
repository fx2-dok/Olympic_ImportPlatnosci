using ExcelDataReader;
using ICSharpCode.NRefactory.TypeSystem;
using Soneta.Business;
using Soneta.Business.UI;
using Soneta.CRM;
using Soneta.Handel;
using Soneta.Kasa;
using Soneta.Magazyny;
using Soneta.Types;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

[assembly: Worker(typeof(ImportPlatnosci.ImportPlatnosciDPD), typeof(Soneta.Kasa.RaportESP))]

namespace ImportPlatnosci
{
    public class ImportPlatnosciDPD
    {
        [Context, Required]
        public NamedStream[] ExelFileNames { get; set; }

        [Context]
        public RaportESP Raport { get; set; }

        [Context]
        public Context Context { get; set; }

        [Action("Import transakcji DPD",
            Description = "Funkcja importuje transakcje DPD",
            Priority = 1000,
            Icon = ActionIcon.Wizard,
            Mode = ActionMode.SingleSession,
            Target = ActionTarget.Menu)]

        public MessageBoxInformation Funkcja()
        {
            IList<Payment> paymentList = new List<Payment>();
            IList<OError> errorsList = new List<OError>();

            foreach (var file in ExelFileNames)
            {
                using (var stream = File.Open(file.FileName, FileMode.Open, FileAccess.Read))
                {
                    var reader = ExcelReaderFactory.CreateReader(stream);
                    var result = reader.AsDataSet();
                    var table = result.Tables[0];

                    string[,] data = new string[table.Rows.Count - 1, table.Columns.Count];

                    int i = 0, j = 0;
                    foreach (DataColumn col in table.Columns)
                    {
                        j = 0;
                        int iTemp = -1;
                        foreach (DataRow row in table.Rows)
                        {
                            iTemp++;
                            if (iTemp == 0)
                            {
                                continue;
                            }
                            data[j, i] = row[col].ToString();
                            j++;
                        }
                        i++;
                    }

                    for (int iter = 0; iter < table.Rows.Count - 1; iter++)
                    {
                        string id = data[iter, 10].Split(',')[0];
                        decimal amountTempDouble = Convert.ToDecimal(data[iter, 4].Replace(".", ","));
                        Currency amount = new Currency(amountTempDouble);
                        Date date = Date.Parse(data[iter, 12]);
                        string description = data[iter, 6];
                        try
                        {
                            paymentList.Add(new Payment(date, id, amount, description, "wpłata"));
                        }
                        catch { errorsList.Add(new OError((iter + 1).ToString(), "Błąd odczytu z pliku", Path.GetFileName(file.FileName))); }
                    }
                }
            }

            using (Session session = Raport.Session.Login.CreateSession(false, false))
            {
                KasaModule km = KasaModule.GetInstance(session);
                CRMModule crm = CRMModule.GetInstance(session);
                HandelModule handel_module = HandelModule.GetInstance(session);
                MagazynyModule mm = MagazynyModule.GetInstance(session);

                EwidencjaSP kasa = Raport.Kasa;
                FromTo okres = Raport.Okres;
                Date data1 = Raport.Data;
                RaportESP raport = km.RaportyESP[Raport.ID];

                DefDokHandlowego invoicesDefinition = handel_module.DefDokHandlowych.WgSymbolu["ZIP"];

                using (ITransaction t = session.Logout(true))
                {
                    foreach (var payment in paymentList)
                    {
                        DokumentHandlowy commercialDocumentZIP = null;
                        DokumentHandlowy commercialDocument = null;

                        Soneta.Business.View view = handel_module.DokHandlowe.CreateView();
                        view.Condition &= new FieldCondition.Equal("Definicja", invoicesDefinition);
                        view.Condition &= new FieldCondition.Greater("Data", payment.Date.AddYears(-1));
                        view.Condition &= new FieldCondition.Equal("Numer", payment.Id);

                        if (view.Count == 0)
                        {
                            errorsList.Add(new OError(payment.Id, "Nie znaleziono dokumentu", "", payment));
                            continue;
                        }

                        foreach (DokumentHandlowy dh in view)
                        {
                            if (dh.BruttoCy.Value == payment.Amount)
                            {
                                //PłatnościDokumentuWorker pdw = new PłatnościDokumentuWorker();
                                //pdw.Dokument = dh;
                                //if (pdw.StanRozliczenia == StanRozliczenia.Nierozliczony)
                                //{
                                commercialDocumentZIP = dh;
                                break;
                                //}
                            }
                        }

                        if (commercialDocumentZIP == null)
                        {
                            errorsList.Add(new OError(payment.Id, "Dokument niezgodny z kwotą płatności z pliku lub rozliczony.", "DPD", payment));
                            continue;
                        }

                        foreach (DokumentHandlowy doc in commercialDocumentZIP.Podrzędne)
                        {
                            if (doc.Definicja.Symbol == "FF")
                            {
                                commercialDocument = doc;
                                break;
                            }
                        }

                        WplataRaport dok = new WplataRaport(raport);
                        km.Zaplaty.AddRow(dok);

                        Kontrahent commercialDocumentPayer = (Kontrahent)commercialDocument.Platnosci.GetFirst().Podmiot;
                        //Kontrahent commercialDocumentPayer = crm.Kontrahenci.WgNIP["5260204110"].GetFirst();  // ####  <---- Płatnik DPD czy osoba zamawiająca? ----> commercialDocument.Kontrahent;

                        dok.Podmiot = commercialDocumentPayer;
                        dok.Kwota = payment.Amount;
                        dok.Opis = payment.Description;
                        dok.NumeryDokumentow = commercialDocument.Numer.NumerPelny;
                        dok.SposobZaplaty = km.SposobyZaplaty.WgNazwy["Przelew"];

                        SubTable st = km.RozrachunkiIdx.WgPodmiot[commercialDocumentPayer, Date.MaxValue];
                        try
                        {
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
                            paymentInRaport.DataDokumentu = payment.Date;     //###  <---  ODBLOKOWAĆ NA PRODUKCJI

                            RozliczenieSP settlement = new RozliczenieSP(due, (Wplata)dok);
                            km.RozliczeniaSP.AddRow(settlement);
                        }
                        catch { errorsList.Add(new OError(payment.Id, "Nie rozliczono dokumentu", "", payment)); }//throw new Exception(e.Message.ToString()); }

                    }
                    t.Commit();
                }
                session.Save();
            }

            if (errorsList.Count != 0)
            {
                string info = "";
                foreach (var error in errorsList)
                {
                    info += error.Payment.Id + " - " + error.ErrorMessage + Environment.NewLine;
                }
                return new MessageBoxInformation("Import płatności DPD")
                {
                    Type = MessageBoxInformationType.Information,
                    Text = "Import zakończono pomyślnie z listą błędów " + errorsList.Count.ToString() + " dokumentów" + Environment.NewLine + info,
                    OKHandler = () => null
                };
            }
            return new MessageBoxInformation("Import płatności DPD")
            {
                Type = MessageBoxInformationType.Information,
                Text = "Zakończono proces importowania płatności PayU." + Environment.NewLine + "Odśwież listę lub naciśnij klawisz F5.",
                OKHandler = () => null
            };

        }
    }
}

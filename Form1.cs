using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Threading;
using Ionic.Zip;

namespace PG2
{
    public partial class Form1 : Form
    {
        public AutoResetEvent Downloading = new AutoResetEvent(true);
        public SortedDictionary<double,double> Data1  =new SortedDictionary<double,double>();
        public SortedDictionary<double, double> Data2 = new SortedDictionary<double,double>();


        public Form1()
        {
            InitializeComponent();
        }

        private void TextBoxID_KeyUp(object sender, EventArgs e)
        {
            
            MaskedTextBox source = (MaskedTextBox)sender;
            if (source.MaskCompleted)
            {
                Thread worker = new Thread(ValidID);
                worker.Start(source);
            }

        }

        private void MyMessage(string message, Label cible)
        {
            this.Invoke((MethodInvoker)delegate {
                cible.Text=message;
            });
        }

        private void ValidID(object inp)
        {
            //MessageBox.Show("hoho");            
            
            MaskedTextBox source = (MaskedTextBox)inp;
            string Filename;
            int iFile=2;
            Label state = label4;
            Label name = label9;
            if (source.Name == "maskedTextBox1")
            {
                iFile = 1;
                state = label2;
                name = label8;
            }

            MyMessage("Synchronisation", state);
            

            Downloading.WaitOne();
            Downloading.Reset();

            MyMessage("Démarrage du téléchargement", state);

            Filename = Filename = Application.UserAppDataPath + "\\stat"+iFile.ToString()+"z.zip";
            HttpWebRequest requete = (HttpWebRequest)WebRequest.Create("http://www.bdm.insee.fr/bdm2/exporterSeries.action");
            requete.Method = "Post";
            string data = "periode=toutes&nbPeriodes=0&liste_formats=txt&idbank="+source.Text;
            byte[] dataToSend = Encoding.UTF8.GetBytes(data);
            requete.ContentType = "application/x-www-form-urlencoded";
            requete.ContentLength = dataToSend.Length;
            Stream Input = requete.GetRequestStream();
            Input.Write(dataToSend, 0, dataToSend.Length);
            Input.Close();
            MyMessage("En Attente de la réponse", state);
            HttpWebResponse réponse = (HttpWebResponse)requete.GetResponse();
            int bufferSize = 1024;
            byte[] buffer = new byte[bufferSize];
            Stream resultat = réponse.GetResponseStream();

            bool fileDLed = false;
            try
            {
                var hd = File.Open(Filename, FileMode.Create);
                int pos;
                MyMessage("Téléchargement", state);
                while ((pos = resultat.Read(buffer, 0, bufferSize)) > 0)
                {
                    hd.Write(buffer, 0, bufferSize);
                }

                hd.Close();
                fileDLed = true;
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message, "ERREUR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                resultat.Close();
                réponse.Close();
            }
            try
            {
                if (fileDLed)
                {
                    MyMessage("Extraction", state);
                    using (var zip = new ZipFile(Filename))
                    {
                        zip.ExtractAll(Application.UserAppDataPath);

                        MyMessage("Suppression des fichiers temporaires",state);
                    }
                }
                else
                {
                    MessageBox.Show("pb");
                }
            }
            catch {
            }
            File.Delete(Application.UserAppDataPath + "\\F" + iFile.ToString() + ".txt");
            File.Delete(Application.UserAppDataPath + "\\Caract.csv");
            bool error = false;
            try
            {
                File.Move(Application.UserAppDataPath + "\\Valeurs.csv", Application.UserAppDataPath + "\\F" + iFile.ToString() + ".txt");
            }
            catch
            {
                MessageBox.Show("Impossible de récupérer les données correspondant à cet identifiant", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MyMessage("Erreur !", state);
                error = true;
            }
            finally
            {
                Downloading.Set();    
            }

            if (!error)
            {
                File.Delete(Filename);
                MyMessage("Analyse des données", state);

                var hand = File.Open(Application.UserAppDataPath + "\\F" + iFile.ToString() + ".txt", FileMode.Open);
                var Reader = new StreamReader(hand, Encoding.GetEncoding(1252));

                string[] anom = Reader.ReadLine().Split(';');
                string nom = anom[anom.Length - 1];
                MyMessage(nom, name);

                Reader.ReadLine();
                string[] headers = Reader.ReadLine().Split(';');

                double[] multiplicators= new double[headers.Length];

                for (int i = 0; i < headers.Length; i++)
                {
                    switch(headers[i])
                    {
                        case "année":
                            multiplicators[i] = 1F;
                            break;
                        case "Mois":
                            multiplicators[i] = (double)(1 / 12F);
                            break;
                        case "Trimestre" :
                            multiplicators[i] = (double)(1 / 4F);
                            break;
                        default :
                            multiplicators[i] = 1F;
                            break;
                    }

                }
                SortedDictionary<double,double> dico = new SortedDictionary<double,double>() ;
                while (!Reader.EndOfStream)
                {
                    var readed = Reader.ReadLine().Split(';');
                    double valeur;
                    if (Double.TryParse(readed[readed.Length - 1], out valeur))
                    {
                        double clé = 1D;
                        for (int i = 0; i < readed.Length - 1; i++)
                        {
                            double headVal;
                            if (!double.TryParse(readed[i], out headVal))
                            {
                                #region switch sur le mois
                                switch (readed[i])
                                {
                                    case "Janvier":
                                        headVal = 1D;
                                        break;
                                    case "Février":
                                        headVal = 2D;
                                        break;
                                    case "Mars":
                                        headVal = 3D;
                                        break;
                                    case "Avril":
                                        headVal = 4D;
                                        break;
                                    case "Mai":
                                        headVal = 5D;
                                        break;
                                    case "Juin":
                                        headVal = 6D;
                                        break;
                                    case "Juillet":
                                        headVal = 7D;
                                        break;
                                    case "Août":
                                        headVal = 8D;
                                        break;
                                    case "Septembre":
                                        headVal = 9D;
                                        break;
                                    case "Octobre":
                                        headVal = 10D;
                                        break;
                                    case "Novembre":
                                        headVal = 11D;
                                        break;
                                    case "Décembre":
                                        headVal = 12D;
                                        break;
                                    default:
                                        MessageBox.Show("Une donnée est corrompue, les résultats peuvent être faussés. Donnée :" + readed[i]);
                                        headVal = 1D;
                                        break;
                                }
                                #endregion
                            }
                            clé+=(headVal-1)*multiplicators[i];
                        }
                        dico.Add(clé, valeur);
                    }
                    else
                    {
                        MessageBox.Show("Une donnée est corrompue, ligne ignorée. Donnée :" + readed[readed.Length-1]);
                    }
                    
                }
                if (iFile == 1)
                {
                    Data1 = dico;
                }
                else
                {
                    Data2 = dico;
                }
                MyMessage("Terminé !", state);
                Reader.Close();
                hand.Close();

                if (label2.Text == "Terminé !" && label4.Text == "Terminé !")
                {
                    Analyse();
                }
            }
        }

        private void Analyse()
        {
            var keyC1 = Data1.Keys.ToList();
            var keyC2 = Data2.Keys.ToList();

            var tData1 = new SortedDictionary<double,double>();
            var tData2 = new SortedDictionary<double,double>();

            var intersected = keyC1.Intersect(keyC2);

            if (intersected.ToList().Count < 2)
            {
                MessageBox.Show("Ces  deux séries sont incompatibles car elles n'ont pas assez de dates en communs");
            }
            else
            {
                var min = intersected.Min();
                var max = intersected.Max();
                for (int i = keyC2.IndexOf(min); i < keyC2.IndexOf(max)+1; i++)
                {
                    tData2.Add(keyC2[i],Data2.Values.ToList()[i]);
                }
                for (int i = keyC1.IndexOf(min); i < keyC1.IndexOf(max)+1; i++)
                {
                    tData1.Add(keyC1[i], Data1.Values.ToList()[i]);
                }

                SortedDictionary<double, double> t1Data, t2Data,temp;
                t2Data = new SortedDictionary<double,double>();

                if (tData1.Count > tData2.Count)
                {
                    t1Data = tData2;
                    temp = tData1;
                }
                else
                {
                    t1Data = tData1;
                    temp = tData2;
                }

                int index = 0;
                var kT1D = t1Data.Keys.ToList();
                var kTT = temp.Keys.ToList();
                var vTT = temp.Values.ToList();

                for (int i = 0; i < kT1D.Count;i++ )
                {
                    double somme = vTT[index];
                    int nb = 1;
                    do{
                        somme+=vTT[index];
                        nb++;
                        index++;
                        if (index >= kTT.Count)
                            break;
                    }while(kT1D[i+1]!=kTT[index]);

                    t2Data.Add(kT1D[i], somme / nb);
                }

                double somme1=0;
                double somme2=0;
                List<double> D1, D2;
                D1 = t1Data.Values.ToList();
                D2 = t2Data.Values.ToList();

                for (int i = 0; i < t2Data.Count - 1; i++)
                {
                    double a, b;
                    a = D1[i];
                    b = D1[i + 1];
                    if (a > b)
                        somme1 += a / b;
                    else
                        somme1 += b / a;

                    a = D2[i];
                    b = D2[i + 1];
                    if (a > b)
                    {
                        somme2 += a / b;
                    }
                    else
                    {
                        somme2 += b / a;
                    }
                }
                somme1 /= D1.Count-1;
                somme2 /= D1.Count-1;

                double similaire = 1D;

                for (int i = 0; i < t2Data.Count - 1; i++)
                {
                    double a,b,c,d;
                     a = D1[i];
                    b = D1[i + 1];
                     c = D2[i];
                    d = D2[i + 1];

                    double varA = (b - a) / somme1;
                    double varB = (d - c) / somme2;

                    if (varA * varB != 0)
                    {
                        int signe;
                        if (varA * varB > 0)
                        {
                            signe = 1;
                        }
                        else
                        {
                            signe = -1;
                        }
                        varA = Math.Abs(varA);
                        varB = Math.Abs(varB);
                        double rapport;
                        if (varA < varB)
                        {
                            rapport = varA / varB;
                        }
                        else
                        {
                            rapport = varB / varA;
                        }

                        similaire += signe * rapport;
                    }
                    else
                    {
                        if (varA == varB)
                        {
                            similaire += 1;
                        }
                    }
                }

                similaire /= (D1.Count);
                similaire *= 100;

                this.Invoke((MethodInvoker)delegate
                {
                    label5.Text = Math.Round(similaire, 1).ToString()+" %";
                    progressBar3.Value =  (int) Math.Round(similaire,0)+100;
                });

            }
        }

    }
}

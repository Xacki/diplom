using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Runtime.InteropServices;
using System.IO;
namespace diplom
{
    public partial class General : Form
    {
        public int pos = 0;
        Word.Application w;
        
        public General()
        {
            InitializeComponent();
        }
        
        class PrPrakt
        {
            public string vid, dolj, dateOt, dateDo;
            public PrPrakt(string vid = "x", string dolj = "x", string dateOt = "x", string dateDo = "x")
            {
                this.vid = vid;
                this.dolj = dolj;
                this.dateOt = dateOt;
                this.dateDo = dateDo;
            }

        }

        public void addPraktStud()
        {
            bool studHavepredm = false;
            foreach (DataRowView predmGrup in prPraktBS)
            {
                studHavepredm = false;
                if(studBS.Count != 0)
                    if (praktBS.Count == 0)
                    {
                        praktBS.AddNew();
                        ((DataRowView)praktBS.Current)["vid"] = ((DataRowView)prPraktBS.Current)["vid"];
                        ((DataRowView)praktBS.Current)["dolj"] = ((DataRowView)prPraktBS.Current)["dolj"];
                        ((DataRowView)praktBS.Current)["dateOt"] = ((DataRowView)prPraktBS.Current)["dateOt"];
                        ((DataRowView)praktBS.Current)["dateDo"] = ((DataRowView)prPraktBS.Current)["dateDo"];
                        praktBS.EndEdit();
                    }
                foreach (DataRowView predmStud in praktBS) 
                {
                    if (predmGrup["vid"].ToString() == predmStud["vid"].ToString() && predmStud["idStud"].ToString() == (studBS.Current as DataRowView)["id"].ToString())
                    {
                        studHavepredm = true;
                        break;
                    }
                }
                if (studBS.Count != 0)
                    if (!studHavepredm)
                    {
                        praktBS.AddNew();
                        ((DataRowView)praktBS.Current)["vid"] = predmGrup["vid"];
                        ((DataRowView)praktBS.Current)["dolj"] = predmGrup["dolj"];
                        ((DataRowView)praktBS.Current)["dateOt"] = predmGrup["dateOt"];
                        ((DataRowView)praktBS.Current)["dateDo"] = predmGrup["dateDo"];
                        praktBS.EndEdit();
                    }
                

            }
            if(praktBS.Count !=0)
                praktTA.Update(portfolioDS.prakt);
        }

        public void addDisciplPrakt()
        {

            List<PrPrakt> pr = new List<PrPrakt>();

            foreach (DataRowView row in prPraktBS)
            {
                pr.Add(new PrPrakt(row["vid"].ToString(), row["dolj"].ToString(), row["dateOt"].ToString(), row["dateDo"].ToString()));
            }
            if (studBS.Count != 0)
                if (praktBS.Count == 0)
                {
                    praktBS.AddNew();
                    ((DataRowView)praktBS.Current)["vid"] = ((DataRowView)prPraktBS.Current)["vid"];
                    ((DataRowView)praktBS.Current)["dolj"] = ((DataRowView)prPraktBS.Current)["dolj"];
                    ((DataRowView)praktBS.Current)["dateOt"] = ((DataRowView)prPraktBS.Current)["dateOt"];
                    ((DataRowView)praktBS.Current)["dateDo"] = ((DataRowView)prPraktBS.Current)["dateDo"];
                    praktBS.EndEdit();
                }
            bool have = true;
            foreach (DataRowView stud in studBS)
            {
                foreach (PrPrakt predm in pr) 
                {
                    foreach (DataRowView rez in praktAllBS) 
                    {
                        if (rez["idStud"].ToString() == stud["id"].ToString())
                        {
                            if (rez["vid"].ToString() == predm.vid)
                            {
                                have = true;
                                break;
                            }
                            have = false;
                        }
                    }

                    if (!have)
                    {
                        praktAllBS.AddNew();
                        ((DataRowView)praktAllBS.Current)["vid"] = predm.vid;
                        ((DataRowView)praktAllBS.Current)["dolj"] = predm.dolj;
                        if (predm.dateOt != "")
                            ((DataRowView)praktAllBS.Current)["dateOt"] = Convert.ToDateTime(predm.dateOt);
                        if (predm.dateOt != "")
                            ((DataRowView)praktAllBS.Current)["dateDo"] = Convert.ToDateTime(predm.dateDo);
                        ((DataRowView)praktAllBS.Current)["idStud"] = stud["id"].ToString();
                        praktAllBS.EndEdit();
                    }
                }
            }
            if (praktAllBS.Count != 0)
                praktTA.Update(portfolioDS.prakt);
        }

        public void delDisciplPrakt()
        {
            if (studBS.Count != 0)
                if (praktBS.Count == 0)
                {
                    praktBS.AddNew();
                    ((DataRowView)praktBS.Current)["vid"] = ((DataRowView)prPraktBS.Current)["vid"];
                    ((DataRowView)praktBS.Current)["dolj"] = ((DataRowView)prPraktBS.Current)["dolj"];
                    ((DataRowView)praktBS.Current)["dateOt"] = ((DataRowView)prPraktBS.Current)["dateOt"];
                    ((DataRowView)praktBS.Current)["dateDo"] = ((DataRowView)prPraktBS.Current)["dateDo"];
                    praktBS.EndEdit();
                }
            List<PrPrakt> pr = new List<PrPrakt>();

            foreach (DataRowView row in prPraktBS)
            {
                pr.Add(new PrPrakt(row["vid"].ToString(), row["dolj"].ToString(), row["dateOt"].ToString(), row["dateDo"].ToString()));
            }
            string id = "";
            bool exist = false;

            foreach (DataRowView stud in studBS)
            {
                foreach (DataRowView rez in praktAllBS)
                {
                    foreach (PrPrakt predm in pr)
                    {
                        if (rez["idStud"].ToString() == stud["id"].ToString())
                        {
                            if (rez["vid"].ToString() == predm.vid)
                            {
                                exist = true;
                                break;
                            }
                            else
                            {
                                exist = false;
                            }
                        }
                    }
                    if (rez["idStud"].ToString() == stud["id"].ToString())
                        if (!exist)
                        {
                            rez.Delete();
                            exist = true;
                        }

                }
            }
            if (praktAllBS.Count != 0)
                praktTA.Update(portfolioDS.prakt);
        }

        class PrKurs
        {
            public string disc, sem;

            public PrKurs(string disc = "x", string sem = "x")
            {
                this.disc = disc;
                this.sem = sem;
            }

        }

        public void addKursStud()
        {
            bool studHavepredm = false;
            foreach (DataRowView predmGrup in prKursBS)
            {
                studHavepredm = false;
                if (studBS.Count != 0)
                    if (kursBS.Count == 0)
                    {
                        kursBS.AddNew();
                        ((DataRowView)kursBS.Current)["disc"] = ((DataRowView)prKursBS.Current)["disc"];
                        ((DataRowView)kursBS.Current)["sem"] = ((DataRowView)prKursBS.Current)["sem"];
                        kursBS.EndEdit();
                    }
                foreach (DataRowView predmStud in kursBS)
                {
                    if (predmGrup["disc"].ToString() == predmStud["disc"].ToString() && predmStud["idStud"].ToString() == (studBS.Current as DataRowView)["id"].ToString())
                    {
                        studHavepredm = true;
                        break;
                    }
                }
                if (kursBS.Count != 0)
                    if (!studHavepredm)
                    {
                        kursBS.AddNew();
                        ((DataRowView)kursBS.Current)["disc"] = predmGrup["disc"];
                        ((DataRowView)kursBS.Current)["sem"] = predmGrup["sem"];
                        kursBS.EndEdit();
                    }


            }
            if (kursBS.Count != 0)
                kursTA.Update(portfolioDS.kurs);
        }

        public void addDisciplKurs()
        {
            if (studBS.Count != 0)
                if (kursBS.Count == 0)
                {
                    kursBS.AddNew();
                    ((DataRowView)kursBS.Current)["disc"] = ((DataRowView)prKursBS.Current)["disc"];
                    ((DataRowView)kursBS.Current)["sem"] = ((DataRowView)prKursBS.Current)["sem"];
                    kursBS.EndEdit();
                }
            List<PrKurs> pr = new List<PrKurs>();

            foreach (DataRowView row in prKursBS)
            {
                pr.Add(new PrKurs(row["disc"].ToString(), row["sem"].ToString()));
            }

            bool have = true;
            foreach (DataRowView stud in studBS)
            {
                foreach (PrKurs predm in pr)
                {
                    foreach (DataRowView rez in kursAllBS)
                    {
                        if (rez["idStud"].ToString() == stud["id"].ToString())
                        {
                            if (rez["disc"].ToString() == predm.disc)
                            {
                                have = true;
                                break;
                            }
                            have = false;
                        }
                    }

                    if (!have)
                    {
                        kursAllBS.AddNew();
                        ((DataRowView)kursAllBS.Current)["disc"] = predm.disc;
                        ((DataRowView)kursAllBS.Current)["sem"] = predm.sem;
                        ((DataRowView)kursAllBS.Current)["idStud"] = stud["id"].ToString();
                        kursAllBS.EndEdit();
                    }
                }
            }
            if (kursAllBS.Count != 0)
                kursTA.Update(portfolioDS.kurs);
        }

        public void delDisciplKurs()
        {
            if (studBS.Count != 0)
                if (kursBS.Count == 0)
                {
                    kursBS.AddNew();
                    ((DataRowView)kursBS.Current)["disc"] = ((DataRowView)prKursBS.Current)["disc"];
                    ((DataRowView)kursBS.Current)["sem"] = ((DataRowView)prKursBS.Current)["sem"];
                    kursBS.EndEdit();
                }
            List<PrKurs> pr = new List<PrKurs>();

            foreach (DataRowView row in prKursBS)
            {
                pr.Add(new PrKurs(row["disc"].ToString(), row["sem"].ToString()));
            }
            bool exist = false;

            foreach (DataRowView stud in studBS)
            {
                foreach (DataRowView rez in kursAllBS)
                {
                    foreach (PrKurs predm in pr)
                    {
                        if (rez["idStud"].ToString() == stud["id"].ToString())
                        {
                            if (rez["disc"].ToString() == predm.disc)
                            {
                                exist = true;
                                break;
                            }
                            else
                            {
                                exist = false;
                            }
                        }
                    }
                    if (rez["idStud"].ToString() == stud["id"].ToString())
                        if (!exist)
                        {
                            rez.Delete();
                            exist = true;
                        }

                }
            }
            if (kursAllBS.Count != 0)
                kursTA.Update(portfolioDS.kurs);
        }

        class PrDisc
        {
            public string disc, nagr,forma,sem;

            public PrDisc(string disc = "x", string nagr = "x", string forma = "x", string sem = "x")
            {
                this.disc = disc;
                this.nagr = nagr;
                this.forma = forma;
                this.sem = sem;
            }
        }

        public void addDisciplStud()
        {
            bool studHavepredm = false;
            foreach (DataRowView predmGrup in prDiscBS)
            {
                studHavepredm = false;
                if (studBS.Count != 0)
                    if (discBS.Count == 0)
                    {
                        discBS.AddNew();
                        ((DataRowView)discBS.Current)["disc"] = ((DataRowView)prDiscBS.Current)["disc"];
                        ((DataRowView)discBS.Current)["nagr"] = ((DataRowView)prDiscBS.Current)["nagr"];
                        ((DataRowView)discBS.Current)["forma"] = ((DataRowView)prDiscBS.Current)["forma"];
                        ((DataRowView)discBS.Current)["sem"] = ((DataRowView)prDiscBS.Current)["sem"];
                        discBS.EndEdit();
                    }
                foreach (DataRowView predmStud in discBS)
                {
                    if (predmGrup["disc"].ToString() == predmStud["disc"].ToString() &&
                        predmGrup["sem"].ToString() == predmStud["sem"].ToString() && 
                        predmStud["idStud"].ToString() == (studBS.Current as DataRowView)["id"].ToString())
                    {
                        studHavepredm = true;
                        break;
                    }
                }
                if (studBS.Count != 0)
                    if (!studHavepredm)
                    {
                        discBS.AddNew();
                        ((DataRowView)discBS.Current)["disc"] = predmGrup["disc"];
                        ((DataRowView)discBS.Current)["nagr"] = predmGrup["nagr"];
                        ((DataRowView)discBS.Current)["forma"] = predmGrup["forma"];
                        ((DataRowView)discBS.Current)["sem"] = predmGrup["sem"];
                        discBS.EndEdit();
                    }


            }
            if (discBS.Count != 0)
                discTA.Update(portfolioDS.disc);
        }

        public void addDisciplDisc()
        {
            if (studBS.Count != 0)
                if (discBS.Count == 0)
                {
                    discBS.AddNew();
                    ((DataRowView)discBS.Current)["disc"] = ((DataRowView)prDiscBS.Current)["disc"];
                    ((DataRowView)discBS.Current)["nagr"] = ((DataRowView)prDiscBS.Current)["nagr"];
                    ((DataRowView)discBS.Current)["forma"] = ((DataRowView)prDiscBS.Current)["forma"];
                    ((DataRowView)discBS.Current)["sem"] = ((DataRowView)prDiscBS.Current)["sem"];
                    discBS.EndEdit();
                }

            List<PrDisc> pr = new List<PrDisc>();
            foreach (DataRowView row in prDiscBS)
            {
                pr.Add(new PrDisc(row["disc"].ToString(), row["nagr"].ToString(), row["forma"].ToString(), row["sem"].ToString()));
            }

            bool have = true;
            foreach (DataRowView stud in studBS)
            {
                foreach (PrDisc predm in pr)
                {
                    foreach (DataRowView rez in discAllBS)
                    {
                        if (rez["idStud"].ToString() == stud["id"].ToString())
                        {
                            if (rez["disc"].ToString() == predm.disc && rez["sem"].ToString() == predm.sem)
                            {
                                have = true;
                                break;
                            }

                            have = false;
                        }
                    }

                    if (!have)
                    {
                        discAllBS.AddNew();
                        ((DataRowView)discAllBS.Current)["disc"] = predm.disc;
                        ((DataRowView)discAllBS.Current)["nagr"] = predm.nagr;
                        ((DataRowView)discAllBS.Current)["sem"] = predm.sem;
                        ((DataRowView)discAllBS.Current)["forma"] = predm.forma;
                        ((DataRowView)discAllBS.Current)["idStud"] = stud["id"].ToString();
                        discAllBS.EndEdit();
                    }
                }
            }
            if (discAllBS.Count != 0)
                discTA.Update(portfolioDS.disc);
        }

        public void delDisciplDisc()
        {
            if (studBS.Count != 0)
                if (discBS.Count == 0)
                {
                    discBS.AddNew();
                    ((DataRowView)discBS.Current)["disc"] = ((DataRowView)prDiscBS.Current)["disc"];
                    ((DataRowView)discBS.Current)["nagr"] = ((DataRowView)prDiscBS.Current)["nagr"];
                    ((DataRowView)discBS.Current)["forma"] = ((DataRowView)prDiscBS.Current)["forma"];
                    ((DataRowView)discBS.Current)["sem"] = ((DataRowView)prDiscBS.Current)["sem"];
                    discBS.EndEdit();
                }

            List<PrDisc> pr = new List<PrDisc>();

            foreach (DataRowView row in prDiscBS)
            {
                pr.Add(new PrDisc(row["disc"].ToString(), row["nagr"].ToString(), row["forma"].ToString(), row["sem"].ToString()));
            }
            string id = "";
            bool exist = false;

            foreach (DataRowView stud in studBS)
            {
                foreach (DataRowView rez in discAllBS)
                {
                    foreach (PrDisc predm in pr)
                    {
                        if (rez["idStud"].ToString() == stud["id"].ToString())
                        {
                            if (rez["disc"].ToString() == predm.disc && rez["sem"].ToString() == predm.sem)
                            {
                                exist = true;
                                break;
                            }
                            else
                            {
                                exist = false;
                            }
                        }
                    }
                    if (rez["idStud"].ToString() == stud["id"].ToString())
                        if (!exist)
                        {
                            rez.Delete();
                            exist = true;
                        }

                }
            }
            if (discAllBS.Count != 0)
                discTA.Update(portfolioDS.disc);
        }

        private void show(string nameFolder,DataRowView curGroup,DataRowView curStud) 
        {
            try
            {
                w = Marshal.GetActiveObject("Word.Application") as Word.Application;
            }
            catch (COMException err)
            {
                w = new Word.Application();
            }
            string s = "";
            int k = 0;
            int i = 0;

            progressBar1.Value = 0;
            progressBar1.Maximum = prDiscBS.Count + 1;
            progressBar1.Visible = true;

            w.Documents.Add(Application.StartupPath + "\\Templates\\anketa.dot");
            
            w.ActiveDocument.Bookmarks["fam"].Range.Text = curStud["fam"].ToString();
            w.ActiveDocument.Bookmarks["name"].Range.Text = curStud["name"].ToString();
            w.ActiveDocument.Bookmarks["otch"].Range.Text = curStud["otch"].ToString();
            w.ActiveDocument.Bookmarks["spec"].Range.Text = (specBS.Current as DataRowView)["naz"].ToString();
            w.ActiveDocument.Bookmarks["nomZachisl"].Range.Text = curStud["nomZachisl"].ToString();
            w.ActiveDocument.Bookmarks["formO"].Range.Text = curGroup["formO"].ToString();
            w.ActiveDocument.Bookmarks["group"].Range.Text = curGroup["naz"].ToString();
            w.ActiveDocument.Bookmarks["tipJaz"].Range.Text = curStud["urVlad"].ToString();
            w.ActiveDocument.Bookmarks["tel"].Range.Text = curStud["tel"].ToString();
            w.ActiveDocument.Bookmarks["email"].Range.Text = curStud["email"].ToString();
            w.ActiveDocument.Bookmarks["fioLast"].Range.Text = curStud["fam"].ToString() + " " + curStud["name"].ToString() + " " + curStud["otch"].ToString();
            if (curStud["dateR"].ToString() != "")
                w.ActiveDocument.Bookmarks["dateR"].Range.Text = Convert.ToDateTime(curStud["dateR"]).ToString("dd/MM/yyyy");

            if (curStud["dateZachisl"].ToString() != "")
                w.ActiveDocument.Bookmarks["dateZachisl"].Range.Text = Convert.ToDateTime(curStud["dateZachisl"]).ToString("dd/MM/yyyy");


            Byte[] blob = null;
            MemoryStream memStream = null;
            Image img = null;
            if (curStud["foto"].ToString() != "")
            {
                blob = (byte[])curStud["foto"];
                memStream = new MemoryStream(blob);
                memStream.Write(blob, 0, blob.Length);
                memStream.Position = 0;
                img = Image.FromStream(memStream);
                Clipboard.SetImage(img);
                w.ActiveDocument.Bookmarks["foto"].Range.Paste();
                w.ActiveDocument.InlineShapes[1].Width = 90;
                w.ActiveDocument.InlineShapes[1].Height = 120;
            }

            foreach (DataRowView row in jazBS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString())
                {
                    s += row["nameJaz"].ToString() + ", ";
                }
            }
            if (s != "")
            {
                s = s.Remove(s.Length - 2);
            }
            
            w.ActiveDocument.Bookmarks["nameJaz"].Range.Text = s;
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in prikazPerevodBS)
            {
                if (row["idGroup"].ToString() == curGroup["id"].ToString())
                {
                    w.ActiveDocument.Tables[3].Rows.Add().SetHeight(0, 0);
                    w.ActiveDocument.Tables[3].Rows[k].Cells[1].Range.Text = row["kurs"].ToString();
                    if (row["date"].ToString() != "")
                        w.ActiveDocument.Tables[3].Rows[k].Cells[2].Range.Text = row["nomer"].ToString() + " от "+ Convert.ToDateTime(row["date"]).ToString("dd.MM.yyyy");
                    w.ActiveDocument.Tables[3].Rows[k].Cells[3].Range.Text = row["soderj"].ToString();
                    k++;
                }
            }
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in prikazStudBS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString() && row["tip"].ToString() == "Перерыв, в академическом отпуске")
                {
                    w.ActiveDocument.Tables[4].Rows.Add().SetHeight(0, 0);
                    w.ActiveDocument.Tables[4].Rows[k].Cells[1].Range.Text = row["kurs"].ToString();
                    if (row["date"].ToString() != "")
                        w.ActiveDocument.Tables[4].Rows[k].Cells[2].Range.Text = row["nomer"].ToString() + " от " + Convert.ToDateTime(row["date"]).ToString("dd.MM.yyyy");
                    w.ActiveDocument.Tables[4].Rows[k].Cells[3].Range.Text = row["soderj"].ToString();
                    k++;
                }
            }
            if (w.ActiveDocument.Tables[4].Rows.Count==1)
	        {
                w.ActiveDocument.Tables[4].Rows.Add().SetHeight(0, 0);
                w.ActiveDocument.Tables[4].Rows[k].Cells[1].Range.Text = "-";
                w.ActiveDocument.Tables[4].Rows[k].Cells[2].Range.Text = "-";
                w.ActiveDocument.Tables[4].Rows[k].Cells[3].Range.Text = "-";
	        }
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in prikazStudBS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString() && row["tip"].ToString() == "Поощрения")
                {
                    w.ActiveDocument.Tables[5].Rows.Add().SetHeight(0, 0);
                    w.ActiveDocument.Tables[5].Rows[k].Cells[1].Range.Text = row["kurs"].ToString();
                    if (row["date"].ToString() != "")
                        w.ActiveDocument.Tables[5].Rows[k].Cells[2].Range.Text = row["nomer"].ToString() + " от " + Convert.ToDateTime(row["date"]).ToString("dd.MM.yyyy");
                    w.ActiveDocument.Tables[5].Rows[k].Cells[3].Range.Text = row["soderj"].ToString();
                    k++;
                }
            }
            if (w.ActiveDocument.Tables[5].Rows.Count == 1)
            {
                w.ActiveDocument.Tables[5].Rows.Add().SetHeight(0, 0);
                w.ActiveDocument.Tables[5].Rows[k].Cells[1].Range.Text = "-";
                w.ActiveDocument.Tables[5].Rows[k].Cells[2].Range.Text = "-";
                w.ActiveDocument.Tables[5].Rows[k].Cells[3].Range.Text = "-";
            }
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in prikazStudBS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString() && row["tip"].ToString() == "Взыскания")
                {
                    w.ActiveDocument.Tables[6].Rows.Add().SetHeight(0, 0);
                    w.ActiveDocument.Tables[6].Rows[k].Cells[1].Range.Text = row["kurs"].ToString();
                    if (row["date"].ToString() != "")
                        w.ActiveDocument.Tables[6].Rows[k].Cells[2].Range.Text = row["nomer"].ToString() + " от " + Convert.ToDateTime(row["date"]).ToString("dd.MM.yyyy");
                    w.ActiveDocument.Tables[6].Rows[k].Cells[3].Range.Text = row["soderj"].ToString();
                    k++;
                }
            }
            if (w.ActiveDocument.Tables[6].Rows.Count == 1)
            {
                w.ActiveDocument.Tables[6].Rows.Add().SetHeight(0, 0);
                w.ActiveDocument.Tables[6].Rows[k].Cells[1].Range.Text = "-";
                w.ActiveDocument.Tables[6].Rows[k].Cells[2].Range.Text = "-";
                w.ActiveDocument.Tables[6].Rows[k].Cells[3].Range.Text = "-";
            }
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in prevObrBS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString())
                {
                    w.ActiveDocument.Tables[7].Rows.Add().SetHeight(0, 0);

                    if (row["dateOt"].ToString() != "" && row["dateDo"].ToString() != "")
                        w.ActiveDocument.Tables[7].Rows[k].Cells[1].Range.Text =
                            Convert.ToDateTime(row["dateOt"]).ToString("dd.MM.yyyy")
                            + " - " +
                            Convert.ToDateTime(row["dateDo"]).ToString("dd.MM.yyyy");

                    w.ActiveDocument.Tables[7].Rows[k].Cells[2].Range.Text = row["ucher"].ToString();
                    w.ActiveDocument.Tables[7].Rows[k].Cells[3].Range.Text = row["kval"].ToString();
                    k++;
                }
            }
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in discAllBS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString() && row["sem"].ToString() == "1")
                {
                    w.ActiveDocument.Tables[8].Rows.Add().SetHeight(0, 0);
                    w.ActiveDocument.Tables[8].Rows[k].Cells[1].Range.Text = (k - 1).ToString();
                    w.ActiveDocument.Tables[8].Rows[k].Cells[2].Range.Text = row["disc"].ToString();
                    w.ActiveDocument.Tables[8].Rows[k].Cells[3].Range.Text = row["nagr"].ToString();
                    w.ActiveDocument.Tables[8].Rows[k].Cells[4].Range.Text = row["forma"].ToString();
                    w.ActiveDocument.Tables[8].Rows[k].Cells[5].Range.Text = row["itog"].ToString();
                    k++;
                    progressBar1.Value++;
                }
            }
            w.ActiveDocument.Tables[8].Rows[k].Delete();
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in discAllBS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString() && row["sem"].ToString() == "2")
                {
                    w.ActiveDocument.Tables[9].Rows.Add().SetHeight(0, 0);
                    w.ActiveDocument.Tables[9].Rows[k].Cells[1].Range.Text = (k - 1).ToString();
                    w.ActiveDocument.Tables[9].Rows[k].Cells[2].Range.Text = row["disc"].ToString();
                    w.ActiveDocument.Tables[9].Rows[k].Cells[3].Range.Text = row["nagr"].ToString();
                    w.ActiveDocument.Tables[9].Rows[k].Cells[4].Range.Text = row["forma"].ToString();
                    w.ActiveDocument.Tables[9].Rows[k].Cells[5].Range.Text = row["itog"].ToString();
                    k++;
                    progressBar1.Value++;
                }
            }
            w.ActiveDocument.Tables[9].Rows[k].Delete();
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in discAllBS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString() && row["sem"].ToString() == "3")
                {
                    w.ActiveDocument.Tables[10].Rows.Add().SetHeight(0, 0);
                    w.ActiveDocument.Tables[10].Rows[k].Cells[1].Range.Text = (k - 1).ToString();
                    w.ActiveDocument.Tables[10].Rows[k].Cells[2].Range.Text = row["disc"].ToString();
                    w.ActiveDocument.Tables[10].Rows[k].Cells[3].Range.Text = row["nagr"].ToString();
                    w.ActiveDocument.Tables[10].Rows[k].Cells[4].Range.Text = row["forma"].ToString();
                    w.ActiveDocument.Tables[10].Rows[k].Cells[5].Range.Text = row["itog"].ToString();
                    k++;
                    progressBar1.Value++;
                }
            }
            w.ActiveDocument.Tables[10].Rows[k].Delete();
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in discAllBS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString() && row["sem"].ToString() == "4")
                {
                    w.ActiveDocument.Tables[11].Rows.Add().SetHeight(0, 0);
                    w.ActiveDocument.Tables[11].Rows[k].Cells[1].Range.Text = (k - 1).ToString();
                    w.ActiveDocument.Tables[11].Rows[k].Cells[2].Range.Text = row["disc"].ToString();
                    w.ActiveDocument.Tables[11].Rows[k].Cells[3].Range.Text = row["nagr"].ToString();
                    w.ActiveDocument.Tables[11].Rows[k].Cells[4].Range.Text = row["forma"].ToString();
                    w.ActiveDocument.Tables[11].Rows[k].Cells[5].Range.Text = row["itog"].ToString();
                    k++;
                    progressBar1.Value++;
                }
            }
            w.ActiveDocument.Tables[11].Rows[k].Delete();
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in discAllBS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString() && row["sem"].ToString() == "5")
                {
                    w.ActiveDocument.Tables[12].Rows.Add().SetHeight(0, 0);
                    w.ActiveDocument.Tables[12].Rows[k].Cells[1].Range.Text = (k - 1).ToString();
                    w.ActiveDocument.Tables[12].Rows[k].Cells[2].Range.Text = row["disc"].ToString();
                    w.ActiveDocument.Tables[12].Rows[k].Cells[3].Range.Text = row["nagr"].ToString();
                    w.ActiveDocument.Tables[12].Rows[k].Cells[4].Range.Text = row["forma"].ToString();
                    w.ActiveDocument.Tables[12].Rows[k].Cells[5].Range.Text = row["itog"].ToString();
                    k++;
                    progressBar1.Value++;
                }
            }
            w.ActiveDocument.Tables[12].Rows[k].Delete();
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in discAllBS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString() && row["sem"].ToString() == "6")
                {
                    w.ActiveDocument.Tables[13].Rows.Add().SetHeight(0, 0);
                    w.ActiveDocument.Tables[13].Rows[k].Cells[1].Range.Text = (k - 1).ToString();
                    w.ActiveDocument.Tables[13].Rows[k].Cells[2].Range.Text = row["disc"].ToString();
                    w.ActiveDocument.Tables[13].Rows[k].Cells[3].Range.Text = row["nagr"].ToString();
                    w.ActiveDocument.Tables[13].Rows[k].Cells[4].Range.Text = row["forma"].ToString();
                    w.ActiveDocument.Tables[13].Rows[k].Cells[5].Range.Text = row["itog"].ToString();
                    k++;
                    progressBar1.Value++;
                }
            }
            w.ActiveDocument.Tables[13].Rows[k].Delete();
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in discAllBS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString() && row["sem"].ToString() == "7")
                {
                    w.ActiveDocument.Tables[14].Rows.Add().SetHeight(0, 0);
                    w.ActiveDocument.Tables[14].Rows[k].Cells[1].Range.Text = (k - 1).ToString();
                    w.ActiveDocument.Tables[14].Rows[k].Cells[2].Range.Text = row["disc"].ToString();
                    w.ActiveDocument.Tables[14].Rows[k].Cells[3].Range.Text = row["nagr"].ToString();
                    w.ActiveDocument.Tables[14].Rows[k].Cells[4].Range.Text = row["forma"].ToString();
                    w.ActiveDocument.Tables[14].Rows[k].Cells[5].Range.Text = row["itog"].ToString();
                    k++;
                    progressBar1.Value++;
                }
            }
            w.ActiveDocument.Tables[14].Rows[k].Delete();
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in discAllBS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString() && row["sem"].ToString() == "8")
                {
                    w.ActiveDocument.Tables[15].Rows.Add().SetHeight(0, 0);
                    w.ActiveDocument.Tables[15].Rows[k].Cells[1].Range.Text = (k - 1).ToString();
                    w.ActiveDocument.Tables[15].Rows[k].Cells[2].Range.Text = row["disc"].ToString();
                    w.ActiveDocument.Tables[15].Rows[k].Cells[3].Range.Text = row["nagr"].ToString();
                    w.ActiveDocument.Tables[15].Rows[k].Cells[4].Range.Text = row["forma"].ToString();
                    w.ActiveDocument.Tables[15].Rows[k].Cells[5].Range.Text = row["itog"].ToString();
                    k++;
                    progressBar1.Value++;
                }
            }
            w.ActiveDocument.Tables[15].Rows[k].Delete();
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in kursBS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString())
                {
                    w.ActiveDocument.Tables[16].Rows.Add();
                    w.ActiveDocument.Tables[16].Rows[k].Cells[1].Range.Text = (k - 1).ToString();
                    w.ActiveDocument.Tables[16].Rows[k].Cells[2].Range.Text = row["disc"].ToString();
                    w.ActiveDocument.Tables[16].Rows[k].Cells[3].Range.Text = row["sem"].ToString();
                    w.ActiveDocument.Tables[16].Rows[k].Cells[3].Range.Text = row["tema"].ToString();
                    w.ActiveDocument.Tables[16].Rows[k].Cells[3].Range.Text = row["ocenka"].ToString();
                    k++;
                }
            }
            w.ActiveDocument.Tables[16].Rows[k].Delete();
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in praktBS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString())
                {
                    w.ActiveDocument.Tables[17].Rows.Add();
                    w.ActiveDocument.Tables[17].Rows[k].Cells[1].Range.Text = (k - 1).ToString();
                    w.ActiveDocument.Tables[17].Rows[k].Cells[2].Range.Text = row["vid"].ToString();
                    w.ActiveDocument.Tables[17].Rows[k].Cells[3].Range.Text = row["baza"].ToString();
                    w.ActiveDocument.Tables[17].Rows[k].Cells[4].Range.Text = row["dolj"].ToString();

                    if (row["dateOt"].ToString() != "" && row["dateDo"].ToString() != "")
                        w.ActiveDocument.Tables[17].Rows[k].Cells[5].Range.Text =
                            Convert.ToDateTime(row["dateOt"]).ToString("dd.MM.yyyy")
                            + " - " +
                            Convert.ToDateTime(row["dateDo"]).ToString("dd.MM.yyyy");

                    w.ActiveDocument.Tables[17].Rows[k].Cells[6].Range.Text = row["ocenka"].ToString();
                    k++;
                }
            }
            w.ActiveDocument.Tables[17].Rows[k].Delete();
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in kvalRabBS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString())
                {
                    w.ActiveDocument.Tables[18].Rows.Add();
                    w.ActiveDocument.Tables[18].Rows[k].Cells[1].Range.Text = (k - 1).ToString();
                    w.ActiveDocument.Tables[18].Rows[k].Cells[2].Range.Text = row["tema"].ToString();
                    w.ActiveDocument.Tables[18].Rows[k].Cells[3].Range.Text = row["ocenka"].ToString();
                    k++;
                }
            }
            w.ActiveDocument.Tables[18].Rows[k].Delete();
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in nauRab2BS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString())
                {
                    w.ActiveDocument.Tables[19].Rows.Add().SetHeight(0, 0);
                    w.ActiveDocument.Tables[19].Rows[k].Cells[1].Range.Text = (k - 1).ToString();
                    w.ActiveDocument.Tables[19].Rows[k].Cells[2].Range.Text = row["tip"].ToString();
                    w.ActiveDocument.Tables[19].Rows[k].Cells[3].Range.Text = row["uroven"].ToString();
                    w.ActiveDocument.Tables[19].Rows[k].Cells[4].Range.Text = row["tema"].ToString();
                    w.ActiveDocument.Tables[19].Rows[k].Cells[6].Range.Text = row["forma"].ToString();
                    w.ActiveDocument.Tables[19].Rows[k].Cells[7].Range.Text = row["rez"].ToString();

                    if (row["dateDo"].ToString() == "01.01.2000 0:00:00" || row["dateDo"].ToString() == "")
                    {
                        if (row["dateOt"].ToString() != "")
                        {
                            w.ActiveDocument.Tables[19].Rows[k].Cells[5].Range.Text =
                                Convert.ToDateTime(row["dateOt"]).ToString("dd.MM.yyyy")
                                + " " +
                                row["mestoProv"].ToString();
                        }
                    }
                    else if (row["dateDo"].ToString() != "01.01.2000 0:00:00" && row["dateDo"].ToString() != "") 
                    {
                        if (row["dateOt"].ToString() != "")
	                    {
                            w.ActiveDocument.Tables[19].Rows[k].Cells[5].Range.Text =
                                Convert.ToDateTime(row["dateOt"]).ToString("dd.MM.yyyy")
                                + " - " +
                                Convert.ToDateTime(row["dateDo"]).ToString("dd.MM.yyyy")
                                + " " +
                                row["mestoProv"].ToString();
	                    }
                    }
                    k++;
                }
            }
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in vneRab3BS)
            {

                if (row["idStud"].ToString() == curStud["id"].ToString())
                {
                    w.ActiveDocument.Tables[20].Rows.Add().SetHeight(0, 0);
                    w.ActiveDocument.Tables[20].Rows[k].Cells[1].Range.Text = (k - 1).ToString();
                    w.ActiveDocument.Tables[20].Rows[k].Cells[2].Range.Text = row["naz"].ToString();
                    w.ActiveDocument.Tables[20].Rows[k].Cells[3].Range.Text = row["forma"].ToString();

                    w.ActiveDocument.Tables[20].Rows[k].Cells[5].Range.Text = row["rez"].ToString();

                    if (row["dateDo"].ToString() == "01.01.2000 0:00:00" || row["dateDo"].ToString() == "")
                    {
                        if (row["dateOt"].ToString() != "")
                        {
                            w.ActiveDocument.Tables[20].Rows[k].Cells[4].Range.Text =
                                Convert.ToDateTime(row["dateOt"]).ToString("dd.MM.yyyy");
                        }
                    }
                    else if (row["dateDo"].ToString() != "01.01.2000 0:00:00" && row["dateDo"].ToString() != "")
                    {
                        if (row["dateOt"].ToString() != "")
                        {
                            w.ActiveDocument.Tables[20].Rows[k].Cells[4].Range.Text =
                                Convert.ToDateTime(row["dateOt"]).ToString("dd.MM.yyyy")
                                + " - " +
                                Convert.ToDateTime(row["dateDo"]).ToString("dd.MM.yyyy");
                        }
                    }
                    k++;
                }
            }
            /////////////////////////////////////////
            k = 2;
            foreach (DataRowView row in dopRab4BS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString())
                {
                    w.ActiveDocument.Tables[21].Rows.Add().SetHeight(0, 0);
                    w.ActiveDocument.Tables[21].Rows[k].Cells[1].Range.Text = (k - 1).ToString();
                    w.ActiveDocument.Tables[21].Rows[k].Cells[2].Range.Text = row["uchir"].ToString();
                    w.ActiveDocument.Tables[21].Rows[k].Cells[3].Range.Text = row["napr"].ToString();
                    w.ActiveDocument.Tables[21].Rows[k].Cells[4].Range.Text = row["oby"].ToString();
                    w.ActiveDocument.Tables[21].Rows[k].Cells[5].Range.Text = row["rez"].ToString();
                    k++;
                }
            }
            /////////////////////////////////////////
            string FIO ="ст."+curStud["fam"].ToString()+curStud["name"].ToString()+ curStud["otch"].ToString();
            if (Directory.Exists(nameFolder + "\\" + FIO))
            {
                Directory.Delete(nameFolder + "\\" + FIO,true);
            }
            Directory.CreateDirectory(nameFolder + "\\" + FIO);
            w.ActiveDocument.SaveAs(nameFolder + "\\" + FIO + "\\" + FIO + ".doc");

            string folderSavePriloj;

            folderSavePriloj = nameFolder + "\\" + FIO + "\\Приложения\\Научно-исследовательская деятельность\\";
            Directory.CreateDirectory(folderSavePriloj);

            folderSavePriloj = nameFolder + "\\" + FIO + "\\Приложения\\Внеаудиторная деятельность\\";
            Directory.CreateDirectory(folderSavePriloj);

            folderSavePriloj = nameFolder + "\\" + FIO + "\\Приложения\\Дополнительное образование, самообразование\\";
            Directory.CreateDirectory(folderSavePriloj);

            progressBar1.Value++;
            progressBar1.Visible = false;
            w.ActiveDocument.Close();
            w = null;

            foreach (DataRowView row in nauRab2BS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString())
                {
                    if (row["foto"].ToString() != "")
                    {
                        blob = (byte[])row["foto"];
                        memStream = new MemoryStream(blob);
                        memStream.Write(blob, 0, blob.Length);
                        memStream.Position = 0;
                        img = Image.FromStream(memStream);
                        img.Save(folderSavePriloj + row["rez"].ToString() + row["fotoF"].ToString(), System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }

            

            foreach (DataRowView row in vneRab3BS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString())
                {
                    if (row["foto"].ToString() != "")
                    {
                        blob = (byte[])row["foto"];
                        memStream = new MemoryStream(blob);
                        memStream.Write(blob, 0, blob.Length);
                        memStream.Position = 0;
                        img = Image.FromStream(memStream);
                        img.Save(folderSavePriloj + row["rez"].ToString() + row["fotoF"].ToString(), System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }

           
            foreach (DataRowView row in dopRab4BS)
            {
                if (row["idStud"].ToString() == curStud["id"].ToString())
                {
                    if (row["foto"].ToString() != "")
                    {
                        blob = (byte[])row["foto"];
                        memStream = new MemoryStream(blob);
                        memStream.Write(blob, 0, blob.Length);
                        memStream.Position = 0;
                        img = Image.FromStream(memStream);
                        img.Save(folderSavePriloj + row["rez"].ToString() + row["fotoF"].ToString(), System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                tableAM.Connection.ConnectionString = "Data Source=" + Program.server + "," + Program.port + ";Initial Catalog=portfolio;Persist Security Info=True;User ID=" + Program.user + ";password=" + Program.password;
                this.specTA.Fill(this.portfolioDS.spec);
                this.groupTA.Fill(this.portfolioDS.group);
                this.studTA.Fill(this.portfolioDS.stud);
                this.prevObrTA.Fill(this.portfolioDS.prevObr);
                this.jazTA.Fill(this.portfolioDS.jaz);
                this.kvalRabTA.Fill(this.portfolioDS.kvalRab);
                this.dopRab4TA.Fill(this.portfolioDS.dopRab4);
                this.nauRab2TA.Fill(this.portfolioDS.nauRab2);
                this.vneRab3TA.Fill(this.portfolioDS.vneRab3);
                this.prikazPerevodTA.Fill(this.portfolioDS.prikazPerevod);
                this.prikazStudTA.Fill(this.portfolioDS.prikazStud);
                this.praktTA.Fill(this.portfolioDS.prakt);
                this.prPraktTA.Fill(this.portfolioDS.prPrakt);
                this.kursTA.Fill(this.portfolioDS.kurs);
                this.prKursTA.Fill(this.portfolioDS.prKurs);
                this.discTA.Fill(this.portfolioDS.disc);
                this.prDiscTA.Fill(this.portfolioDS.prDisc);
                Properties.Settings.Default.Save();
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Неверные идентификационные данные", "Ошибка входа", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
            label16.Text = "Количество студентов: "+studBS.Count.ToString();
            discBS.Filter = "sem = '1'";
        }

        private void general_KeyUp(object sender, KeyEventArgs e)
        {
                if (e.KeyCode == Keys.Return)
                {
                    dopRab4BS_CurrentChanged(sender, e);
                    specBS_CurrentChanged(sender, e);
                    groupBS_CurrentChanged(sender, e);
                    studBS_CurrentChanged(sender, e);
                    jazBS_CurrentChanged(sender, e);
                    prevObrBS_CurrentChanged(sender, e);
                    nauRab2BS_CurrentChanged(sender, e);
                    vneRab3BS_CurrentChanged(sender, e);
                    prikazStudBS_CurrentChanged(sender, e);
                    kvalRabBS_CurrentChanged(sender, e);
                    discBS_CurrentChanged(sender, e);
                }
        }
    
        #region bt
        
        private void button1_Click(object sender, EventArgs e)
        {
            specBS.AddNew();
            button1.Enabled = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (specBS.Count == 0)
            {
                MessageBox.Show("Список специальностей пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show(
                (specBS.Current as DataRowView)["naz"].ToString(),
                "Вы точно хотите удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                specBS.RemoveCurrent();
                specTA.Update(portfolioDS.spec);

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (specBS.Count==0)
            {
                MessageBox.Show("Список специальностей пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            groupBS.AddNew();
            button3.Enabled = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (groupBS.Count == 0)
            {
                MessageBox.Show("Список групп пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show(
                (groupBS.Current as DataRowView)["naz"].ToString(),
                "Вы точно хотите удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                groupBS.RemoveCurrent();
                groupTA.Update(portfolioDS.group);

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (groupBS.Count == 0)
            {
                MessageBox.Show("Список групп пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Perevod perevod = new Perevod();
            perevod.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (groupBS.Count == 0)
            {
                MessageBox.Show("Список групп пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Predm predm = new Predm();
            predm.ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (groupBS.Count == 0)
            {
                MessageBox.Show("Список групп пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            pos = studBS.Position;
            studBS.AddNew();
            Profile profile = new Profile();
            profile.ShowDialog();

        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (studBS.Count == 0)
            {
                MessageBox.Show("Список студентов пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show(
                (studBS.Current as DataRowView)["fam"].ToString() + " " +
                (studBS.Current as DataRowView)["name"].ToString() + " " +
                (studBS.Current as DataRowView)["otch"].ToString(),
                "Вы точно хотите удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                studBS.RemoveCurrent();
                studTA.Update(portfolioDS.stud);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (kvalRabBS.Count == 0)
            {
                MessageBox.Show("Список работ пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show(
            (kvalRabBS.Current as DataRowView)["tema"].ToString(),
            "Вы точно хотите удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                kvalRabBS.RemoveCurrent();
                kvalRabTA.Update(portfolioDS.kvalRab);

            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (studBS.Count == 0)
            {
                MessageBox.Show("Список студентов пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            kvalRabBS.AddNew();
            button10.Enabled = false;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (studBS.Count == 0)
            {
                MessageBox.Show("Список студентов пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            jazBS.AddNew();
            button11.Enabled = false;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (jazBS.Count == 0)
            {
                MessageBox.Show("Список языков пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show(
            (jazBS.Current as DataRowView)["nameJaz"].ToString(),
            "Вы точно хотите удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                jazBS.RemoveCurrent();
                jazTA.Update(portfolioDS.jaz);

            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (studBS.Count == 0)
            {
                MessageBox.Show("Список студентов пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            prevObrBS.AddNew();
            button13.Enabled = false;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            if (prevObrBS.Count == 0)
            {
                MessageBox.Show("Список учереждений пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show(
                        (prevObrBS.Current as DataRowView)["ucher"].ToString(),
                        "Вы точно хотите удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                prevObrBS.RemoveCurrent();
                prevObrTA.Update(portfolioDS.prevObr);

            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            if (nauRab2BS.Count == 0)
            {
                MessageBox.Show("Список достижений пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show(
                        (nauRab2BS.Current as DataRowView)["tip"].ToString(),
                        "Вы точно хотите удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                nauRab2BS.RemoveCurrent();
                nauRab2TA.Update(portfolioDS.nauRab2);

            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (studBS.Count == 0)
            {
                MessageBox.Show("Список студентов пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            pos = nauRab2BS.Position;
            nauRab2BS.AddNew();
            NauRab nauRab = new NauRab();
            nauRab.ShowDialog();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (vneRab3BS.Count == 0)
            {
                MessageBox.Show("Список достижений пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show(
                        (vneRab3BS.Current as DataRowView)["naz"].ToString(),
                        "Вы точно хотите удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                vneRab3BS.RemoveCurrent();
                vneRab3TA.Update(portfolioDS.vneRab3);

            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            if (studBS.Count == 0)
            {
                MessageBox.Show("Список студентов пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            pos = vneRab3BS.Position;
            vneRab3BS.AddNew();
            VneRab vneRab = new VneRab();
            vneRab.ShowDialog();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            if (dopRab4BS.Count == 0)
            {
                MessageBox.Show("Список достижений пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show(
                        (dopRab4BS.Current as DataRowView)["uchir"].ToString(),
                        "Вы точно хотите удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                dopRab4BS.RemoveCurrent();
                dopRab4TA.Update(portfolioDS.dopRab4);

            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            
            if (studBS.Count == 0)
            {
                MessageBox.Show("Список студентов пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            pos = dopRab4BS.Position;
            dopRab4BS.AddNew();
            DopRab dopRab = new DopRab();
            dopRab.ShowDialog();
        }

        private void button21_Click(object sender, EventArgs e)
        {
            if (groupBS.Count == 0)
            {
                MessageBox.Show("Список групп пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }


            if (folderBrowserDialog1.ShowDialog() != DialogResult.OK)return;

            string nameFolder = folderBrowserDialog1.SelectedPath;

            show(nameFolder, (groupBS.Current as DataRowView), (studBS.Current as DataRowView));

        }

        private void button22_Click(object sender, EventArgs e)
        {
            if (groupBS.Count == 0)
            {
                MessageBox.Show("Список групп пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (folderBrowserDialog1.ShowDialog() != DialogResult.OK)return;

            string nameFolder = folderBrowserDialog1.SelectedPath + "\\гр." + (groupBS.Current as DataRowView)["naz"].ToString();
            Directory.CreateDirectory(nameFolder);

            foreach (DataRowView row in studBS)
            {
                if (row["idGroup"].ToString() == (groupBS.Current as DataRowView)["id"].ToString())
                {
                    show(nameFolder, (groupBS.Current as DataRowView), row);
                }
            }
        }
        
        private void button23_Click(object sender, EventArgs e)
        {
            if (prikazStudBS.Count == 0)
            {
                MessageBox.Show("Список приказов пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show(
                (prikazStudBS.Current as DataRowView)["nomer"].ToString(),
                "Вы точно хотите удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                prikazStudBS.RemoveCurrent();
                prikazStudTA.Update(portfolioDS.prikazStud);
            }
        }
        
        private void button24_Click(object sender, EventArgs e)
        {
            if (studBS.Count == 0)
            {
                MessageBox.Show("Список студентов Пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            prikazStudBS.AddNew();
            button24.Enabled = false;

        }
        
        #endregion
        
        private void specBS_CurrentChanged(object sender, EventArgs e)
        {
            if (portfolioDS.spec.GetChanges() != null)
            {
                Validate();
                specBS.EndEdit();
                specTA.Update(portfolioDS.spec);
            }
            button1.Enabled = true;
        }

        private void groupBS_CurrentChanged(object sender, EventArgs e)
        {
            if (portfolioDS.group.GetChanges() != null)
            {
                Validate();
                groupBS.EndEdit();
                groupTA.Update(portfolioDS.group);
            }
            button3.Enabled = true;
            
        }

        private void studBS_CurrentChanged(object sender, EventArgs e)
        {
            label16.Text = "Количество студентов: " + studBS.Count.ToString();
            if (portfolioDS.stud.GetChanges() != null)
            {
                Validate();
                studBS.EndEdit();
                studTA.Update(portfolioDS.stud);
            }
            button7.Enabled = true;
        }

        private void prevObrBS_CurrentChanged(object sender, EventArgs e)
        {
            if (portfolioDS.prevObr.GetChanges() != null)
            {
                Validate();
                prevObrBS.EndEdit();
                prevObrTA.Update(portfolioDS.prevObr);
            }
            button13.Enabled = true;
        }

        private void nauRab2BS_CurrentChanged(object sender, EventArgs e)
        {
            if (portfolioDS.nauRab2.GetChanges() != null)
            {
                Validate();
                nauRab2BS.EndEdit();
                nauRab2TA.Update(portfolioDS.nauRab2);
            }
            button16.Enabled = true;
        }

        private void vneRab3BS_CurrentChanged(object sender, EventArgs e)
        {
            if (portfolioDS.vneRab3.GetChanges() != null)
            {
                Validate();
                vneRab3BS.EndEdit();
                vneRab3TA.Update(portfolioDS.vneRab3);
            }
            button18.Enabled = true;
        }

        private void dopRab4BS_CurrentChanged(object sender, EventArgs e)
        {
            if (portfolioDS.dopRab4.GetChanges() != null)
            {
                Validate();
                dopRab4BS.EndEdit();
                dopRab4TA.Update(portfolioDS.dopRab4);
            }
            button20.Enabled = true;
        }

        private void prikazStudBS_CurrentChanged(object sender, EventArgs e)
        {
            if (portfolioDS.prikazStud.GetChanges() != null)
            {
                Validate();
                prikazStudBS.EndEdit();
                prikazStudTA.Update(portfolioDS.prikazStud);
            }
            button24.Enabled = true;
        }
        
        private void kvalRabBS_CurrentChanged(object sender, EventArgs e)
        {
            if (portfolioDS.kvalRab.GetChanges() != null)
            {
                Validate();
                kvalRabBS.EndEdit();
                kvalRabTA.Update(portfolioDS.kvalRab);
            }
            button10.Enabled = true;
        }

        public void prPraktBS_CurrentChanged(object sender, EventArgs e)
        {
            if (portfolioDS.prPrakt.GetChanges() != null)
            {
                Validate();
                prPraktBS.EndEdit();
                prPraktTA.Update(portfolioDS.prPrakt);
            }
        }

        private void jazBS_CurrentChanged(object sender, EventArgs e)
        {
            if (portfolioDS.jaz.GetChanges() != null)
            {
                Validate();
                jazBS.EndEdit();
                jazTA.Update(portfolioDS.jaz);
            }
            button11.Enabled = true;
        }

        public void prKursBS_CurrentChanged(object sender, EventArgs e)
        {
            if (portfolioDS.prKurs.GetChanges() != null)
            {
                Validate();
                prKursBS.EndEdit();
                prKursTA.Update(portfolioDS.prKurs);
            }
        }

        public void prDiscBS_CurrentChanged(object sender, EventArgs e)
        {
            if (portfolioDS.prDisc.GetChanges() != null)
            {
                Validate();
                prDiscBS.EndEdit();
                prDiscTA.Update(portfolioDS.prDisc);
            }
        }

        private void discBS_CurrentChanged(object sender, EventArgs e)
        {
            if (portfolioDS.disc.GetChanges() != null)
            {
                Validate();
                discBS.EndEdit();
                discTA.Update(portfolioDS.disc);
            }
        }
        
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                Program.general.discBS.RemoveFilter();
                Program.general.discBS.Filter = "sem = '1'";
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                Program.general.discBS.RemoveFilter();
                Program.general.discBS.Filter = "sem = '2'";
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                Program.general.discBS.RemoveFilter();
                Program.general.discBS.Filter = "sem = '3'";
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked)
            {
                Program.general.discBS.RemoveFilter();
                Program.general.discBS.Filter = "sem = '4'";
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked)
            {
                Program.general.discBS.RemoveFilter();
                Program.general.discBS.Filter = "sem = '5'";
            }

        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton6.Checked)
            {
               Program.general.discBS.RemoveFilter();
               Program.general.discBS.Filter = "sem = '6'";
            }
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton7.Checked)
            {
                Program.general.discBS.RemoveFilter();
                Program.general.discBS.Filter = "sem = '7'";
            }
        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton8.Checked)
            {
                Program.general.discBS.RemoveFilter();
                Program.general.discBS.Filter = "sem = '8'";
            }
        }
        
        private void dataGridView3_DoubleClick(object sender, EventArgs e)
        {
            pos = studBS.Position;
            Profile profile = new Profile();
            profile.ShowDialog();
        }

        private void dataGridView8_DoubleClick(object sender, EventArgs e)
        {
            pos = nauRab2BS.Position;
            NauRab nauRab = new NauRab();
            nauRab.ShowDialog();
        }

        private void dataGridView9_DoubleClick(object sender, EventArgs e)
        {
            pos = vneRab3BS.Position;
            VneRab vneRab = new VneRab();
            vneRab.ShowDialog();
        }

        private void dataGridView10_DoubleClick(object sender, EventArgs e)
        {
            pos = dopRab4BS.Position;
            DopRab dopRab4 = new DopRab();
            dopRab4.ShowDialog();
        }

    }

}
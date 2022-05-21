using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace diplom
{
    public partial class Predm : Form
    {
        public Predm()
        {
            InitializeComponent();
        }

        private void Predm_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Program.general.prKursBS;
            dataGridView2.DataSource = Program.general.prPraktBS;
            dataGridView3.DataSource = Program.general.prDiscBS;

            Program.general.prDiscBS.Filter = "sem = '1'";
        }

        private void Predm_KeyUp(object sender, KeyEventArgs e)
        {
            prPraktBS_CurrentChanged(sender, e);
            prKursBS_CurrentChanged(sender, e);
            if (e.KeyCode == Keys.Enter) prDiscBS_CurrentChanged(sender, e);
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            Program.general.discBS.RemoveFilter();
            Program.general.prDiscBS.RemoveFilter();

            Program.general.Validate();
            Program.general.prPraktBS.EndEdit();
            Program.general.prKursBS.EndEdit();
            Program.general.prDiscBS.EndEdit();

            Program.general.addDisciplPrakt();
            Program.general.delDisciplPrakt();

            Program.general.addDisciplKurs();
            Program.general.delDisciplKurs();

            Program.general.addDisciplDisc();
            Program.general.delDisciplDisc();

            Program.general.prPraktTA.Update(Program.general.portfolioDS.prPrakt);
            Program.general.prKursTA.Update(Program.general.portfolioDS.prKurs);
            Program.general.prDiscTA.Update(Program.general.portfolioDS.prDisc);

            Program.general.praktTA.Update(Program.general.portfolioDS.prakt);
            Program.general.kursTA.Update(Program.general.portfolioDS.kurs);
            Program.general.discTA.Update(Program.general.portfolioDS.disc);

            Program.general.radioButton1.Checked = true;
            Program.general.discBS.Filter = "sem = '1'";
            Close();
            Dispose();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Program.general.prPraktBS.CancelEdit();
            Program.general.prKursBS.CancelEdit();
            Close();
            Dispose();
        }
        
        private void button3_Click(object sender, EventArgs e)
        {
            Program.general.prKursBS.AddNew();
            button3.Enabled = false;
            button4.Enabled = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (Program.general.prKursBS.Count == 0)
            {
                MessageBox.Show("Список дисциплин пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show(
                        (Program.general.prKursBS.Current as DataRowView)["disc"].ToString(),
                        "Вы точно хотите удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Program.general.prKursBS.RemoveCurrent();
                Program.general.prKursTA.Update(Program.general.portfolioDS.prKurs);

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Program.general.prPraktBS.AddNew();
            button5.Enabled = false;
            button6.Enabled = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (Program.general.prPraktBS.Count == 0)
            {
                MessageBox.Show("Список дисциплин пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show(
            (Program.general.prPraktBS.Current as DataRowView)["vid"].ToString(),
            "Вы точно хотите удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Program.general.prPraktBS.RemoveCurrent();
                Program.general.prPraktTA.Update(Program.general.portfolioDS.prPrakt);

            }
        }
        
        private void button7_Click(object sender, EventArgs e)
        {
            Program.general.prDiscBS.AddNew();

            if (radioButton1.Checked)
	        {
                ((DataRowView)Program.general.prDiscBS.Current)["sem"] = "1";
	        }
            
            if (radioButton2.Checked)
	        {
                ((DataRowView)Program.general.prDiscBS.Current)["sem"] = "2";
	        }
            
            if (radioButton3.Checked)
            {
                ((DataRowView)Program.general.prDiscBS.Current)["sem"] = "3";
            }
            
            if (radioButton4.Checked)
            {
                ((DataRowView)Program.general.prDiscBS.Current)["sem"] = "4";
            }
            
            if (radioButton5.Checked)
            {
                ((DataRowView)Program.general.prDiscBS.Current)["sem"] = "5";
            }
            
            if (radioButton6.Checked)
            {
                ((DataRowView)Program.general.prDiscBS.Current)["sem"] = "6";
            }
           
            if (radioButton7.Checked)
            {
                ((DataRowView)Program.general.prDiscBS.Current)["sem"] = "7";
            }
            
            if (radioButton8.Checked)
            {
                ((DataRowView)Program.general.prDiscBS.Current)["sem"] = "8";
            }

            Validate();
            Program.general.prDiscBS.EndEdit();
            Program.general.prDiscTA.Update(Program.general.portfolioDS.prDisc);
            dataGridView3.DataSource = Program.general.prDiscBS;
            button7.Enabled = false;
            button8.Enabled = false;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (dataGridView3.Rows.Count==0)
            {
                MessageBox.Show("Список дисциплин пуст", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Program.general.prDiscBS.CancelEdit();
            if (MessageBox.Show(
            (Program.general.prDiscBS.Current as DataRowView)["disc"].ToString(),
            "Вы точно хотите удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Program.general.prDiscBS.RemoveCurrent();
                Program.general.prDiscTA.Update(Program.general.portfolioDS.prDisc);

            }
            button7.Enabled = true;
            button8.Enabled = true;
        }

        private void prPraktBS_CurrentChanged(object sender, EventArgs e)
        {
            Program.general.prPraktBS_CurrentChanged(sender, e);
            button5.Enabled = true;
            button6.Enabled = true; 
        }

        private void prKursBS_CurrentChanged(object sender, EventArgs e)
        {
            Program.general.prKursBS_CurrentChanged(sender, e);
            button3.Enabled = true;
            button4.Enabled = true; 
        }

        private void prDiscBS_CurrentChanged(object sender, EventArgs e)
        {
            Program.general.prDiscBS_CurrentChanged(sender, e);
            button7.Enabled = true;
            button8.Enabled = true;
        }

        #region rb
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                Program.general.prDiscBS.Filter = "sem = '1'";
                //Program.general.prDiscTA.Update(Program.general.portfolioDS.prDisc);
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                Program.general.prDiscBS.Filter = "sem = '2'";
                //Program.general.prDiscTA.Update(Program.general.portfolioDS.prDisc);
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                Program.general.prDiscBS.Filter = "sem = '3'";
                //Program.general.prDiscTA.Update(Program.general.portfolioDS.prDisc);
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked)
            {
                Program.general.prDiscBS.Filter = "sem = '4'";
                //Program.general.prDiscTA.Update(Program.general.portfolioDS.prDisc);
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked)
            {
                Program.general.prDiscBS.Filter = "sem = '5'";
                //Program.general.prDiscTA.Update(Program.general.portfolioDS.prDisc);
            }
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton6.Checked)
            {
                Program.general.prDiscBS.Filter = "sem = '6'";
                //Program.general.prDiscTA.Update(Program.general.portfolioDS.prDisc);
            }
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton7.Checked)
            {
                Program.general.prDiscBS.Filter = "sem = '7'";
                //Program.general.prDiscTA.Update(Program.general.portfolioDS.prDisc);
            }
        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton8.Checked)
            {
                Program.general.prDiscBS.Filter = "sem = '8'";
                //Program.general.prDiscTA.Update(Program.general.portfolioDS.prDisc);
            }
        }

        #endregion

        
       
        
        

        
        

    }
}

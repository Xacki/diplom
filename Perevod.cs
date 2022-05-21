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
    public partial class Perevod : Form
    {
        public Perevod()
        {
            InitializeComponent();
        }
       
        private void Perevod_Load(object sender, EventArgs e)
        {
            this.prikazPerevodTA.Fill(this.portfolioDataSet.prikazPerevod);
            Perevod.ActiveForm.Text += (Program.general.groupBS.Current as DataRowView)["naz"].ToString();
            dataGridView1.DataSource = Program.general.prikazPerevodBS;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Program.general.prikazPerevodBS.CancelEdit();
            Close();
            Dispose();
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            Program.general.Validate();
            Program.general.prikazPerevodBS.EndEdit();
            Program.general.prikazPerevodTA.Update(Program.general.portfolioDS.prikazPerevod);
            Close();
            Dispose();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Program.general.prikazPerevodBS.AddNew();
            button3.Enabled = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(
                (Program.general.prikazPerevodBS.Current as DataRowView)["nomer"].ToString(),
                "Вы точно хотите удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Program.general.prikazPerevodBS.RemoveCurrent();
                Program.general.prikazPerevodTA.Update(Program.general.portfolioDS.prikazPerevod);

            }
        }

        private void Perevod_KeyUp(object sender, KeyEventArgs e)
        {
            prikazPerevodBS_CurrentChanged(sender, e);
        }

        private void prikazPerevodBS_CurrentChanged(object sender, EventArgs e)
        {
            if (Program.general.portfolioDS.prikazPerevod.GetChanges() != null)
            {
                Validate();
                Program.general.prikazPerevodBS.EndEdit();
                Program.general.prikazPerevodTA.Update(Program.general.portfolioDS.prikazPerevod);
            }
            button3.Enabled = true;
        }










    }
}

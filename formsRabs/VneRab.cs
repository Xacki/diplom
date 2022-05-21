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
    public partial class VneRab : Form
    {
        public VneRab()
        {
            InitializeComponent();
        }

        private void VneRab_Load(object sender, EventArgs e)
        {
            textBox1.DataBindings.Add(new Binding("Text", Program.general.vneRab3BS, "naz", true));
            textBox2.DataBindings.Add(new Binding("Text", Program.general.vneRab3BS, "forma", true));
            textBox3.DataBindings.Add(new Binding("Text", Program.general.vneRab3BS, "rez", true));
            pictureBox1.DataBindings.Add(new Binding("Image", Program.general.vneRab3BS, "foto", true));

            dateTimePicker1.DataBindings.Add(new Binding("Value", Program.general.vneRab3BS, "dateOt", true));
            dateTimePicker2.DataBindings.Add(new Binding("Value", Program.general.vneRab3BS, "dateDo", true));

            if ((Program.general.vneRab3BS.Current as DataRowView)["dateDo"].ToString() == "01.01.2000 0:00:00"
                ||
                (Program.general.vneRab3BS.Current as DataRowView)["dateDo"].ToString() == ""
                )
            {
                dateTimePicker2.Enabled = false;
                checkBox1.Checked = true;
            }
            else
            {
                dateTimePicker2.Enabled = true;
                checkBox1.Checked = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                dateTimePicker2.Value = Convert.ToDateTime("01.01.2000 0:00:00");
            }
            Program.general.Validate();
            Program.general.vneRab3BS.EndEdit();
            Program.general.vneRab3TA.Update(Program.general.portfolioDS.vneRab3);
            Close();
            Dispose();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Program.general.vneRab3BS.CancelEdit();
            Close();
            Dispose();
            Program.general.vneRab3BS.Position = Program.general.pos; 
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Bitmap pict = new Bitmap(openFileDialog1.FileName);
                pictureBox1.Image = (Image)pict;
                (Program.general.vneRab3BS.Current as DataRowView)["fotoF"] = openFileDialog1.FileName.Substring(openFileDialog1.FileName.LastIndexOf('.'));
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            pictureBox1.Image = null;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                dateTimePicker2.Enabled = false;
                dateTimePicker2.Value = Convert.ToDateTime("01.01.2000 0:00:00");
            }

            if (!checkBox1.Checked)
            {
                dateTimePicker2.Enabled = true;
                dateTimePicker2.Value = DateTime.Now;
            }
        }
    }
}

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
    public partial class DopRab : Form
    {
        public DopRab()
        {
            InitializeComponent();
        }

        private void DopRab_Load(object sender, EventArgs e)
        {
            textBox1.DataBindings.Add(new Binding("Text", Program.general.dopRab4BS, "uchir", true));
            textBox2.DataBindings.Add(new Binding("Text", Program.general.dopRab4BS, "napr", true));
            textBox3.DataBindings.Add(new Binding("Text", Program.general.dopRab4BS, "oby", true));
            textBox4.DataBindings.Add(new Binding("Text", Program.general.dopRab4BS, "rez", true));
            pictureBox1.DataBindings.Add(new Binding("Image", Program.general.dopRab4BS, "foto", true));
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Program.general.Validate();
            Program.general.dopRab4BS.EndEdit();
            Program.general.dopRab4TA.Update(Program.general.portfolioDS.dopRab4);
            Close();
            Dispose();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Program.general.dopRab4BS.CancelEdit();
            Close();
            Dispose();
            Program.general.dopRab4BS.Position = Program.general.pos; 
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Bitmap pict = new Bitmap(openFileDialog1.FileName);
                pictureBox1.Image = (Image)pict;
                (Program.general.dopRab4BS.Current as DataRowView)["fotoF"] = openFileDialog1.FileName.Substring(openFileDialog1.FileName.LastIndexOf('.'));
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            pictureBox1.Image = null;
        }
    }
}

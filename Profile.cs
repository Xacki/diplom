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
    public partial class Profile : Form
    {
        public Profile()
        {
            InitializeComponent();
        }

        private void Profile_Load(object sender, EventArgs e)
        {

            textBox1.DataBindings.Add(new Binding("Text", Program.general.studBS, "fam", true));
            textBox2.DataBindings.Add(new Binding("Text", Program.general.studBS, "name", true));
            textBox3.DataBindings.Add(new Binding("Text", Program.general.studBS, "otch", true));
            textBox4.DataBindings.Add(new Binding("Text", Program.general.studBS, "tel", true));
            textBox5.DataBindings.Add(new Binding("Text", Program.general.studBS, "email", true));
            textBox6.DataBindings.Add(new Binding("Text", Program.general.studBS, "nomZachisl", true));
            dateTimePicker1.DataBindings.Add(new Binding("Value", Program.general.studBS, "dateR", true));
            dateTimePicker2.DataBindings.Add(new Binding("Value", Program.general.studBS, "dateZachisl", true));
            comboBox1.DataBindings.Add(new Binding("Text", Program.general.studBS, "urVlad", true));
            pictureBox1.DataBindings.Add(new Binding("Image", Program.general.studBS, "foto", true));

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Bitmap pict = new Bitmap(openFileDialog1.FileName);
                pictureBox1.Image = (Image)pict;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            pictureBox1.Image = null;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Program.general.discBS.RemoveFilter();

            Program.general.Validate();
            Program.general.studBS.EndEdit();
            Program.general.studTA.Update(Program.general.portfolioDS.stud);

            Program.general.addPraktStud();
            Program.general.praktTA.Update(Program.general.portfolioDS.prakt);

            Program.general.addKursStud();
            Program.general.kursTA.Update(Program.general.portfolioDS.kurs);

            Program.general.addDisciplStud();
            Program.general.discTA.Update(Program.general.portfolioDS.disc);

            Program.general.discBS.Filter = "sem = '1'";
            Program.general.radioButton1.Checked = true;
            Close();
            Dispose();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Program.general.studBS.CancelEdit();
            Close();
            Dispose();
            Program.general.studBS.Position = Program.general.pos; 
        } 
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace diplom
{

    static class Program
    {
        public static string user, password, server, port;
        public static General general;
        public static Profile profile;
        public static DopRab dopRab;
        public static NauRab nauRab;
        public static VneRab vneRab;
        public static Perevod perevod;
        public static Predm predm;
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Zast zast = new Zast();
            DateTime end = DateTime.Now + TimeSpan.FromSeconds(1);
            zast.Show();
            while (end>DateTime.Now)
            {
                Application.DoEvents();
            }
            zast.Close();
            zast.Dispose();


            Login login = new Login();
            DopRab dopRab = new DopRab();
            NauRab nauRab = new NauRab();
            Perevod fillPredm = new Perevod();
            VneRab vneRab = new VneRab();
            Predm predm = new Predm();
            login.ShowDialog();
            if (login.DialogResult != DialogResult.OK)
                Environment.Exit(0);

            Program.general = new General();
            Application.Run(general);
        }
    }
}

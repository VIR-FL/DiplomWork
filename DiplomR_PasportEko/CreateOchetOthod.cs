using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DiplomR_PasportEko
{
    public partial class CreateOchetOthod : Form
    {
        public CreateOchetOthod()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                var helper = new WordHelper("DogovirOthodov.docx");
                var items = new Dictionary<string, string>
                {
                {"<QRUCURLIZ>", textBox1.Text},
                {"<DATADOGOVOR>", maskedTextBox1.Text},
                {"<OTHOD>", textBox2.Text},
                {"<KODFKKOOTHOD>", textBox3.Text},
                {"<QPROISHOTHOD>", textBox4.Text},
                {"<QFIZXIMSOST>", textBox5.Text},
                {"<CLASSOPAS>", textBox6.Text},
                {"<QFULLORGANIZATION>", textBox7.Text},
                {"<ORGANIZATION>", textBox8.Text},
                {"<QINN>", textBox9.Text},
                {"<QOKPO>", textBox10.Text},
                {"<QOKVED>", textBox11.Text},
                {"<QADRES>", textBox12.Text},
                {"<QPOHTADRES>", textBox13.Text},
                {"<QFACTADRES>", textBox14.Text},
                };
                helper.Process(items);
                MessageBox.Show("Отчет был успешно создан и сохранен в папку Документы на вашем компьютере!", "Уведомление", MessageBoxButtons.OK);
                Close();
            }
            catch(Exception)
            {
                MessageBox.Show("Не все поля заполенены!", "Уведомление", MessageBoxButtons.OK);
            }
            
        }

        private void посмотретьГотовыйОтчетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ProsmotrOtcheta prosmotrOtcheta = new ProsmotrOtcheta();
            prosmotrOtcheta.ShowDialog();
        }
    }
}

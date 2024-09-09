using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace DiplomR_PasportEko.InsertPasport
{
    public partial class InsertKlient : Form
    {
        public static string connect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=ЭкологПаспортНИКА2.mdb";
        private OleDbConnection myconect;
        public InsertKlient()
        {
            myconect = new OleDbConnection(connect);
            myconect.Open();
            InitializeComponent();
            //comboBox1.Items.Add("Новый"); comboBox1.Items.Add("Постояный");
            //comboBox1.SelectedIndex = 0;
            myconect.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {     
            try
            {
                myconect.Open();
                OleDbCommand command = new OleDbCommand();
                int kod = Convert.ToInt32(textBox1.Text);
                string organizethion = textBox3.Text;
                string Familia = textBox2.Text;
                string telefon = Convert.ToString(maskedTextBox1.Text);
                string email = textBox6.Text;
                string tip_klienta = Convert.ToString(textBox4.Text);
                command.Connection = myconect;
                //запрос на добавление
                command.CommandText = "INSERT INTO [Клиенты] (Код_Клиента, Организация, Фамилия_и_Инициалы, Телефон, Почта, Тип_Клиента)" +
                    " VALUES('" + kod + "','" + organizethion+ "','" + Familia + "','"  + telefon + "','" + email + "','" + tip_klienta + "')";
                command.ExecuteNonQuery();
                MessageBox.Show("Данные добавлены!"); Close();
                myconect.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}

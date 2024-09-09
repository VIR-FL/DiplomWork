using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlTypes;

namespace DiplomR_PasportEko.InsertPasport
{
    public partial class InsertDeal : Form
    {
        public static string connect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=ЭкологПаспортНИКА2.mdb";
        private OleDbConnection myconect;
        public InsertDeal()
        {
            myconect = new OleDbConnection(connect);
            myconect.Open();
            InitializeComponent();
            myconect.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                myconect.Open();
                OleDbCommand command = new OleDbCommand();
                int nomaedeal = Convert.ToInt32(textBox1.Text);
                int client = Convert.ToInt32(textBox2.Text);
                int otvechsotrud = Convert.ToInt32(textBox3.Text);
                int okazuslug = Convert.ToInt32(textBox4.Text);
                int othod = Convert.ToInt32(textBox7.Text);
                int koluslug = Convert.ToInt32(textBox5.Text);
                int price = Convert.ToInt32(textBox6.Text);
                string dateokazdeal = Convert.ToString(maskedTextBox1.Text);
                command.Connection = myconect;
                //запрос на добавление
                command.CommandText = "INSERT INTO [Сделки] (Номер_Сделки, Клиент, Отвечающий_Сотрудник, Оказанная_Услуга, Отход, Количество_Услуги, Цена, Дата_Оказания_Услуги)" +
                    " VALUES('" + nomaedeal + "','" + client + "','" + otvechsotrud + "','" + okazuslug + "','"+ othod + "','" + koluslug + "','" + price + "','" + dateokazdeal + "')";
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

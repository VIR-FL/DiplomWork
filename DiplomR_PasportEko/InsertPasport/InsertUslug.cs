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
    public partial class InsertUslug : Form
    {
        public static string connect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=ЭкологПаспортНИКА2.mdb";
        private OleDbConnection myconect;
        public InsertUslug()
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
                int kod = Convert.ToInt32(textBox1.Text);
                string Usluga = textBox2.Text;
                int price = Convert.ToInt32(textBox3.Text);
                string description = textBox4.Text;
                command.Connection = myconect;
                //запрос на добавление
                command.CommandText = "INSERT INTO [Услуги] (Код_Услуги, Услуга, Описание_Услуги, Цена)" +
                    " VALUES('" + kod + "','" + Usluga + "','" + description + "','" + price + "')";
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

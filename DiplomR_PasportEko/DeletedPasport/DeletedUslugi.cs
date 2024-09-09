using DiplomR_PasportEko.DeletedPasport.Potverdite;
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

namespace DiplomR_PasportEko.DeletedPasport
{
    public partial class DeletedUslugi : Form
    {
        public static string connect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=ЭкологПаспортНИКА2.mdb";
        private OleDbConnection myconect;
        public DeletedUslugi()
        {
            myconect = new OleDbConnection(connect);
            myconect.Open();
            InitializeComponent();
            myconect.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PotverditeUsligi potverditeUsligi = new PotverditeUsligi();
            if (potverditeUsligi.ShowDialog() == DialogResult.Yes)
            {
                try
                {
                    myconect.Open(); // открытин соединения с бд
                    int kod = Convert.ToInt32(textBox1.Text);
                    string quary1 = "DELETE FROM Услуги WHERE Код_Услуги = " + kod; // запрос на удаление записи из таблицы 
                    OleDbCommand command = new OleDbCommand(quary1, myconect);
                    command.ExecuteNonQuery();
                    MessageBox.Show("Данные о Услуге удалены!", "Уведомление", MessageBoxButtons.OK);// сообщение что запись удалена
                    myconect.Close();//закрыие соединения с бд
                    Close();
                }
                catch (Exception)
                {
                    MessageBox.Show("Введите значение!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);// окно ошибки
                }
            }
            else if (potverditeUsligi.ShowDialog() == DialogResult.No)
            {
                Close();
            }
        }
    }
}

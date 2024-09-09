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
    public partial class Avtoriz : Form
    {
        List<Avtorization> avtorizations = new List<Avtorization>
        {
            new Avtorization("freikagroicougi-8536@gmail.com","2lfaa41f"),
            new Avtorization("gouleummuyouge-4768@gmail.com","4fww1b"),
            new Avtorization("vouzepaunebre-7194@gmail.com","eg24vgh")
        };
        public Avtoriz()
        {
            InitializeComponent();
        }

       

        private void button1_Click(object sender, EventArgs e)
        {
            CenterFormPassport centerFormPassport = new CenterFormPassport();
            string login, password; int schet = 0;
            login = textBox1.Text;
            password = textBox2.Text;
            foreach (Avtorization avto in avtorizations)
            {
                if (login == avto.email && password == avto.password)
                {
                    centerFormPassport.ShowDialog();
                    break;
                }
                else if (login == "Admin" && password == "Admin")
                {
                    centerFormPassport.ShowDialog();
                    break;
                }
                schet++;
            }
            if (schet >= 3)
            {
                MessageBox.Show("Логин или пароль введены неправильно!", "Уведомление", MessageBoxButtons.OK);
            }
             textBox2.Clear(); 
        }
        private void Avtoriz_Load(object sender, EventArgs e)
        {
            textBox1.Text = "Логин";
            textBox1.ForeColor = Color.Gray;
            textBox2.Text = "Пароль";
            textBox2.ForeColor = Color.Gray;
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            if (textBox1.Text == "Логин")
            {
                textBox1.Text = null;
                textBox1.ForeColor = Color.Black;
            }

        }

        private void textBox2_Enter(object sender, EventArgs e)
        {
            if (textBox2.Text == "Пароль")
            {
                textBox2.Text = null;
                textBox2.ForeColor = Color.Black;
            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                textBox1.Text = "Логин";
                textBox1.ForeColor = Color.Gray;
            }
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            if (textBox2.Text == "")
            {
                textBox2.Text = "Пароль";
                textBox2.ForeColor = Color.Gray;
            }
        }
    }


    class Avtorization
    {
        public string email, password;

        public Avtorization(string _email, string _password)
        {
            email = _email;
            password = _password;
        }
    }
}

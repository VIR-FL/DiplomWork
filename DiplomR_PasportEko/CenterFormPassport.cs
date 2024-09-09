using DiplomR_PasportEko.DeletedPasport;
using DiplomR_PasportEko.InsertPasport;
using DiplomR_PasportEko.Redaction;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml.Serialization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace DiplomR_PasportEko
{
    public partial class CenterFormPassport : Form
    {
        public static string connect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=ЭкологПаспортНИКА2.mdb";
        private OleDbConnection myconect;
        OleDbCommand command;
        public CenterFormPassport()
        {   
            myconect = new OleDbConnection(connect);
            myconect.Open();
            InitializeComponent();
            tableotchet();
            setingsdataGridViews();
            setingscomboBox();
            myconect.Close();
        }

        private void CenterFormPassport_Load(object sender, EventArgs e)
        {
           
            // TODO: данная строка кода позволяет загрузить данные в таблицу "экологПаспортНИКА2DataSet.Отходы". При необходимости она может быть перемещена или удалена.
            this.отходыTableAdapter.Fill(this.экологПаспортНИКА2DataSet.Отходы);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "экологПаспортНИКА2DataSet.Клиенты". При необходимости она может быть перемещена или удалена.
            this.клиентыTableAdapter.Fill(this.экологПаспортНИКА2DataSet.Клиенты);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "экологПаспортНИКА2DataSet.Услуги". При необходимости она может быть перемещена или удалена.
            this.услугиTableAdapter.Fill(this.экологПаспортНИКА2DataSet.Услуги);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "экологПаспортНИКА2DataSet.Сделки". При необходимости она может быть перемещена или удалена.
            this.сделкиTableAdapter.Fill(this.экологПаспортНИКА2DataSet.Сделки);
            
        }

        public void Updates()
        {
            this.отходыTableAdapter.Update(this.экологПаспортНИКА2DataSet.Отходы);
            this.клиентыTableAdapter.Update(this.экологПаспортНИКА2DataSet.Клиенты);
            this.услугиTableAdapter.Update(экологПаспортНИКА2DataSet.Услуги);
            this.сделкиTableAdapter.Update(this.экологПаспортНИКА2DataSet.Сделки);
            this.отходыTableAdapter.Fill(this.экологПаспортНИКА2DataSet.Отходы);
            this.клиентыTableAdapter.Fill(this.экологПаспортНИКА2DataSet.Клиенты);
            this.услугиTableAdapter.Fill(this.экологПаспортНИКА2DataSet.Услуги);
            this.сделкиTableAdapter.Fill(this.экологПаспортНИКА2DataSet.Сделки);
        }

        private void tabControl1_DrawItem(object sender, DrawItemEventArgs e) //Метод который смещает в бок вкладки в Tabcontrol1 и помечание выделенной вкладки синим
        {
            Graphics g;
            string sText;
            int iX;
            float iY;

            SizeF sizeText;
            TabControl ctlTab;

            ctlTab = (TabControl)sender;

            g = e.Graphics;

            sText = ctlTab.TabPages[e.Index].Text;
            sizeText = g.MeasureString(sText, ctlTab.Font);
            iX = e.Bounds.Left + 6;
            iY = e.Bounds.Top + (e.Bounds.Height - sizeText.Height) / 2;
            g.DrawString(sText, ctlTab.Font, Brushes.Black, iX, iY);

            e.Graphics.SetClip(e.Bounds);
            string text = tabControl1.TabPages[e.Index].Text;
            SizeF sz = e.Graphics.MeasureString(text, e.Font);

            bool bSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            using (SolidBrush b = new SolidBrush(bSelected ? SystemColors.Highlight : SystemColors.Control))
                e.Graphics.FillRectangle(b, e.Bounds);

            using (SolidBrush b = new SolidBrush(bSelected ? SystemColors.HighlightText : SystemColors.ControlText))
                e.Graphics.DrawString(text, e.Font, b, e.Bounds.X + 2, e.Bounds.Y + (e.Bounds.Height - sz.Height) / 2);

            if (tabControl1.SelectedIndex == e.Index)
                e.DrawFocusRectangle();

            e.Graphics.ResetClip();
        }

        public void MyExecuteNonQuery(string SqlText)
        {
            myconect = new OleDbConnection(connect);
            myconect.Open(); // открыть источник данных
            command = myconect.CreateCommand(); // задать SQL-команду
            command.CommandText = SqlText; // задать командную строку
            command.ExecuteNonQuery(); // выполнить SQL-команду
            myconect.Close(); // закрыть источник данных
        }

        private void setingscomboBox()
        {
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox1.Items.AddRange(new string[] {"По Отходу", "По Калассу Отхода", "По Коду Отхода по ФККО"});
            comboBox1.SelectedIndex = 0; comboBox1.DropDownHeight = 60; comboBox1.DropDownWidth = 220;
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.Items.AddRange(new string[] {"По Организации", "По Фамилии", "По Типу Клиента"});
            comboBox2.SelectedIndex = 0; comboBox2.DropDownHeight = 60; comboBox2.DropDownWidth = 220;
            comboBox3.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox3.Items.AddRange(new string[] {"По Услуге", "По Описанию Услуги"});
            comboBox3.SelectedIndex = 0; comboBox3.DropDownHeight = 60; comboBox3.DropDownWidth = 220;
            comboBox4.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox4.Items.AddRange(new string[] {"По Номеру Сделки", "По Клиенту", "По Оказанной Услуге"});
            comboBox4.SelectedIndex = 0; comboBox4.DropDownHeight = 60; comboBox4.DropDownWidth = 220;
        }

        private void setingsdataGridViews()
        {
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView5.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView2.ForeColor = Color.Black;
        }

        public void tableotchet()
        {
            string query;
            query = "SELECT Сделки.Номер_Сделки, Сделки.Дата_Оказания_Услуги, Клиенты.Организация, Отходы.Отход, Отходы.Каласс_Отхода, Отходы.Код_Отходо_По_ФККО " +
                "FROM Услуги INNER JOIN (Отходы INNER JOIN (Клиенты INNER JOIN Сделки ON Клиенты.[Код_Клиента] = Сделки.[Клиент]) ON Отходы.[Код_Отхода] = Сделки.[Отход]) ON Услуги.[Код_Услуги] = Сделки.[Оказанная_Услуга] WHERE Услуги.Код_Услуги = 101 ";
            OleDbDataAdapter comand = new OleDbDataAdapter(query, myconect);
            DataTable dat = new DataTable();
            comand.Fill(dat);
            dataGridView5.DataSource = dat;
        }

        private void button13_Click_1(object sender, EventArgs e)
        {
            int index, num;
            string nomersdel, datauslug, organization, othod, classothod, kodothodsFFFKO;
            num = dataGridView5.Rows.Count;
            if (num == 1) return;
            CreateOchetOthod createOchetOthod = new CreateOchetOthod();
            index = dataGridView5.CurrentRow.Index;
            nomersdel = dataGridView5[0, index].Value.ToString();
            datauslug = dataGridView5[1, index].Value.ToString();
            organization = dataGridView5[2, index].Value.ToString();
            othod = dataGridView5[3, index].Value.ToString();
            classothod = dataGridView5[4, index].Value.ToString();
            kodothodsFFFKO = dataGridView5[5, index].Value.ToString();
            createOchetOthod.maskedTextBox1.Text = datauslug;
            createOchetOthod.textBox2.Text = othod;
            createOchetOthod.textBox3.Text = kodothodsFFFKO;
            createOchetOthod.textBox6.Text = classothod;
            createOchetOthod.textBox7.Text = organization;
            createOchetOthod.textBox8.Text = organization;
            createOchetOthod.ShowDialog();
        }

        

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string query;
            if (comboBox2.SelectedIndex == 0)
            {
                query = "SELECT * FROM Клиенты WHERE [Организация] LIKE '%" + textBox1.Text + "%' ";
            }
            else if (comboBox2.SelectedIndex == 1)
            {
                query = "SELECT * FROM Клиенты WHERE [Фамилия_и_Инициалы] LIKE '%" + textBox1.Text + "%' ";
            }
            else
            {
                query = "SELECT * FROM Клиенты WHERE [Тип_Клиента] LIKE '%" + textBox1.Text + "%' ";  
            }
            OleDbDataAdapter comand = new OleDbDataAdapter(query, myconect);
            DataTable dat = new DataTable();
            comand.Fill(dat);
            dataGridView3.DataSource = dat;
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            string query;
            if (comboBox3.SelectedIndex == 0)
            {
                query = "SELECT * FROM Услуги WHERE [Услуга] LIKE '%" + textBox2.Text + "%' ";
            }
            else
            {
                query = "SELECT * FROM Услуги WHERE [Описание_Услуги] LIKE '%" + textBox2.Text + "%' ";
            }
            OleDbDataAdapter comand = new OleDbDataAdapter(query, myconect);
            DataTable dat = new DataTable();
            comand.Fill(dat);
            dataGridView2.DataSource = dat;
        }
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            string query;
            if (comboBox4.SelectedIndex == 0)
            {
                query = "SELECT * FROM Сделки WHERE [Номер_Сделки] LIKE '%" + textBox3.Text + "%' ";
            }
            else if (comboBox4.SelectedIndex == 1)
            {
                query = "SELECT * FROM Сделки WHERE [Клиент] LIKE '%" + textBox3.Text + "%' ";
            }
            else
            {
                query = "SELECT * FROM Сделки WHERE [Оказанная_Услуга] LIKE '%" + textBox3.Text + "%' ";
            }
            OleDbDataAdapter comand = new OleDbDataAdapter(query, myconect);
            DataTable dat = new DataTable();
            comand.Fill(dat);
            dataGridView1.DataSource = dat;
        }
        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            string query;
            if (comboBox1.SelectedIndex == 0)
            {
                query = "SELECT * FROM Отходы WHERE [Отход] LIKE '%" + textBox4.Text + "%' ";
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                query = "SELECT * FROM Отходы WHERE [Каласс_Отхода] LIKE '%" + textBox4.Text + "%' ";
            }
            else
            {
                query = "SELECT * FROM Отходы WHERE [Код_Отходо_По_ФККО] LIKE '%" + textBox4.Text + "%' ";
            }
            OleDbDataAdapter comand = new OleDbDataAdapter(query, myconect);
            DataTable dat = new DataTable();
            comand.Fill(dat);
            dataGridView4.DataSource = dat;
        }
      
        private void button11_Click(object sender, EventArgs e)
        {
            int index, num;
            string SqlText = "UPDATE [Отходы] SET "; // часть запроса на редаетирование 
            string kod, othod, classothod, kodothodFKKO;
            num = dataGridView4.Rows.Count;
            if (num == 1) return;
            RedactionOthod redactionOthod = new RedactionOthod();
            // нахождения записи  по индексу в таблице
            index = dataGridView4.CurrentRow.Index;
            kod = dataGridView4[0, index].Value.ToString();
            othod = dataGridView4[1, index].Value.ToString();
            classothod = dataGridView4[2, index].Value.ToString();
            kodothodFKKO = dataGridView4[3, index].Value.ToString();
            // присвоение найденых записей в textbox чтобы из редактировать 
            redactionOthod.textBox2.Text = othod;
            redactionOthod.textBox3.Text = classothod;
            redactionOthod.textBox4.Text = kodothodFKKO;
            if (redactionOthod.ShowDialog() == DialogResult.OK)// открытие  формы для редактирования  
            {
                // присвоение отредаетироаных записей из textbox  в переменую
                othod = redactionOthod.textBox2.Text;
                classothod = redactionOthod.textBox3.Text;
                kodothodFKKO = redactionOthod.textBox4.Text;

                //выполение второй части запроса для редактирования
                SqlText += "[Отход] = '" + othod + "\', [Каласс_Отхода] = '" + int.Parse(classothod) + "\', [Код_Отходо_По_ФККО] = '" + kodothodFKKO + "\' ";
                SqlText += "WHERE [Отходы].Код_Отхода = " + int.Parse(kod);
                MyExecuteNonQuery(SqlText);
                Updates();
                MessageBox.Show("Данные изменины!", "Уведомление", MessageBoxButtons.OK);
            }
        }
        
        private void button3_Click(object sender, EventArgs e)
        {
            int index, num;
            string SqlText = "UPDATE [Услуги] SET "; // часть запроса на редаетирование 
            string kod, uslug, opisanieuslug, price;
            num = dataGridView2.Rows.Count;
            if (num == 1) return;
            RedationUslugi redationUslugi = new RedationUslugi();
            // нахождения записи  по индексу в таблице
            index = dataGridView2.CurrentRow.Index;
            kod = dataGridView2[0, index].Value.ToString();
            uslug = dataGridView2[1, index].Value.ToString();
            opisanieuslug = dataGridView2[2, index].Value.ToString();
            price = dataGridView2[3, index].Value.ToString();
            // присвоение найденых записей в textbox чтобы из редактировать 
            redationUslugi.textBox2.Text = uslug;
            redationUslugi.textBox4.Text = opisanieuslug;
            redationUslugi.textBox3.Text = price;
            if (redationUslugi.ShowDialog() == DialogResult.OK)// открытие  формы для редактирования  
            {
                // присвоение отредаетироаных записей из textbox  в переменую
                uslug = redationUslugi.textBox2.Text;
                opisanieuslug = redationUslugi.textBox4.Text;
                price = redationUslugi.textBox3.Text;

                //выполение второй части запроса для редактирования
                SqlText += "[Услуга] = '" + uslug + "\', [Описание_Услуги] = '" + opisanieuslug + "\', [Цена] = '" + int.Parse(price) + "\' ";
                SqlText += "WHERE [Услуги].Код_Услуги = " + int.Parse(kod);
                MyExecuteNonQuery(SqlText);
                Updates();
                MessageBox.Show("Данные изменины!", "Уведомление", MessageBoxButtons.OK);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            int index, num;
            string SqlText = "UPDATE [Клиенты] SET ";
            string kod, orgabization, FamilaAndInizial, Number, email, tipkilent;
            num = dataGridView3.Rows.Count;
            if (num == 1) return;
            RedationKlient redationKlient = new RedationKlient();
            index = dataGridView3.CurrentRow.Index;
            kod = dataGridView3[0, index].Value.ToString();
            orgabization = dataGridView3[1, index].Value.ToString();
            FamilaAndInizial = dataGridView3[2, index].Value.ToString();
            Number = dataGridView3[3, index].Value.ToString();
            email = dataGridView3[4, index].Value.ToString();
            tipkilent = dataGridView3[5, index].Value.ToString();
            redationKlient.textBox3.Text = orgabization;
            redationKlient.textBox2.Text = FamilaAndInizial;
            redationKlient.maskedTextBox1.Text = Number;
            redationKlient.textBox6.Text = email;
            redationKlient.textBox4.Text = tipkilent;
            if (redationKlient.ShowDialog() == DialogResult.OK)// открытие  формы для редактирования  
            {
                // присвоение отредаетироаных записей из textbox  в переменую
                orgabization = redationKlient.textBox3.Text;
                FamilaAndInizial = redationKlient.textBox2.Text;
                Number = redationKlient.maskedTextBox1.Text;
                email = redationKlient.textBox6.Text;
                tipkilent = redationKlient.textBox4.Text;

                //выполение второй части запроса для редактирования
                SqlText += "[Организация] = '" + orgabization + "\', [Фамилия_и_Инициалы] = '" + FamilaAndInizial + "\', [Телефон] = '" + Number + "\', " +
                        "[Почта] = '" + email + "\', [Тип_Клиента] = \'" + tipkilent + "\' ";
                SqlText += "WHERE [Клиенты].Код_Клиента = " + int.Parse(kod);
                MyExecuteNonQuery(SqlText);
                Updates();
                MessageBox.Show("Данные изменины!", "Уведомление", MessageBoxButtons.OK);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int index, num;
            string SqlText = "UPDATE [Сделки] SET ";
            string kod, klient, otvechsotrud, okazusluga, othod, koluslug, price, dataokazuslug;
            num = dataGridView1.Rows.Count;
            if (num == 1) return;
            RedactionDeal redactionDeal = new RedactionDeal();
            index = dataGridView1.CurrentRow.Index;
            kod = dataGridView1[0, index].Value.ToString();
            klient = dataGridView1[1, index].Value.ToString();
            otvechsotrud = dataGridView1[2, index].Value.ToString();
            okazusluga = dataGridView1[3, index].Value.ToString();
            othod = dataGridView1[4, index].Value.ToString();
            koluslug = dataGridView1[5, index].Value.ToString();
            price = dataGridView1[6, index].Value.ToString();
            dataokazuslug = dataGridView1[7, index].Value.ToString();
            redactionDeal.textBox2.Text = klient;
            redactionDeal.textBox3.Text = otvechsotrud;
            redactionDeal.textBox4.Text = okazusluga;
            redactionDeal.textBox7.Text = othod;
            redactionDeal.textBox5.Text = koluslug;
            redactionDeal.textBox6.Text = price;
            redactionDeal.maskedTextBox1.Text = dataokazuslug;
            if (redactionDeal.ShowDialog() == DialogResult.OK)// открытие  формы для редактирования  
            {
                // присвоение отредаетироаных записей из textbox  в переменую
                klient = redactionDeal.textBox2.Text;
                otvechsotrud = redactionDeal.textBox3.Text;
                okazusluga = redactionDeal.textBox4.Text;
                othod = redactionDeal.textBox7.Text;
                koluslug = redactionDeal.textBox5.Text;
                price = redactionDeal.textBox6.Text;
                dataokazuslug = redactionDeal.maskedTextBox1.Text;

                //выполение второй части запроса для редактирования
                SqlText += "[Клиент] = '" + int.Parse(klient) + "\', [Отвечающий_Сотрудник] = '"
                    + int.Parse(otvechsotrud) + "\', [Оказанная_Услуга] = '" + int.Parse(okazusluga) + "\', " +
                        "[Отход] = '" + int.Parse(othod) + "\', [Количество_Услуги] = \'"
                        + int.Parse(koluslug) + "\', [Цена] = \'" + int.Parse(price) + "\', [Дата_Оказания_Услуги] = \'" + dataokazuslug + "\' ";
                SqlText += "WHERE [Сделки].Номер_Сделки = " + int.Parse(kod);
                MyExecuteNonQuery(SqlText);
                Updates();
                MessageBox.Show("Данные изменины!", "Уведомление", MessageBoxButtons.OK);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            InsertDeal insertDeal = new InsertDeal();
            insertDeal.ShowDialog();
            Updates();
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {}

        private void button1_Click(object sender, EventArgs e)
        {
            InsertUslug insertUslug = new InsertUslug();
            insertUslug.ShowDialog();
            Updates();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            InsertKlient insertKlient = new InsertKlient();
            insertKlient.ShowDialog();
            Updates();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            InsertOthod insertOthod = new InsertOthod();
            insertOthod.ShowDialog();
            Updates();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DeletedDeal deletedDeal = new DeletedDeal();
            deletedDeal.ShowDialog();
            Updates();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DeletedUslugi deletedUslugi = new DeletedUslugi();
            deletedUslugi.ShowDialog();
            Updates();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DeletedKlient deletedKlient = new DeletedKlient();
            deletedKlient.ShowDialog();
            Updates();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            DeletedOthod deletedOthod = new DeletedOthod();
            deletedOthod.ShowDialog();
            Updates();
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 aboutBox1 = new AboutBox1();
            aboutBox1.ShowDialog();
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        public void button13_Click(object sender, EventArgs e)
        {


        }

        private void настройкаПечатиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pageSetupDialog1.ShowDialog();
        }

        private void печатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (printDialog1.ShowDialog() == DialogResult.OK) printDocument1.Print();
        }
    }
}

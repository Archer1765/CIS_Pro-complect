using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace Kursovaya
{
    public partial class Authorization : Form
    {
        DB db;
        MySqlDataAdapter msda;
        MySqlCommand command;
        string id, fio, login, password, user_role;
        public string Fio 
        { 
            get { return fio; } 
            set { fio = value; } 
        }
        public string Id
        {
            get { return id; }
            set { id = value; }
        }
        public string Login
        {
            get { return login; }
            set { login = value; }
        }
        public string Pass
        {
            get { return password; }
            set { password = value; }
        }
        public string UsRole
        {
            get { return user_role; }
            set { user_role = value; }
        }
        private void textBox1_Enter(object sender, EventArgs e)
        {
            if (textBox1.Text == "login")
            {
                textBox1.Text = "";
                textBox1.ForeColor = Color.Black;
            }
        }
        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                textBox1.Text = "login";
                textBox1.ForeColor = Color.Gray;
            }
        }
        private void textBox2_Enter(object sender, EventArgs e)
        {
            if (textBox2.Text == "password")
            {
                textBox2.Text = "";
                textBox2.ForeColor = Color.Black;
            }
        }
        private void textBox2_Leave(object sender, EventArgs e)
        {
            if (textBox2.Text == "")
            {
                textBox2.Text = "password";
                textBox2.ForeColor = Color.Gray;
            }
        }
        public Authorization()
        {
            db = new DB();
            msda = new MySqlDataAdapter();
            InitializeComponent();
            
        }
        private void Form1_Load(object sender, EventArgs e)
        {

            label4.Visible = false;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            db.OpenConnection();
            try
            {
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    string procedure_name = "select id_users, FIO, login, password, user_role_id from users WHERE login ='" + textBox1.Text + "' AND password ='" + textBox2.Text + "'";
                    //MySqlConnection Connection = new MySqlConnection(procedure_name);
                    command = new MySqlCommand(procedure_name, db.GetConnection());
                    MySqlDataReader reader;
                    reader = command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            id = reader["id_users"].ToString();
                            Fio = reader["FIO"].ToString();
                            login = reader["login"].ToString();
                            password = reader["password"].ToString();
                            user_role = reader["user_role_id"].ToString();
                        }
                        //MessageBox.Show("Ваши данные найдены в базе: логин - " + login + " " + " и пароль - " + password);
                        
                        UserRole(user_role);
                        
                        reader.Close();
                    }
                    else
                    {
                        MessageBox.Show("Данные не найдены", "Information");
                        //label4.Visible = true;
                        //reader.Close();
                    }
                }
                else
                {
                    MessageBox.Show("Логин или пароль пусты", "Information");
                    //label4.Text = "Логин или пароль пусты";
                    //label4.Visible = true;
                }
            }
            catch
            {
                MessageBox.Show("Ошибка соединения", "Information");
            }
        }
        #region
        /*
        private void buttonForms_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Form2 form2 = new Form2();
            form2.ShowDialog();
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            this.Close();
            Application.Exit();
        }

        private void buttonGraphic_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Form3 form3 = new Form3();
            form3.ShowDialog();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void buttonView_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Form4 form4 = new Form4();
            form4.ShowDialog();
        }*/
        #endregion
        private void UserRole(string user_role)
        {
            #region
            /*
            if (user_role == "1")
            {
                this.Visible = false;
                db.CloseConnection();
                Administrator a = new Administrator();
                a.ShowDialog();
            }
            else
            {
                this.Visible = false;
                db.CloseConnection();
                Seller s = new Seller();
                s.ShowDialog();
            }
            */
            #endregion
            switch (user_role)
            {
                case "1": db.CloseConnection(); this.Visible = false; Administrator a = new Administrator(id, fio, login, password, user_role); a.Show(); break;
                case "2": db.CloseConnection(); this.Visible = false; Seller s = new Seller(id, fio, login, password, user_role); s.Show(); break;
            }
        }
    }
}

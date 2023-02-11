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
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Kursovaya
{
    public partial class Administrator : Form
    {
        DB db;

        Point location = new Point(205, 12);
        Size size = new Size(700, 550);
        Color c = Color.FromArgb(192, 192, 255);
        Color c2 = Color.FromArgb(240, 248, 255);

        public Administrator(string id, string fio, string login, string password, string user_role)
        {
            db = new DB();

            InitializeComponent();

            //this.BackColor = c2;

            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            panel5.Visible = false;
            panel6.Visible = false;

            panel1.Parent = this;
            panel2.Parent = this;
            panel3.Parent = this;
            panel4.Parent = this;
            panel5.Parent = this;
            panel6.Parent = this;

            panel1.Location = location;
            panel2.Location = location;
            panel3.Location = location;
            panel4.Location = location;
            panel5.Location = location;
            panel6.Location = location;

            panel1.Size = size;
            panel2.Size = size;
            panel3.Size = size;
            panel4.Size = size;
            panel5.Size = size;
            panel6.Size = size;

            panel1.BackColor = c;
            panel2.BackColor = c;
            panel3.BackColor = c;
            panel4.BackColor = c;
            panel5.BackColor = c;
            panel6.BackColor = c;

            label12.Text = fio;

            label22.Visible = false;
            label23.Visible = false;
            label24.Visible = false;
            label25.Visible = false;
            label26.Visible = false;
            label27.Visible = false;
            label28.Visible = false;
            label29.Visible = false;

            label33.Visible = false;
            label34.Visible = false;
        }
        private void Administrator_Load(object sender, EventArgs e)
        {
            this.dataGridView1.DefaultCellStyle.Font = new Font("Bookman Old Style", 10);
            this.dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Bookman Old Style", 10);

            this.dataGridView2.DefaultCellStyle.Font = new Font("Bookman Old Style", 10);
            this.dataGridView2.ColumnHeadersDefaultCellStyle.Font = new Font("Bookman Old Style", 10);

            this.dataGridView3.DefaultCellStyle.Font = new Font("Bookman Old Style", 10);
            this.dataGridView3.ColumnHeadersDefaultCellStyle.Font = new Font("Bookman Old Style", 10);

            this.dataGridView4.DefaultCellStyle.Font = new Font("Bookman Old Style", 10);
            this.dataGridView4.ColumnHeadersDefaultCellStyle.Font = new Font("Bookman Old Style", 10);

            this.dataGridView4.DefaultCellStyle.Font = new Font("Bookman Old Style", 10);
            this.dataGridView4.ColumnHeadersDefaultCellStyle.Font = new Font("Bookman Old Style", 10);

            this.dataGridView5.DefaultCellStyle.Font = new Font("Bookman Old Style", 10);
            this.dataGridView5.ColumnHeadersDefaultCellStyle.Font = new Font("Bookman Old Style", 10);

            this.dataGridView6.DefaultCellStyle.Font = new Font("Bookman Old Style", 10);
            this.dataGridView6.ColumnHeadersDefaultCellStyle.Font = new Font("Bookman Old Style", 10);

            comboBox1.Items.Add("По продажам товара за период");
            comboBox1.Items.Add("По остаткам товара за период");
            comboBox1.Items.Add("О заказанных у поставщиков товара за период");
            comboBox1.Items.Add("По продажам сотрудников за период");
            comboBox1.Items.Add("По продажам по каждому сотруднику за период");

            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;

            ShowUsers();
        }
        private void ShowUsers()
        {
            MySqlDataAdapter msda = new MySqlDataAdapter();
            string sql = "SELECT FIO FROM users WHERE user_role_id = 2;";

            DataTable dt = new DataTable();

            db.OpenConnection();
            MySqlCommand command = new MySqlCommand(sql, db.GetConnection());
            msda.SelectCommand = command;
            msda.Fill(dt);

            MySqlDataReader reader = command.ExecuteReader();
            //List<string[]> data = new List<string[]>();

            while (reader.Read())
            {
                //data.Add(new string[1]);

                //data[data.Count - 1][0] = reader[0].ToString();
                comboBox2.Items.Add(reader[0].ToString());
            }

            reader.Close();

            db.CloseConnection();
            //comboBox2.DataSource = data;
            //foreach (string[] s in data)
            //    comboBox2.Items.Add(s);
        }
        private void ShowDataUsers()
        {
            dataGridView4.Rows.Clear();
            MySqlDataAdapter msda = new MySqlDataAdapter();
            string sql = "SELECT id_users, FIO, login, password, (SELECT user_role_name FROM user_role WHERE user_role.id_user_role = users.user_role_id) FROM users;";

            DataTable dt = new DataTable();

            db.OpenConnection();
            MySqlCommand command = new MySqlCommand(sql, db.GetConnection());
            msda.SelectCommand = command;
            msda.Fill(dt);

            MySqlDataReader reader = command.ExecuteReader();
            List<string[]> data = new List<string[]>();

            while (reader.Read())
            {
                data.Add(new string[5]);

                data[data.Count - 1][0] = reader[0].ToString();
                data[data.Count - 1][1] = reader[1].ToString();
                data[data.Count - 1][2] = reader[2].ToString();
                data[data.Count - 1][3] = reader[3].ToString();
                data[data.Count - 1][4] = reader[4].ToString();
            }

            reader.Close();

            db.CloseConnection();

            foreach (string[] s in data)
                dataGridView4.Rows.Add(s);
        }
        private void buttonUsers_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            panel5.Visible = true;
            panel6.Visible = false;

            ShowDataUsers();
        }
        private void buttonAddUsers_Click(object sender, EventArgs e)
        {
            db.OpenConnection();

            string procedure_name = "insert_users";
            MySqlCommand comm_Add = new MySqlCommand(procedure_name, db.GetConnection());
            comm_Add.CommandType = CommandType.StoredProcedure;

            MySqlParameter f_i_o_param = new MySqlParameter
            {
                ParameterName = "f_i_o",
                Value = textBoxFIO.Text
            };
            MySqlParameter log_param = new MySqlParameter
            {
                ParameterName = "log",
                Value = textBoxLogin.Text
            };
            MySqlParameter pass_param = new MySqlParameter
            {
                ParameterName = "pass",
                Value = textBoxPassword.Text
            };
            MySqlParameter na_us_ro_param = new MySqlParameter
            {
                ParameterName = "na_us_ro",
                Value = textBoxUserRole.Text
            };

            comm_Add.Parameters.Add(f_i_o_param);
            comm_Add.Parameters.Add(log_param);
            comm_Add.Parameters.Add(pass_param);
            comm_Add.Parameters.Add(na_us_ro_param);

            comm_Add.ExecuteNonQuery();

            db.CloseConnection();

            ShowDataUsers();
        }
        private void buttonDeleteUsers_Click(object sender, EventArgs e)
        {
            db.OpenConnection();

            string procedure_name = "delete_users";
            MySqlCommand comm_Del = new MySqlCommand(procedure_name, db.GetConnection());
            comm_Del.CommandType = CommandType.StoredProcedure;

            MySqlParameter f_i_o_param = new MySqlParameter
            {
                ParameterName = "f_i_o",
                Value = textBoxFIO.Text
            };
            MySqlParameter log_param = new MySqlParameter
            {
                ParameterName = "log",
                Value = textBoxLogin.Text
            };
            MySqlParameter pass_param = new MySqlParameter
            {
                ParameterName = "pass",
                Value = textBoxPassword.Text
            };
            MySqlParameter na_us_ro_param = new MySqlParameter
            {
                ParameterName = "na_us_ro",
                Value = textBoxUserRole.Text
            };

            comm_Del.Parameters.Add(f_i_o_param);
            comm_Del.Parameters.Add(log_param);
            comm_Del.Parameters.Add(pass_param);
            comm_Del.Parameters.Add(na_us_ro_param);

            comm_Del.ExecuteNonQuery();

            db.CloseConnection();

            ShowDataUsers();
        }
        private void buttonUpdateUsers_Click(object sender, EventArgs e)
        {
            db.OpenConnection();

            string procedure_name = "update_users";
            MySqlCommand comm_Upd = new MySqlCommand(procedure_name, db.GetConnection());
            comm_Upd.CommandType = CommandType.StoredProcedure;

            string value = dataGridView4.CurrentRow.Cells[0].Value.ToString();

            MySqlParameter id_param = new MySqlParameter
            {
                ParameterName = "id",
                Value = value
            };
            MySqlParameter f_i_o_param = new MySqlParameter
            {
                ParameterName = "new_f_i_o",
                Value = textBoxFIO.Text
            };
            MySqlParameter log_param = new MySqlParameter
            {
                ParameterName = "new_log",
                Value = textBoxLogin.Text
            };
            MySqlParameter pass_param = new MySqlParameter
            {
                ParameterName = "new_pass",
                Value = textBoxPassword.Text
            };
            MySqlParameter na_us_ro_param = new MySqlParameter
            {
                ParameterName = "new_na_us_ro",
                Value = textBoxUserRole.Text
            };

            comm_Upd.Parameters.Add(id_param);
            comm_Upd.Parameters.Add(f_i_o_param);
            comm_Upd.Parameters.Add(log_param);
            comm_Upd.Parameters.Add(pass_param);
            comm_Upd.Parameters.Add(na_us_ro_param);

            comm_Upd.ExecuteNonQuery();

            db.CloseConnection();

            ShowDataUsers();
        }
        private void ShowDataTovar()
        {
            dataGridView1.Rows.Clear();
            MySqlDataAdapter msda = new MySqlDataAdapter();
            string sql = "SELECT id_tovar, tovar_name, price, (SELECT gruppa_tovar_name FROM gruppa_tovar WHERE gruppa_tovar.id_gruppa_tovar = tovar.gruppa_tovar_id), (SELECT provider_name FROM provider WHERE provider.id_provider = tovar.provider_id) FROM tovar;";

            DataTable dt = new DataTable();

            db.OpenConnection();
            MySqlCommand command = new MySqlCommand(sql, db.GetConnection());
            msda.SelectCommand = command;
            msda.Fill(dt);

            MySqlDataReader reader = command.ExecuteReader();
            List<string[]> data = new List<string[]>();

            while (reader.Read())
            {
                data.Add(new string[5]);

                data[data.Count - 1][0] = reader[0].ToString();
                data[data.Count - 1][1] = reader[1].ToString();
                data[data.Count - 1][2] = reader[2].ToString();
                data[data.Count - 1][3] = reader[3].ToString();
                data[data.Count - 1][4] = reader[4].ToString();
            }

            reader.Close();

            db.CloseConnection();

            foreach (string[] s in data)
                dataGridView1.Rows.Add(s);
        }
        private void buttonTovar_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            panel5.Visible = false;
            panel6.Visible = false;

            ShowDataTovar();
        }
        private void buttonAddTovar_Click(object sender, EventArgs e)
        {
            db.OpenConnection();

            string procedure_name = "insert_tovar";
            MySqlCommand comm_Add = new MySqlCommand(procedure_name, db.GetConnection());
            comm_Add.CommandType = CommandType.StoredProcedure;

            MySqlParameter t_n_param = new MySqlParameter
            {
                ParameterName = "t_n",
                Value = textBox_tovar_name.Text
            };
            MySqlParameter pr_param = new MySqlParameter
            {
                ParameterName = "pr",
                Value = Convert.ToInt32(textBox_price.Text)//цена
            };
            MySqlParameter na_gr_tov_param = new MySqlParameter
            {
                ParameterName = "na_gr_tov",
                Value = textBox_gruppa_tovar_id.Text
            };
            MySqlParameter na_prov_param = new MySqlParameter
            {
                ParameterName = "na_prov",
                Value = textBox_provider_id.Text
            };

            comm_Add.Parameters.Add(t_n_param);
            comm_Add.Parameters.Add(pr_param);
            comm_Add.Parameters.Add(na_gr_tov_param);
            comm_Add.Parameters.Add(na_prov_param);

            comm_Add.ExecuteNonQuery();

            db.CloseConnection();

            ShowDataTovar();
        }
        private void buttonDeleteTovar_Click(object sender, EventArgs e)
        {
            db.OpenConnection();

            string procedure_name = "delete_tovar";
            MySqlCommand comm_Del = new MySqlCommand(procedure_name, db.GetConnection());
            comm_Del.CommandType = CommandType.StoredProcedure;

            MySqlParameter t_n_param = new MySqlParameter
            {
                ParameterName = "t_n",
                Value = textBox_tovar_name.Text
            };
            MySqlParameter p_param = new MySqlParameter
            {
                ParameterName = "pr",
                Value = Convert.ToInt32(textBox_price.Text)//цена
            };
            MySqlParameter na_gr_tov_param = new MySqlParameter
            {
                ParameterName = "na_gr_tov",
                Value = textBox_gruppa_tovar_id.Text
            };
            MySqlParameter na_prov_param = new MySqlParameter
            {
                ParameterName = "na_prov",
                Value = textBox_provider_id.Text
            };

            comm_Del.Parameters.Add(t_n_param);
            comm_Del.Parameters.Add(p_param);
            comm_Del.Parameters.Add(na_gr_tov_param);
            comm_Del.Parameters.Add(na_prov_param);

            comm_Del.ExecuteNonQuery();

            db.CloseConnection();

            ShowDataTovar();
        }
        private void buttonUpdateTovar_Click(object sender, EventArgs e)
        {
            db.OpenConnection();

            string procedure_name = "update_tovar";
            MySqlCommand comm_Upd = new MySqlCommand(procedure_name, db.GetConnection());
            comm_Upd.CommandType = CommandType.StoredProcedure;

            string value = dataGridView1.CurrentRow.Cells[0].Value.ToString();

            MySqlParameter id_param = new MySqlParameter
            {
                ParameterName = "id",
                Value = value
            };
            MySqlParameter new_t_n_param = new MySqlParameter
            {
                ParameterName = "new_t_n",
                Value = textBox_tovar_name.Text
            };
            MySqlParameter new_p_param = new MySqlParameter
            {
                ParameterName = "new_p",
                Value = Convert.ToInt32(textBox_price.Text)//цена
            };
            MySqlParameter new_g_t_i_param = new MySqlParameter
            {
                ParameterName = "new_g_t_i",
                Value = textBox_gruppa_tovar_id.Text
            };
            MySqlParameter new_p_i_param = new MySqlParameter
            {
                ParameterName = "new_p_i",
                Value = textBox_provider_id.Text
            };

            comm_Upd.Parameters.Add(id_param);
            comm_Upd.Parameters.Add(new_t_n_param);
            comm_Upd.Parameters.Add(new_p_param);
            comm_Upd.Parameters.Add(new_g_t_i_param);
            comm_Upd.Parameters.Add(new_p_i_param);

            comm_Upd.ExecuteNonQuery();

            db.CloseConnection();

            ShowDataTovar();
        }
        private void ShowDataRashod()
        {
            dataGridView2.Rows.Clear();
            MySqlDataAdapter msda = new MySqlDataAdapter();
            string sql = "SELECT id_rashod, rashod_date, rashod_colvo, (SELECT tovar_name FROM tovar WHERE tovar.id_tovar = rashod.tovar_id), (SELECT FIO FROM users WHERE users.id_users = rashod.users_id) FROM rashod; ";

            DataTable dt = new DataTable();

            db.OpenConnection();
            MySqlCommand command = new MySqlCommand(sql, db.GetConnection());
            msda.SelectCommand = command;
            msda.Fill(dt);

            MySqlDataReader reader = command.ExecuteReader();
            List<string[]> data = new List<string[]>();

            while (reader.Read())
            {
                data.Add(new string[5]);

                data[data.Count - 1][0] = reader[0].ToString();
                data[data.Count - 1][1] = reader[1].ToString();
                data[data.Count - 1][2] = reader[2].ToString();
                data[data.Count - 1][3] = reader[3].ToString();
                data[data.Count - 1][4] = reader[4].ToString();
            }

            reader.Close();

            db.CloseConnection();

            foreach (string[] s in data)
                dataGridView2.Rows.Add(s);
        }
        private void buttonRashod_Click(object sender, EventArgs e)
        {
            panel2.Visible = true;
            panel1.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            panel5.Visible = false;
            panel6.Visible = false;

            ShowDataRashod();
        }
        private void buttonAddRashod_Click(object sender, EventArgs e)
        {
            db.OpenConnection();

            string procedure_name = "insert_rashod";
            MySqlCommand comm_Add = new MySqlCommand(procedure_name, db.GetConnection());
            comm_Add.CommandType = CommandType.StoredProcedure;

            MySqlParameter r_d_param = new MySqlParameter
            {
                ParameterName = "r_d",
                Value = textBox_date.Text
            };
            MySqlParameter r_c_param = new MySqlParameter
            {
                ParameterName = "r_c",
                Value = Convert.ToInt32(textBox_colvo.Text)//кол-во
            };
            MySqlParameter na_tov_param = new MySqlParameter
            {
                ParameterName = "na_tov",
                Value = textBox_tvar_name.Text
            };
            MySqlParameter na_us_param = new MySqlParameter
            {
                ParameterName = "na_us",
                Value = textBox_user.Text
            };

            comm_Add.Parameters.Add(r_d_param);
            comm_Add.Parameters.Add(r_c_param);
            comm_Add.Parameters.Add(na_tov_param);
            comm_Add.Parameters.Add(na_us_param);

            comm_Add.ExecuteNonQuery();

            db.CloseConnection();

            ShowDataRashod();
        }
        private void buttonDeleteRashod_Click(object sender, EventArgs e)
        {
            db.OpenConnection();

            string procedure_name = "delete_rashod";
            MySqlCommand comm_Del = new MySqlCommand(procedure_name, db.GetConnection());
            comm_Del.CommandType = CommandType.StoredProcedure;

            MySqlParameter r_d_param = new MySqlParameter
            {
                ParameterName = "r_d",
                Value = textBox_date.Text
            };
            MySqlParameter r_c_param = new MySqlParameter
            {
                ParameterName = "r_c",
                Value = Convert.ToInt32(textBox_colvo.Text)//кол-во
            };
            MySqlParameter na_tov_param = new MySqlParameter
            {
                ParameterName = "na_tov",
                Value = textBox_tvar_name.Text
            };
            MySqlParameter na_us_param = new MySqlParameter
            {
                ParameterName = "na_us",
                Value = textBox_user.Text
            };

            comm_Del.Parameters.Add(r_d_param);
            comm_Del.Parameters.Add(r_c_param);
            comm_Del.Parameters.Add(na_tov_param);
            comm_Del.Parameters.Add(na_us_param);

            comm_Del.ExecuteNonQuery();

            db.CloseConnection();

            ShowDataRashod();
        }
        private void buttonUpdateRashod_Click(object sender, EventArgs e)
        {
            db.OpenConnection();

            string procedure_name = "update_rashod";
            MySqlCommand comm_Upd = new MySqlCommand(procedure_name, db.GetConnection());
            comm_Upd.CommandType = CommandType.StoredProcedure;

            string value = dataGridView2.CurrentRow.Cells[0].Value.ToString();

            MySqlParameter id_param = new MySqlParameter
            {
                ParameterName = "id",
                Value = value
            };
            MySqlParameter new_r_d_param = new MySqlParameter
            {
                ParameterName = "new_r_d",
                Value = textBox_date.Text
            };
            MySqlParameter new_r_c_param = new MySqlParameter
            {
                ParameterName = "new_r_c",
                Value = Convert.ToInt32(textBox_colvo.Text)//цена
            };
            MySqlParameter new_tov_param = new MySqlParameter
            {
                ParameterName = "new_tov",
                Value = textBox_tvar_name.Text
            };
            MySqlParameter new_us_param = new MySqlParameter
            {
                ParameterName = "new_us",
                Value = textBox_user.Text
            };

            comm_Upd.Parameters.Add(id_param);
            comm_Upd.Parameters.Add(new_r_d_param);
            comm_Upd.Parameters.Add(new_r_c_param);
            comm_Upd.Parameters.Add(new_tov_param);
            comm_Upd.Parameters.Add(new_us_param);

            comm_Upd.ExecuteNonQuery();

            db.CloseConnection();

            ShowDataRashod();
        }
        private void ShowDataPrihod()
        {
            dataGridView3.Rows.Clear();
            MySqlDataAdapter msda = new MySqlDataAdapter();
            string sql = "SELECT id_prihod, prihod_date, prihod_colvo, (SELECT tovar_name FROM tovar WHERE tovar.id_tovar = prihod.tovar_id) FROM prihod; ";

            DataTable dt = new DataTable();

            db.OpenConnection();
            MySqlCommand command = new MySqlCommand(sql, db.GetConnection());
            msda.SelectCommand = command;
            msda.Fill(dt);

            MySqlDataReader reader = command.ExecuteReader();
            List<string[]> data = new List<string[]>();

            while (reader.Read())
            {
                data.Add(new string[4]);

                data[data.Count - 1][0] = reader[0].ToString();
                data[data.Count - 1][1] = reader[1].ToString();
                data[data.Count - 1][2] = reader[2].ToString();
                data[data.Count - 1][3] = reader[3].ToString();
            }

            reader.Close();

            db.CloseConnection();

            foreach (string[] s in data)
                dataGridView3.Rows.Add(s);
        }
        private void buttonPrihod_Click(object sender, EventArgs e)
        {
            panel3.Visible = true;
            panel1.Visible = false;
            panel2.Visible = false;
            panel4.Visible = false;
            panel5.Visible = false;
            panel6.Visible = false;

            ShowDataPrihod();
        }
        private void buttonAddPrihod_Click(object sender, EventArgs e)
        {
            db.OpenConnection();

            string procedure_name = "insert_prihod";
            MySqlCommand comm_Add = new MySqlCommand(procedure_name, db.GetConnection());
            comm_Add.CommandType = CommandType.StoredProcedure;

            MySqlParameter p_d_param = new MySqlParameter
            {
                ParameterName = "p_d",
                Value = textBox_date2.Text
            };
            MySqlParameter p_c_param = new MySqlParameter
            {
                ParameterName = "p_c",
                Value = Convert.ToInt32(textBox_colvo_prihod.Text)//кол-во
            };
            MySqlParameter na_tov_param = new MySqlParameter
            {
                ParameterName = "na_tov",
                Value = textBox_t_name.Text
            };

            comm_Add.Parameters.Add(p_d_param);
            comm_Add.Parameters.Add(p_c_param);
            comm_Add.Parameters.Add(na_tov_param);

            comm_Add.ExecuteNonQuery();

            db.CloseConnection();

            ShowDataPrihod();
        }
        private void buttonDeletePrihod_Click(object sender, EventArgs e)
        {
            db.OpenConnection();

            string procedure_name = "delete_prihod";
            MySqlCommand comm_Del = new MySqlCommand(procedure_name, db.GetConnection());
            comm_Del.CommandType = CommandType.StoredProcedure;

            MySqlParameter p_d_param = new MySqlParameter
            {
                ParameterName = "p_d",
                Value = textBox_date2.Text
            };
            MySqlParameter p_c_param = new MySqlParameter
            {
                ParameterName = "p_c",
                Value = Convert.ToInt32(textBox_colvo_prihod.Text)//кол-во
            };
            MySqlParameter na_tov_param = new MySqlParameter
            {
                ParameterName = "na_tov",
                Value = textBox_t_name.Text
            };

            comm_Del.Parameters.Add(p_d_param);
            comm_Del.Parameters.Add(p_c_param);
            comm_Del.Parameters.Add(na_tov_param);

            comm_Del.ExecuteNonQuery();

            db.CloseConnection();

            ShowDataTovar();
        }
        private void buttonUpdatePrihod_Click(object sender, EventArgs e)
        {
            db.OpenConnection();

            string procedure_name = "update_prihod";
            MySqlCommand comm_Upd = new MySqlCommand(procedure_name, db.GetConnection());
            comm_Upd.CommandType = CommandType.StoredProcedure;

            string value = dataGridView3.CurrentRow.Cells[0].Value.ToString();

            MySqlParameter id_param = new MySqlParameter
            {
                ParameterName = "id",
                Value = value
            };
            MySqlParameter new_p_d_param = new MySqlParameter
            {
                ParameterName = "new_p_d",
                Value = textBox_date2.Text
            };
            MySqlParameter new_p_c_param = new MySqlParameter
            {
                ParameterName = "new_p_c",
                Value = Convert.ToInt32(textBox_colvo_prihod.Text)//цена
            };
            MySqlParameter new_na_tov_param = new MySqlParameter
            {
                ParameterName = "new_na_tov",
                Value = textBox_t_name.Text
            };

            comm_Upd.Parameters.Add(id_param);
            comm_Upd.Parameters.Add(new_p_d_param);
            comm_Upd.Parameters.Add(new_p_c_param);
            comm_Upd.Parameters.Add(new_na_tov_param);

            comm_Upd.ExecuteNonQuery();

            db.CloseConnection();

            ShowDataPrihod();
        }
        private void report_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = true;
            panel5.Visible = false;
            panel6.Visible = false;
        }
        private void textBox_date_Enter(object sender, EventArgs e)
        {
            if (textBox_date.Text == "гггг-мм-дд")
            {
                textBox_date.Text = "";
                textBox_date.ForeColor = Color.Black;
            }
        }
        private void textBox_date_Leave(object sender, EventArgs e)
        {
            if (textBox_date.Text == "")
            {
                textBox_date.Text = "гггг-мм-дд";
                textBox_date.ForeColor = Color.Gray;
            }
        }
        private void textBox_user_Enter(object sender, EventArgs e)
        {
            if (textBox_user.Text == "Фамилия И.О.")
            {
                textBox_user.Text = "";
                textBox_user.ForeColor = Color.Black;
            }
        }
        private void textBox_user_Leave(object sender, EventArgs e)
        {
            if (textBox_user.Text == "")
            {
                textBox_user.Text = "Фамилия И.О.";
                textBox_user.ForeColor = Color.Gray;
            }
        }
        private void textBox_date2_Enter(object sender, EventArgs e)
        {
            if (textBox_date2.Text == "гггг-мм-дд")
            {
                textBox_date2.Text = "";
                textBox_date2.ForeColor = Color.Black;
            }
        }
        private void textBox_date2_Leave(object sender, EventArgs e)
        {
            if (textBox_date2.Text == "")
            {
                textBox_date2.Text = "гггг-мм-дд";
                textBox_date2.ForeColor = Color.Gray;
            }
        }
        private void button_Excel_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "По продажам товара за период")
            {
                SaveToExcel("Отчет по продажам за период " + textBox_date_ot.Text.ToString() + "-" + textBox_date_do.Text.ToString());
            }
            if (comboBox1.Text == "По остаткам товара за период")
            {
                SaveToExcel("Отчет по остаткам товара за период " + textBox_date_ot.Text.ToString() + "-" + textBox_date_do.Text.ToString());
            }
            if (comboBox1.Text == "О заказанных у поставщиков товара за период")
            {
                SaveToExcel("Отчет о заказанных у поставщиков товара за период " + textBox_date_ot.Text.ToString() + "-" + textBox_date_do.Text.ToString());
            }
            if (comboBox1.Text == "По продажам сотрудников за период")
            {
                SaveToExcel("Отчет по продажам сотрудников за период " + textBox_date_ot.Text.ToString() + "-" + textBox_date_do.Text.ToString());
            }
            if (comboBox1.Text == "По продажам по каждому сотруднику за период")
            {
                SaveToExcel("Отчет по продажам " + comboBox2.Text.ToString() + " за период " + textBox_date_ot.Text.ToString() + "-" + textBox_date_do.Text.ToString());
            }
        }
        private void textBox_date_ot_Enter(object sender, EventArgs e)
        {
            if (textBox_date_ot.Text == "гггг-мм-дд")
            {
                textBox_date_ot.Text = "";
                textBox_date_ot.ForeColor = Color.Black;
            }
        }
        private void textBox_date_ot_Leave(object sender, EventArgs e)
        {
            if (textBox_date_ot.Text == "")
            {
                textBox_date_ot.Text = "гггг-мм-дд";
                textBox_date_ot.ForeColor = Color.Gray;
            }
        }
        private void textBox_date_do_Enter(object sender, EventArgs e)
        {
            if (textBox_date_do.Text == "гггг-мм-дд")
            {
                textBox_date_do.Text = "";
                textBox_date_do.ForeColor = Color.Black;
            }
        }
        private void textBox_date_do_Leave(object sender, EventArgs e)
        {
            if (textBox_date_do.Text == "")
            {
                textBox_date_do.Text = "гггг-мм-дд";
                textBox_date_do.ForeColor = Color.Gray;
            }
        }
        private void textBoxFIO_Enter(object sender, EventArgs e)
        {
            if (textBoxFIO.Text == "Фамилия И.О.")
            {
                textBoxFIO.Text = "";
                textBoxFIO.ForeColor = Color.Black;
            }
        }
        private void textBoxFIO_Leave(object sender, EventArgs e)
        {
            if (textBoxFIO.Text == "")
            {
                textBoxFIO.Text = "Фамилия И.О.";
                textBoxFIO.ForeColor = Color.Gray;
            }
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox_tovar_name.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            textBox_price.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            textBox_gruppa_tovar_id.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            textBox_provider_id.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
        }
        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox_date.Text = dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString();
            textBox_colvo.Text = dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString();
            textBox_tvar_name.Text = dataGridView2.Rows[e.RowIndex].Cells[3].Value.ToString();
            textBox_user.Text = dataGridView2.Rows[e.RowIndex].Cells[4].Value.ToString();
        }
        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox_date2.Text = dataGridView3.Rows[e.RowIndex].Cells[1].Value.ToString();
            textBox_colvo_prihod.Text = dataGridView3.Rows[e.RowIndex].Cells[2].Value.ToString();
            textBox_t_name.Text = dataGridView3.Rows[e.RowIndex].Cells[3].Value.ToString();
        }
        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            textBoxFIO.Text = dataGridView4.Rows[e.RowIndex].Cells[1].Value.ToString();
            textBoxLogin.Text = dataGridView4.Rows[e.RowIndex].Cells[2].Value.ToString();
            textBoxPassword.Text = dataGridView4.Rows[e.RowIndex].Cells[3].Value.ToString();
            textBoxUserRole.Text = dataGridView4.Rows[e.RowIndex].Cells[4].Value.ToString();
        }
        private void buttonShowReport_Click(object sender, EventArgs e)
        {
            db.OpenConnection();
            if (comboBox1.Text == "По продажам товара за период")
            {
                dataGridView5.Rows.Clear();
                dataGridView5.ColumnCount = 5;

                dataGridView5.Columns[0].HeaderText = "Наименование товара";
                dataGridView5.Columns[1].HeaderText = "Количество";
                dataGridView5.Columns[2].HeaderText = "Цена";
                dataGridView5.Columns[3].HeaderText = "Сумма";
                dataGridView5.Columns[4].HeaderText = "Дата продажи";

                string sql1 = "SELECT t.tovar_name, r.rashod_colvo, t.price, (r.rashod_colvo * t.price), r.rashod_date FROM (titan.rashod r JOIN titan.tovar t) WHERE ((r.rashod_date >= '" + textBox_date_ot.Text + "') and(r.rashod_date <= '" + textBox_date_do.Text + "') and (r.tovar_id = t.id_tovar));";
                MySqlCommand command1 = new MySqlCommand(sql1, db.GetConnection());

                MySqlDataReader reader1 = command1.ExecuteReader();
                List<string[]> data1 = new List<string[]>();

                while (reader1.Read())
                {
                    data1.Add(new string[5]);

                    data1[data1.Count - 1][0] = reader1[0].ToString();
                    data1[data1.Count - 1][1] = reader1[1].ToString();
                    data1[data1.Count - 1][2] = reader1[2].ToString();
                    data1[data1.Count - 1][3] = reader1[3].ToString();
                    data1[data1.Count - 1][4] = reader1[4].ToString();
                }
                reader1.Close();

                db.CloseConnection();

                foreach (string[] s in data1)
                    dataGridView5.Rows.Add(s);
            }
            else if (comboBox1.Text == "По остаткам товара за период")
            {
                dataGridView5.Rows.Clear();
                dataGridView5.ColumnCount = 3;
                dataGridView5.Columns[0].HeaderText = "Наименование товара";
                dataGridView5.Columns[1].HeaderText = "Цена";
                dataGridView5.Columns[2].HeaderText = "Количество";

                string sql2 = "SELECT t.tovar_name, t.price, (p.prihod_colvo - r.rashod_colvo) FROM((titan.rashod r JOIN titan.tovar t) JOIN titan.prihod p) WHERE((r.rashod_date >= '" + textBox_date_ot.Text + "') and (r.rashod_date <= '" + textBox_date_do.Text + "') and(r.tovar_id = t.id_tovar) and(t.id_tovar = p.tovar_id));";
                MySqlCommand command2 = new MySqlCommand(sql2, db.GetConnection());

                MySqlDataReader reader2 = command2.ExecuteReader();
                List<string[]> data2 = new List<string[]>();

                while (reader2.Read())
                {
                    data2.Add(new string[5]);

                    data2[data2.Count - 1][0] = reader2[0].ToString();
                    data2[data2.Count - 1][1] = reader2[1].ToString();
                    data2[data2.Count - 1][2] = reader2[2].ToString();
                }
                reader2.Close();

                db.CloseConnection();

                foreach (string[] s in data2)
                    dataGridView5.Rows.Add(s);
            }
            else if (comboBox1.Text == "О заказанных у поставщиков товара за период")
            {
                dataGridView5.Rows.Clear();
                dataGridView5.ColumnCount = 5;

                dataGridView5.Columns[0].HeaderText = "Наименование товара";
                dataGridView5.Columns[1].HeaderText = "Количество";
                dataGridView5.Columns[2].HeaderText = "Цена";
                dataGridView5.Columns[3].HeaderText = "Поставщик";
                dataGridView5.Columns[4].HeaderText = "Дата продажи";

                string sql1 = "SELECT t.tovar_name, p.prihod_colvo, t.price, pr.provider_name, p.prihod_date FROM ((titan.prihod p JOIN titan.tovar t) JOIN titan.provider pr) WHERE ((p.prihod_date >= '" + textBox_date_ot.Text + "') and (p.prihod_date <= '" + textBox_date_do.Text + "') and (pr.id_provider = t.provider_id) and (t.id_tovar = p.tovar_id));";
                MySqlCommand command1 = new MySqlCommand(sql1, db.GetConnection());

                MySqlDataReader reader1 = command1.ExecuteReader();
                List<string[]> data1 = new List<string[]>();

                while (reader1.Read())
                {
                    data1.Add(new string[5]);

                    data1[data1.Count - 1][0] = reader1[0].ToString();
                    data1[data1.Count - 1][1] = reader1[1].ToString();
                    data1[data1.Count - 1][2] = reader1[2].ToString();
                    data1[data1.Count - 1][3] = reader1[3].ToString();
                    data1[data1.Count - 1][4] = reader1[4].ToString();
                }
                reader1.Close();

                db.CloseConnection();

                foreach (string[] s in data1)
                    dataGridView5.Rows.Add(s);
            }
            else if (comboBox1.Text == "По продажам сотрудников за период")
            {
                dataGridView5.Rows.Clear();
                dataGridView5.ColumnCount = 5;

                dataGridView5.Columns[0].HeaderText = "Сотрудник";
                dataGridView5.Columns[1].HeaderText = "Наименование товара";
                dataGridView5.Columns[2].HeaderText = "Количество";
                dataGridView5.Columns[3].HeaderText = "Сумма";
                dataGridView5.Columns[4].HeaderText = "Дата продажи";

                string sql1 = "SELECT u.FIO, t.tovar_name, r.rashod_colvo, (r.rashod_colvo * t.price), r.rashod_date FROM ((titan.users u JOIN titan.tovar t) JOIN titan.rashod r) WHERE ((r.rashod_date >= '" + textBox_date_ot.Text + "') and (r.rashod_date <= '" + textBox_date_do.Text + "') and (u.id_users = r.users_id) and (t.id_tovar = r.tovar_id));";
                MySqlCommand command1 = new MySqlCommand(sql1, db.GetConnection());

                MySqlDataReader reader1 = command1.ExecuteReader();
                List<string[]> data1 = new List<string[]>();

                while (reader1.Read())
                {
                    data1.Add(new string[5]);

                    data1[data1.Count - 1][0] = reader1[0].ToString();
                    data1[data1.Count - 1][1] = reader1[1].ToString();
                    data1[data1.Count - 1][2] = reader1[2].ToString();
                    data1[data1.Count - 1][3] = reader1[3].ToString();
                    data1[data1.Count - 1][4] = reader1[4].ToString();
                }
                reader1.Close();

                db.CloseConnection();

                foreach (string[] s in data1)
                    dataGridView5.Rows.Add(s);
            }
            //comboBox1.Items.Add("По продажам по каждому сотруднику за период");
            else //comboBox1.Items.Add("По продажам по каждому сотруднику за период");
            {
                //label18.Visible = true;
                //comboBox2.Visible = true;

                dataGridView5.Rows.Clear();
                dataGridView5.ColumnCount = 4;

                dataGridView5.Columns[0].HeaderText = "Наименование товара";
                dataGridView5.Columns[1].HeaderText = "Количество";
                dataGridView5.Columns[2].HeaderText = "Сумма";
                dataGridView5.Columns[3].HeaderText = "Дата продажи";

                string sql1 = "SELECT t.tovar_name, r.rashod_colvo, (r.rashod_colvo * t.price), r.rashod_date FROM ((titan.users u JOIN titan.tovar t) JOIN titan.rashod r) WHERE ((r.rashod_date >= '" + textBox_date_ot.Text + "') and (r.rashod_date <= '" + textBox_date_do.Text + "') and (u.FIO = '" + comboBox2.Text + "') and (u.id_users = r.users_id) and (t.id_tovar = r.tovar_id));";
                MySqlCommand command1 = new MySqlCommand(sql1, db.GetConnection());

                MySqlDataReader reader1 = command1.ExecuteReader();
                List<string[]> data1 = new List<string[]>();

                while (reader1.Read())
                {
                    data1.Add(new string[4]);

                    data1[data1.Count - 1][0] = reader1[0].ToString();
                    data1[data1.Count - 1][1] = reader1[1].ToString();
                    data1[data1.Count - 1][2] = reader1[2].ToString();
                    data1[data1.Count - 1][3] = reader1[3].ToString();
                }
                reader1.Close();

                db.CloseConnection();

                foreach (string[] s in data1)
                    dataGridView5.Rows.Add(s);
            }
            db.CloseConnection();
        }
        private void SaveToExcel(string file_name)
        {
            saveFileDialog1.InitialDirectory = "F:\\УНИВЕР\\4 КУРС\\8 СЕМЕСТР\\КОРПОРАТИВНЫЕ ИС";
            saveFileDialog1.Title = "SAVE AS EXCEL FILE";
            //saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx";
            saveFileDialog1.FileName = file_name;
            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                Cursor.Current = Cursors.WaitCursor;
                Excel.Application exApp = new Excel.Application();
                exApp.Application.Workbooks.Add(Type.Missing);
                exApp.Columns.ColumnWidth = 28;
                //exApp.Visible = true;
                for (int i = 1; i < dataGridView5.Columns.Count + 1; i++)
                {
                    exApp.Cells[1, i] = dataGridView5.Columns[i - 1].HeaderText;
                    //exApp.Cells[1, i].Style.Font.Bolt = true;
                }
                for (int i = 0; i < dataGridView5.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView5.Columns.Count; j++)
                    {
                        exApp.Cells[i + 2, j + 1] = dataGridView5.Rows[i].Cells[j].Value.ToString();
                    }

                }
                exApp.ActiveWorkbook.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                exApp.ActiveWorkbook.Saved = true;
                exApp.Quit();
            }
            Cursor.Current = Cursors.Default;
        }
        private void textBox_price_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (Char.IsDigit(number))
            {
                //e.Handled = true;
                label22.Visible = false;
            }
            else
            {
                label22.Visible = true;
            }
        }
        private void textBox_colvo_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (Char.IsDigit(number))
            {
                //e.Handled = true;
                label23.Visible = false;
            }
            else
            {
                label23.Visible = true;
            }
        }
        private void textBox_date_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (Char.IsDigit(number) || number == 45)
            {
                //e.Handled = true;
                label23.Visible = false;
            }
            else
            {
                label24.Visible = true;
            }
        }
        private void textBox_user_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (number > 'А' || number < 'я' || number == 46)
            {
                //e.Handled = true;
                label25.Visible = false;
            }
            else
            {
                label25.Visible = true;
            }
        }
        private void textBox_date2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (Char.IsDigit(number) || number == 45)
            {
                //e.Handled = true;
                label26.Visible = false;
            }
            else
            {
                label26.Visible = true;
            }
        }
        private void textBox_colvo_prihod_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (Char.IsDigit(number))
            {
                //e.Handled = true;
                label27.Visible = false;
            }
            else
            {
                label27.Visible = true;
            }
        }
        private void textBox_date_ot_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (Char.IsDigit(number) || number == 45)
            {
                //e.Handled = true;
                label28.Visible = false;
            }
            else
            {
                label28.Visible = true;
            }
        }
        private void textBox_date_do_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (Char.IsDigit(number) || number == 45)
            {
                //e.Handled = true;
                label28.Visible = false;
            }
            else
            {
                label28.Visible = true;
            }
        }
        private void textBoxFIO_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (number > 'А' || number < 'я' || number == 46)
            {
                //e.Handled = true;
                label29.Visible = false;
            }
            else
            {
                label29.Visible = true;
            }
        }
        private void textBoxLogin_KeyPress(object sender, KeyPressEventArgs e)
        {

        }
        private void textBoxPassword_KeyPress(object sender, KeyPressEventArgs e)
        {

        }
        private void buttonProvider_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            panel5.Visible = false;
            panel6.Visible = true;

            ShowProvider();
        }
        private void ShowProvider()
        {
            dataGridView6.Rows.Clear();
            MySqlDataAdapter msda = new MySqlDataAdapter();
            string sql = "SELECT * FROM provider;";

            DataTable dt = new DataTable();

            db.OpenConnection();
            MySqlCommand command = new MySqlCommand(sql, db.GetConnection());
            msda.SelectCommand = command;
            msda.Fill(dt);

            MySqlDataReader reader = command.ExecuteReader();
            List<string[]> data = new List<string[]>();

            while (reader.Read())
            {
                data.Add(new string[4]);

                data[data.Count - 1][0] = reader[0].ToString();
                data[data.Count - 1][1] = reader[1].ToString();
                data[data.Count - 1][2] = reader[2].ToString();
                data[data.Count - 1][3] = reader[3].ToString();
            }

            reader.Close();

            db.CloseConnection();

            foreach (string[] s in data)
                dataGridView6.Rows.Add(s);
        }
        private void buttonAddProvider_Click(object sender, EventArgs e)
        {
            MySqlDataAdapter msda = new MySqlDataAdapter();
            string sql = "INSERT INTO provider (provider_name, predstavitel, city) VALUES ('" + textBox_provider_name.Text + "', '"+ textBox_predstavitel.Text + "', '"+ textBox_city.Text + "');";

            DataTable dt = new DataTable();

            db.OpenConnection();
            MySqlCommand command = new MySqlCommand(sql, db.GetConnection());
            msda.SelectCommand = command;
            msda.Fill(dt);

            //MySqlDataReader reader = command.ExecuteReader();

            db.CloseConnection();

            ShowProvider();
        }
        private void buttonDeleteProvider_Click(object sender, EventArgs e)
        {
            string value = dataGridView6.CurrentRow.Cells[0].Value.ToString();

            MySqlDataAdapter msda = new MySqlDataAdapter();
            string sql = "DELETE FROM provider WHERE id_provider = " + value;

            DataTable dt = new DataTable();

            db.OpenConnection();
            MySqlCommand command = new MySqlCommand(sql, db.GetConnection());
            msda.SelectCommand = command;
            msda.Fill(dt);

            //MySqlDataReader reader = command.ExecuteReader();

            db.CloseConnection();

            ShowProvider();
        }
        private void buttonUpdateProvider_Click(object sender, EventArgs e)
        {
            string value = dataGridView6.CurrentRow.Cells[0].Value.ToString();

            MySqlDataAdapter msda = new MySqlDataAdapter();
            string sql = "UPDATE provider SET provider_name = '" + textBox_provider_name.Text + "', predstavitel = '" + textBox_predstavitel.Text + "', city = '" + textBox_city.Text + "' WHERE id_provider = " + value;

            DataTable dt = new DataTable();

            db.OpenConnection();
            MySqlCommand command = new MySqlCommand(sql, db.GetConnection());
            msda.SelectCommand = command;
            msda.Fill(dt);

            //MySqlDataReader reader = command.ExecuteReader();

            db.CloseConnection();

            ShowProvider();
        }
        private void dataGridView6_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox_provider_name.Text = dataGridView6.Rows[e.RowIndex].Cells[1].Value.ToString();
            textBox_predstavitel.Text = dataGridView6.Rows[e.RowIndex].Cells[2].Value.ToString();
            textBox_city.Text = dataGridView6.Rows[e.RowIndex].Cells[3].Value.ToString();
        }

        private void textBox_gruppa_tovar_id_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number))
            {
                //e.Handled = true;
                label33.Visible = false;
            }
            else
            {
                label33.Visible = true;
            }
        }

        private void textBox_provider_id_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number))
            {
                //e.Handled = true;
                label34.Visible = false;
            }
            else
            {
                label34.Visible = true;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace ComputerTechnique
{
    public partial class Employee : Form
    {
        int id = 0;
        String view = "";
        SqlDataAdapter sqlDataAdapter;
        public Employee()
        {
            InitializeComponent();
            comboBoxLoad();
        }
        public Employee(int id, String view)
        {
            InitializeComponent();
            this.id = id;
            this.view = view;
            comboBoxLoad();
            textBoxLoad();
            if (view == "watch")
            {
                button2.Visible = false;
                linkLabel1.Visible = false;
                linkLabel2.Visible = false;
                textBox1.ReadOnly = true;
                textBox2.ReadOnly = true;
                textBox3.ReadOnly = true;
                textBox4.ReadOnly = true;
                textBox5.ReadOnly = true;
                comboBox1.Enabled = false;
                comboBox2.Enabled = false;
            }
        }
        

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            NewType type = new NewType("Department");
            type.ShowDialog();
            comboBoxUpdate("Department");
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            NewType type = new NewType("Post");
            type.ShowDialog();
            comboBoxUpdate("Post");
        }
        private void comboBoxUpdate(String linkName)
        {
            switch (linkName)
            {
                case "Department":
                    comboBoxDepartment();
                    break;
                default:
                    comboBoxPost();
                    break;
            }
        }
        private void comboBoxLoad()
        {
            comboBoxDepartment();
            comboBoxPost();
        }
        private void comboBoxDepartment()
        {
            sqlDataAdapter = new SqlDataAdapter("Select ID_Department, Name from Department", Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            comboBox1.DataSource = dataTable;
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "ID_Department";
        }
        private void comboBoxPost()
        {
            sqlDataAdapter = new SqlDataAdapter("Select ID_Post, Name from Post", Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            comboBox2.DataSource = dataTable;
            comboBox2.DisplayMember = "Name";
            comboBox2.ValueMember = "ID_Post";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "" || textBox5.Text == "")
            {
                MessageBox.Show("Все поля должны быть заполнены!", "Сообщение");
            }
            else
            {
                DialogResult result = MessageBox.Show("Сохранить изменения?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    if (id == 0)
                    {
                        addEmployee();
                    }
                    else
                    {
                        changeEmployee();
                    }
                }
            }   
        }

        private void addEmployee()
        {
                    try
                    {
                        Connection.connectOpen();
                        sqlDataAdapter = new SqlDataAdapter("select * from Employee", Connection.sqlConnection);
                        sqlDataAdapter.InsertCommand = new SqlCommand("Insert into Employee (Surname,Name,Patronymic,PhoneNumber,Email,Department,Post) values(@surname,@name,@patronymic,@phonenumber,@email,@department,@post)", Connection.sqlConnection);
                        sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@surname", SqlDbType.VarChar));
                        sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@name", SqlDbType.VarChar));
                        sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@patronymic", SqlDbType.VarChar));
                        sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@phonenumber", SqlDbType.VarChar));
                        sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@email", SqlDbType.VarChar));
                        sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@department", SqlDbType.Int));
                        sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@post", SqlDbType.Int));
                        sqlDataAdapter.InsertCommand.Parameters["@surname"].Value = textBox1.Text;
                        sqlDataAdapter.InsertCommand.Parameters["@name"].Value = textBox2.Text;
                        sqlDataAdapter.InsertCommand.Parameters["@patronymic"].Value = textBox3.Text;
                        sqlDataAdapter.InsertCommand.Parameters["@phonenumber"].Value = textBox4.Text;
                        sqlDataAdapter.InsertCommand.Parameters["@email"].Value = textBox5.Text;
                        sqlDataAdapter.InsertCommand.Parameters["@department"].Value = (int)comboBox1.SelectedValue;
                        sqlDataAdapter.InsertCommand.Parameters["@post"].Value = (int)comboBox2.SelectedValue;
                        sqlDataAdapter.InsertCommand.ExecuteNonQuery();
                        MessageBox.Show("Запись была успешно добавлена!", "Сообщение");
                    }
                    catch
                    {
                        MessageBox.Show("Произошла ошибка при добавлении записи!", "Сообщение");
                    }
        }

        private void changeEmployee()
        {
            try
            {
                Connection.connectOpen();
                sqlDataAdapter = new SqlDataAdapter("select * from Employee", Connection.sqlConnection);
                sqlDataAdapter.UpdateCommand = new SqlCommand("Update Employee set Surname=@surname,Name=@name,Patronymic=@patronymic,PhoneNumber=@phonenumber,Email=@email,Department=@department,Post=@post where ID_Employee = "+id, Connection.sqlConnection);
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@surname", SqlDbType.VarChar));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@name", SqlDbType.VarChar));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@patronymic", SqlDbType.VarChar));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@phonenumber", SqlDbType.VarChar));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@email", SqlDbType.VarChar));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@department", SqlDbType.Int));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@post", SqlDbType.Int));
                sqlDataAdapter.UpdateCommand.Parameters["@surname"].Value = textBox1.Text;
                sqlDataAdapter.UpdateCommand.Parameters["@name"].Value = textBox2.Text;
                sqlDataAdapter.UpdateCommand.Parameters["@patronymic"].Value = textBox3.Text;
                sqlDataAdapter.UpdateCommand.Parameters["@phonenumber"].Value = textBox4.Text;
                sqlDataAdapter.UpdateCommand.Parameters["@email"].Value = textBox5.Text;
                sqlDataAdapter.UpdateCommand.Parameters["@department"].Value = (int)comboBox1.SelectedValue;
                sqlDataAdapter.UpdateCommand.Parameters["@post"].Value = (int)comboBox2.SelectedValue;
                sqlDataAdapter.UpdateCommand.ExecuteNonQuery();
                MessageBox.Show("Запись была успешно изменена!", "Сообщение");
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при изменении записи!", "Сообщение");
            }
        }

        private void Employee_Load(object sender, EventArgs e)
        {
            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void textBoxLoad()
        {
            sqlDataAdapter = new SqlDataAdapter("Select ID_Employee, e.Surname,e.Name,e.Patronymic,PhoneNumber,Email,d.name,p.name Должность from Employee e join Post p on p.ID_Post = e.Post join Department d on d.ID_Department = e.Department where ID_Employee = " + id, Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            textBox1.Text = Convert.ToString(dataTable.Rows[0][1]);
            textBox2.Text = Convert.ToString(dataTable.Rows[0][2]);
            textBox3.Text = Convert.ToString(dataTable.Rows[0][3]);
            textBox4.Text = Convert.ToString(dataTable.Rows[0][4]);
            textBox5.Text = Convert.ToString(dataTable.Rows[0][5]);
            comboBox1.Text = Convert.ToString(dataTable.Rows[0][6]);
            comboBox2.Text = Convert.ToString(dataTable.Rows[0][7]);
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if ((c < 'А' || c > 'я') && c != '\b' && c != '-'  && c != 32)
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if ((c < 'А' || c > 'я') && c != '\b' && c != '-' && c != 32)
            {
                e.Handled = true;
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if ((c < 'А' || c > 'я') && c != '\b' && c != '-' && c != 32)
            {
                e.Handled = true;
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if ((c < 'A' || c > 'z') && !Char.IsDigit(c) && c != '-' && c != '@'  && c != '.'  && c != 32)
            {
                e.Handled = true;
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if ( !Char.IsDigit(c) && c != '-' && c != '(' && c != ')'  && c != 32)
            {
                e.Handled = true;
            }
        }

        private void Employee_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (view != "watch")
            {
                DialogResult result = MessageBox.Show("Отменить изменения?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.No)
                {
                    if (e.CloseReason == CloseReason.UserClosing)
                        e.Cancel = true;
                }
            }
        }
    }
}

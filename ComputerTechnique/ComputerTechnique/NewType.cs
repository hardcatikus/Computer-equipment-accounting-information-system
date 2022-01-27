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
    public partial class NewType : Form
    {
        String table = "";
        public NewType(String table)
        {
            InitializeComponent();
            this.table = table;
        }
        SqlDataAdapter sqlDataAdapter;

        private void button2_Click(object sender, EventArgs e)
        {
            Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Все поля должны быть заполнены!", "Сообщение");
            }
            else
            {
                Connection.connectOpen();
                sqlDataAdapter = new SqlDataAdapter("Select * from " + table, Connection.sqlConnection);
                sqlDataAdapter.InsertCommand = new SqlCommand("Insert into " + table + "(Name) values(@name)", Connection.sqlConnection);
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@name", SqlDbType.VarChar));
                sqlDataAdapter.InsertCommand.Parameters["@name"].Value = textBox1.Text.ToString();
                sqlDataAdapter.InsertCommand.ExecuteNonQuery();
                Hide();
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if ((c < 'А' || c > 'я') && (c < 'A' || c > 'z') && c != '\b' && !Char.IsDigit(c) && c != '-' && c != '/' && c != '"' && c != '.' && c != ',' && c != '(' && c != ')' && c != ':' && c != ';' && c != '%' && c != '*' && c != 32)
            {
                e.Handled = true;
            }
        }

    }
}

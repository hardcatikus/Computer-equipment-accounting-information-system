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
    public partial class DeviceMovement : Form
    {
        int idDevice, idEmployee, idMovement, prevRoom, newRoom = 0;
        int rowOfDevice, rowOfEmployee = -1;
        String view = "";
        SqlDataAdapter sqlDataAdapter;
        
        public DeviceMovement()
        {
            InitializeComponent();
            textBox5.ReadOnly = true;
            textBox6.ReadOnly = true;
        }
        public DeviceMovement(int idMovement,String view)
        {
            InitializeComponent();
            this.idMovement = idMovement;
            this.view = view;
            if (view == "watch")
            {
                textBox1.ReadOnly = true;
                textBox2.ReadOnly = true;
                textBox3.ReadOnly = true;
                textBox4.ReadOnly = true;
                checkBox1.Enabled = false;
                dateTimePicker1.Enabled = false;
                checkBox1.Enabled = false;
                button2.Visible = false;
                button3.Enabled = false;
                button4.Enabled = false;
                dataGridView1.Enabled = false;
                dataGridView2.Enabled = false;
            }
            textBoxLoad();
            textBox5.ReadOnly=true;
            textBox6.ReadOnly = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == "" || textBox4.Text == "" || textBox5.Text == "" || textBox6.Text == "")
            {
                MessageBox.Show("Все поля должны быть заполнены!", "Сообщение");
            }
            else
            {
                DialogResult result = MessageBox.Show("Сохранить изменения?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    if (idMovement == 0)
                    {
                        addMovement();
                    }
                    else
                    {
                        changeMovement();
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void DeviceMovement_Load(object sender, EventArgs e)
        {
            loadGridDevice();
            loadEmployee();
            dataGridView1.ReadOnly = true;
            dataGridView2.ReadOnly = true;
        }

        private void loadGridDevice()
        {
            string select = "";
            select = "Select ID_Device, InventoryNumber 'Инвентарный номер', d.Name Наименование, dt.Name Тип, ds.Name Статус, r.Name 'Текущее местоположение' from Device d join DeviceType dt on dt.ID_DeviceType = d.DeviceType join DeviceStatus ds on ds.ID_DeviceStatus = d.DeviceStatus join Room r on r.ID_Room = d.Room";
            sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable;
            dataGridView1.Columns[0].Visible = false;
            rowOfDevice = -1;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            searchDevice();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            searchEmployee();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            searchDevice();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            searchEmployee();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            rowOfDevice= e.RowIndex;
            textBox5.Text = dataGridView1.Rows[rowOfDevice].Cells[1].Value.ToString()+" "+ dataGridView1.Rows[rowOfDevice].Cells[2].Value.ToString();
            idDevice = Convert.ToInt16(dataGridView1.Rows[rowOfDevice].Cells[0].Value);
            textBox3.Text = dataGridView1.Rows[rowOfDevice].Cells[5].Value.ToString();
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            rowOfEmployee = e.RowIndex;
            textBox6.Text = dataGridView2.Rows[rowOfEmployee].Cells[1].Value.ToString();
            idEmployee = Convert.ToInt16(dataGridView2.Rows[rowOfEmployee].Cells[0].Value);
        }

        private void loadEmployee()
        {
            string select = "Select ID_Employee, e.Surname+' '+e.Name+' '+e.Patronymic ФИО,p.name Должность, d.name Отдел from Employee e join Post p on p.ID_Post = e.Post join Department d on d.ID_Department=e.Department";
            sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            dataGridView2.DataSource = dataTable;
            dataGridView2.Columns[0].Visible = false;
            rowOfEmployee = -1;
        }

        private void searchDevice()
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().Contains(textBox1.Text))
                        {
                            dataGridView1.Rows[i].Selected = true;
                            break;
                        }
            }
        }
        private void searchEmployee()
        {
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                dataGridView2.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                    if (dataGridView2.Rows[i].Cells[j].Value != null)
                        if (dataGridView2.Rows[i].Cells[j].Value.ToString().Contains(textBox2.Text))
                        {
                            dataGridView2.Rows[i].Selected = true;
                            break;
                        }
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if ((c < 'А' || c > 'я') && c != '\b' && !Char.IsDigit(c) && c != '-' && c != '/' && c != '"' && c != '.' && c != ',' && c != '(' && c != ')' && c != ':' && c != ';' && c != '%' && c != '*' && c != 32)
            {
                e.Handled = true;
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if ((c < 'А' || c > 'я') && c != '\b' && !Char.IsDigit(c) && c != '-' && c != '/' && c != '"' && c != '.' && c != ',' && c != '(' && c != ')' && c != ':' && c != ';' && c != '%' && c != '*' && c != 32)
            {
                e.Handled = true;
            }
        }

        private void DeviceMovement_FormClosing(object sender, FormClosingEventArgs e)
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

        private void addMovement()
        {
            try
            {
                roomCheck();
                Connection.connectOpen();
                sqlDataAdapter = new SqlDataAdapter("select * from DeviceMovement", Connection.sqlConnection);
                sqlDataAdapter.InsertCommand = new SqlCommand("Insert into DeviceMovement (Device,Employee,DateOfMovement,PreviousRoom,NewRoom) values(@device,@employee,@dateofmovement,@previousroom,@newroom)", Connection.sqlConnection);
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@device", SqlDbType.Int));
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@employee", SqlDbType.Int));
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@dateofmovement", SqlDbType.DateTime));
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@previousroom", SqlDbType.Int));
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@newroom", SqlDbType.Int));
                sqlDataAdapter.InsertCommand.Parameters["@device"].Value = idDevice.ToString();
                sqlDataAdapter.InsertCommand.Parameters["@employee"].Value = idEmployee.ToString();
                sqlDataAdapter.InsertCommand.Parameters["@dateofmovement"].Value = dateTimePicker1.Value.ToShortDateString();
                sqlDataAdapter.InsertCommand.Parameters["@previousroom"].Value = prevRoom;
                sqlDataAdapter.InsertCommand.Parameters["@newroom"].Value = newRoom;
                sqlDataAdapter.InsertCommand.ExecuteNonQuery();
                checkBoxCheck();
                MessageBox.Show("Запись была успешно добавлена!", "Сообщение");
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при добавлении записи!", "Сообщение");
            }
        }
        private void checkBoxCheck()
        {
            if (checkBox1.Checked is true)
            {
                Connection.connectOpen();
                sqlDataAdapter = new SqlDataAdapter("select * from Device", Connection.sqlConnection);
                sqlDataAdapter.UpdateCommand = new SqlCommand("Update Device set Room=@room where ID_Device= " + idDevice, Connection.sqlConnection);
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@room", SqlDbType.Int));
                sqlDataAdapter.UpdateCommand.Parameters["@room"].Value = newRoom;
                sqlDataAdapter.UpdateCommand.ExecuteNonQuery();
            }
        }

        private void roomCheck()
        {
            try
            {
                sqlDataAdapter = new SqlDataAdapter("select Id_Room from Room where name = '" + textBox3.Text + "'", Connection.sqlConnection);
                DataTable dataTable = new DataTable();
                sqlDataAdapter.Fill(dataTable);
                newRoom = Convert.ToInt32(dataTable.Rows[0][0]);
            }
            catch
            {
                Connection.connectOpen();
                sqlDataAdapter = new SqlDataAdapter("select * from Room", Connection.sqlConnection);
                sqlDataAdapter.InsertCommand = new SqlCommand("Insert into Room (Name) values(@name)", Connection.sqlConnection);
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@name", SqlDbType.VarChar));
                sqlDataAdapter.InsertCommand.Parameters["@name"].Value = textBox3.Text;
                sqlDataAdapter.InsertCommand.ExecuteNonQuery();
                sqlDataAdapter = new SqlDataAdapter("select Id_Room from Room where name = '" + textBox3.Text + "'", Connection.sqlConnection);
                DataTable dataTable = new DataTable();
                sqlDataAdapter.Fill(dataTable);
                newRoom = Convert.ToInt32(dataTable.Rows[0][0]);
            }
            try
            {
                sqlDataAdapter = new SqlDataAdapter("select Id_Room from Room where name = '" + textBox4.Text + "'", Connection.sqlConnection);
                DataTable dataTable = new DataTable();
                sqlDataAdapter.Fill(dataTable);
                prevRoom = Convert.ToInt32(dataTable.Rows[0][0]);
            }
            catch
            {
                Connection.connectOpen();
                sqlDataAdapter = new SqlDataAdapter("select * from Room", Connection.sqlConnection);
                sqlDataAdapter.InsertCommand = new SqlCommand("Insert into Room (Name) values(@name)", Connection.sqlConnection);
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@name", SqlDbType.VarChar));
                sqlDataAdapter.InsertCommand.Parameters["@name"].Value = textBox4.Text;
                sqlDataAdapter.InsertCommand.ExecuteNonQuery();
                sqlDataAdapter = new SqlDataAdapter("select Id_Room from Room where name = '" + textBox4.Text + "'", Connection.sqlConnection);
                DataTable dataTable = new DataTable();
                sqlDataAdapter.Fill(dataTable);
                prevRoom = Convert.ToInt32(dataTable.Rows[0][0]);
            }
        }

        private void changeMovement()
        {
            try
            {
                roomCheck();
                Connection.connectOpen();
                sqlDataAdapter = new SqlDataAdapter("select * from DeviceMovement", Connection.sqlConnection);
                sqlDataAdapter.UpdateCommand = new SqlCommand("Update DeviceMovement set Device=@device,Employee=@employee,DateOfMovement=@dateofmovement,PreviousRoom=@previousroom,NewRoom=@newroom where ID_DeviceMovement = "+idMovement, Connection.sqlConnection);
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@device", SqlDbType.Int));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@employee", SqlDbType.Int));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@dateofmovement", SqlDbType.DateTime));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@previousroom", SqlDbType.Int));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@newroom", SqlDbType.Int));
                sqlDataAdapter.UpdateCommand.Parameters["@device"].Value = idDevice.ToString();
                sqlDataAdapter.UpdateCommand.Parameters["@employee"].Value = idEmployee.ToString();
                sqlDataAdapter.UpdateCommand.Parameters["@dateofmovement"].Value = dateTimePicker1.Value.ToShortDateString();
                sqlDataAdapter.UpdateCommand.Parameters["@previousroom"].Value = prevRoom;
                sqlDataAdapter.UpdateCommand.Parameters["@newroom"].Value = newRoom;
                sqlDataAdapter.UpdateCommand.ExecuteNonQuery();
                checkBoxCheck();
                MessageBox.Show("Запись была успешно обновлена!", "Сообщение");
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при обновлении записи!", "Сообщение");
            }
        }
        private void textBoxLoad()
        {
            sqlDataAdapter = new SqlDataAdapter("Select ID_DeviceMovement,ID_Device,ID_Employee,d.Name,d.InventoryNumber, DateOfMovement, r1.Name,r2.Name,e.Surname+' '+e.Name+' '+e.Patronymic from DeviceMovement dm join Device d on d.ID_Device = dm.Device join Employee E on e.ID_Employee = dm.Employee join Room r1 on r1.ID_Room = dm.PreviousRoom join Room r2 on r2.ID_Room = dm.NewRoom where dm.ID_DeviceMovement = " + idMovement, Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            textBox3.Text = Convert.ToString(dataTable.Rows[0][6]);
            textBox4.Text = Convert.ToString(dataTable.Rows[0][7]);
            textBox5.Text = Convert.ToString(dataTable.Rows[0][4])+" "+ Convert.ToString(dataTable.Rows[0][3]);
            textBox6.Text = Convert.ToString(dataTable.Rows[0][8]);
            dateTimePicker1.Text = Convert.ToString(dataTable.Rows[0][5]);
            idDevice= Convert.ToInt16(dataTable.Rows[0][1]);
            idEmployee = Convert.ToInt16(dataTable.Rows[0][2]);
        }
    }
}

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
    public partial class Repair : Form
    {
        int idDevice, idEmployee, idRepair = 0;
        int rowOfDevice, rowOfEmployee = -1;
        String view = "";
        SqlDataAdapter sqlDataAdapter;

        public Repair()
        {
            InitializeComponent();
            dateTimePicker2.Enabled = false;
            textBox5.ReadOnly = true;
            textBox6.ReadOnly = true;
            loadStatusDevice();
        }

        public Repair(int idRepair, String view)
        {
            InitializeComponent();
            this.idRepair = idRepair;
            this.view = view;
            textBox5.ReadOnly = true;
            textBox6.ReadOnly = true;
            if (view == "watch")
            {
                textBox1.ReadOnly = true;
                textBox2.ReadOnly = true;
                textBox3.ReadOnly = true;
                comboBox1.Enabled = false;
                checkBox1.Enabled = false;
                checkBox2.Enabled = false;
                dateTimePicker1.Enabled = false;
                dateTimePicker2.Enabled = false;                
                button2.Visible = false;
                button3.Enabled = false;
                button4.Enabled = false;
                dataGridView1.Enabled = false;
                dataGridView2.Enabled = false;
                linkLabel2.Visible = false;
            }
            loadStatusDevice();
            textBoxLoad();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            searchDevice();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            searchDevice();
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

        private void button4_Click(object sender, EventArgs e)
        {
            searchEmployee();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            searchEmployee();
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            rowOfDevice = e.RowIndex;
            textBox5.Text = dataGridView1.Rows[rowOfDevice].Cells[1].Value.ToString() + " " + dataGridView1.Rows[rowOfDevice].Cells[2].Value.ToString();
            idDevice = Convert.ToInt16(dataGridView1.Rows[rowOfDevice].Cells[0].Value);
            comboBox1.Text = dataGridView1.Rows[rowOfDevice].Cells[4].Value.ToString();
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            rowOfEmployee = e.RowIndex;
            textBox6.Text = dataGridView2.Rows[rowOfEmployee].Cells[1].Value.ToString();
            idEmployee = Convert.ToInt16(dataGridView2.Rows[rowOfEmployee].Cells[0].Value);
        }

        private void Repair_Load(object sender, EventArgs e)
        {
            loadGridDevice();
            loadEmployee();            
            checkBox2Check();
            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            dataGridView1.ReadOnly = true;
            dataGridView2.ReadOnly = true;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked is true)
            {
                comboBox1.Enabled = true;
            }
            else
            {
                comboBox1.Enabled = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            checkBox2Check();
        }
        private void checkBox2Check()
        {
            if (checkBox2.Checked is true)
            {
                dateTimePicker2.Visible = false;
                label8.Visible = false;
            }
            else
            {
                dateTimePicker2.Visible = true;
                label8.Visible = true;
                if (view != "watch")
                {
                    dateTimePicker2.Enabled = true;
                }
                else
                {
                    dateTimePicker2.Enabled = false;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox6.Text == "" || textBox3.Text == "" || textBox5.Text == "")
            {
                MessageBox.Show("Все поля должны быть заполнены!", "Сообщение");
            }
            else
            {
                DialogResult result = MessageBox.Show("Сохранить изменения?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    if (idRepair == 0)
                    {
                        addRepair();
                    }
                    else
                    {
                        changeRepair();
                    }
                }
            }
        }

        private void loadStatusDevice()
        {
            sqlDataAdapter = new SqlDataAdapter("Select ID_DeviceStatus, Name from DeviceStatus", Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            comboBox1.DataSource = dataTable;
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "ID_DeviceStatus";
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

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            NewType type = new NewType("DeviceStatus");
            type.ShowDialog();
            comboBoxStatusLoad();
        }

        private void comboBoxStatusLoad()
        {
            sqlDataAdapter = new SqlDataAdapter("Select ID_DeviceStatus, Name from DeviceStatus", Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            comboBox1.DataSource = dataTable;
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "ID_DeviceStatus";
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if ((c < 'А' || c > 'я') && (c < 'A' || c > 'z') && c != '\b' && !Char.IsDigit(c) && c != '-' && c != '/' && c != '"' && c != '.' && c != ',' && c != '(' && c != ')' && c != ':' && c != ';' && c != '%' && c != '*' && c != 32)
            {
                e.Handled = true;
            }
        }

        private void Repair_FormClosing(object sender, FormClosingEventArgs e)
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

        private void addRepair()
        {
            try
            {
                String query = "";
                if (checkBox2.Checked is true)
                {
                    query = "Insert into RepairWork (Device,Master,StartOfWork,Description) values(@device,@master,@startofwork,@description)";
                }
                else
                {
                    query = "Insert into RepairWork (Device,Master,StartOfWork,EndofWork,Description) values(@device,@master,@startofwork,@endofwork,@description)";
                }
                Connection.connectOpen();
                sqlDataAdapter = new SqlDataAdapter("select * from RepairWork", Connection.sqlConnection);
                sqlDataAdapter.InsertCommand = new SqlCommand(query, Connection.sqlConnection);
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@device", SqlDbType.Int));
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@master", SqlDbType.Int));
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@startofwork", SqlDbType.DateTime));
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@description", SqlDbType.VarChar));
                sqlDataAdapter.InsertCommand.Parameters["@device"].Value = idDevice.ToString();
                sqlDataAdapter.InsertCommand.Parameters["@master"].Value = idEmployee.ToString();
                sqlDataAdapter.InsertCommand.Parameters["@startofwork"].Value = dateTimePicker1.Value.ToShortDateString();
                if (checkBox2.Checked is false)
                {
                    sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@endofwork", SqlDbType.DateTime));
                    sqlDataAdapter.InsertCommand.Parameters["@endofwork"].Value = dateTimePicker2.Value.ToShortDateString();
                }
                sqlDataAdapter.InsertCommand.Parameters["@description"].Value = textBox3.Text;
                sqlDataAdapter.InsertCommand.ExecuteNonQuery();
                if (checkBox1.Checked is true)
                {
                    Connection.connectOpen();
                    sqlDataAdapter = new SqlDataAdapter("select * from Device", Connection.sqlConnection);
                    sqlDataAdapter.UpdateCommand = new SqlCommand("Update Device set DeviceStatus=@status where ID_Device= " + idDevice, Connection.sqlConnection);
                    sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@status", SqlDbType.Int));
                    sqlDataAdapter.UpdateCommand.Parameters["@status"].Value = (int)comboBox1.SelectedValue;
                    sqlDataAdapter.UpdateCommand.ExecuteNonQuery();
                    loadGridDevice();
                }
                MessageBox.Show("Запись была успешно добавлена!", "Сообщение");
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при добавлении записи!", "Сообщение");
            }
        }
        private void textBoxLoad()
        {
            sqlDataAdapter = new SqlDataAdapter("Select ID_RepairWork,d.InventoryNumber, d.Name,StartOfWork, EndOfWork,e.Surname+' '+e.Name+' '+e.Patronymic, Description, ds.Name, ID_Device,ID_Employee from RepairWork rw join Employee e on e.ID_Employee = rw.Master join Device d on d.ID_Device = rw.Device join DeviceStatus ds on ds.ID_DeviceStatus=d.DeviceStatus where ID_RepairWork = " + idRepair, Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            textBox3.Text = Convert.ToString(dataTable.Rows[0][6]);
            textBox5.Text = Convert.ToString(dataTable.Rows[0][1]) + " " + Convert.ToString(dataTable.Rows[0][2]);
            textBox6.Text = Convert.ToString(dataTable.Rows[0][5]);            
            dateTimePicker1.Text = Convert.ToString(dataTable.Rows[0][3]);
            if (Convert.ToString(dataTable.Rows[0][4]) == "")
            {
                checkBox2.Checked = true;
                dateTimePicker2.Enabled = false;
            }
            else
            {
                dateTimePicker2.Text = Convert.ToString(dataTable.Rows[0][4]);
                checkBox2.Checked = false;                
            }
            checkBox1.Checked = false;
            comboBox1.Text = Convert.ToString(dataTable.Rows[0][7]);
            idDevice = Convert.ToInt16(dataTable.Rows[0][8]);
            idEmployee = Convert.ToInt16(dataTable.Rows[0][9]);
        }

        private void changeRepair() 
        {
            try
            {
                String query;
                if (checkBox2.Checked is true)
                {
                    query = "Update RepairWork set Device=@device,Master=@master,StartOfWork=@startofwork,EndofWork=NULL,Description=@description where ID_RepairWork = " + idRepair.ToString();
                }
                else
                {
                    query = "Update RepairWork set Device=@device,Master=@master,StartOfWork=@startofwork,EndofWork=@endofwork,Description=@description where ID_RepairWork = " + idRepair.ToString();
                }
                    Connection.connectOpen();
                sqlDataAdapter = new SqlDataAdapter("select * from RepairWork", Connection.sqlConnection);
                sqlDataAdapter.UpdateCommand = new SqlCommand(query, Connection.sqlConnection);
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@device", SqlDbType.Int));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@master", SqlDbType.Int));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@startofwork", SqlDbType.DateTime));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@description", SqlDbType.VarChar));
                sqlDataAdapter.UpdateCommand.Parameters["@device"].Value = idDevice.ToString();
                sqlDataAdapter.UpdateCommand.Parameters["@master"].Value = idEmployee.ToString();
                sqlDataAdapter.UpdateCommand.Parameters["@startofwork"].Value = dateTimePicker1.Value.ToShortDateString();
                if (checkBox2.Checked is false)
                {
                    sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@endofwork", SqlDbType.DateTime));
                    sqlDataAdapter.UpdateCommand.Parameters["@endofwork"].Value = dateTimePicker2.Value.ToShortDateString();
                }             
                sqlDataAdapter.UpdateCommand.Parameters["@description"].Value = textBox3.Text;
                sqlDataAdapter.UpdateCommand.ExecuteNonQuery();
                if (checkBox1.Checked is true)
                {
                    Connection.connectOpen();
                    sqlDataAdapter = new SqlDataAdapter("select * from Device", Connection.sqlConnection);
                    sqlDataAdapter.UpdateCommand = new SqlCommand("Update Device set DeviceStatus=@status where ID_Device= " + idDevice, Connection.sqlConnection);
                    sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@status", SqlDbType.Int));
                    sqlDataAdapter.UpdateCommand.Parameters["@status"].Value = (int)comboBox1.SelectedValue;
                    sqlDataAdapter.UpdateCommand.ExecuteNonQuery();
                    loadGridDevice();
                }
                
                MessageBox.Show("Запись была успешно обновлена!", "Сообщение");
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при обновлении записи!", "Сообщение");
            }
        }
    }
}

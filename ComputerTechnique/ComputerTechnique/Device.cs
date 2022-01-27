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
    public partial class Device : Form
    {
        int id, rowOfCharacteristic, room, idDevice = 0;
        string view = "";
        DataTable dataTableAdded = new DataTable();
        SqlDataAdapter sqlDataAdapter;

        public Device()
        {
            InitializeComponent();
            comboBoxLoad();
            dataGridLoad();
        }

        public Device(int id, string view)
        {
            InitializeComponent();
            this.id = id;
            this.view = view;
            comboBoxLoad();
            textBoxLoad();
            loadGridCharacteristics();
            if (view == "watch")
            {
                button2.Visible = false;
                linkLabel1.Visible = false;
                linkLabel2.Visible = false;
                linkLabel3.Visible = false;
                textBox1.ReadOnly = true;
                textBox2.ReadOnly = true;
                textBox3.ReadOnly = true;
                textBox4.ReadOnly = true;
                textBox5.ReadOnly = true;
                comboBox1.Enabled = false;
                comboBox2.Enabled = false;
                comboBox3.Enabled = false;
                dateTimePicker1.Enabled = false;
                button3.Enabled = false;
                dataGridView1.ReadOnly = true;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            NewType type = new NewType("DeviceType");
            type.ShowDialog();
            comboBoxUpdate("Type");
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            NewType type = new NewType("DeviceStatus");
            type.ShowDialog();
            comboBoxUpdate("Status");
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Employee type = new Employee();
            type.ShowDialog();
            comboBoxUpdate("Employee");
        }
        private void comboBoxUpdate(String linkName)
        {
            switch (linkName)
            {
                case "Type":
                    comboBoxTypeLoad();
                    break;
                case "Status":
                    comboBoxStatusLoad();
                    break;
                default:
                    comboBoxEmployeeLoad();
                    break;
            }
        }
        private void comboBoxLoad()
        {
            comboBoxTypeLoad();
            comboBoxStatusLoad();
            comboBoxEmployeeLoad();
        }
        private void comboBoxTypeLoad()
        {
            sqlDataAdapter = new SqlDataAdapter("Select ID_DeviceType, Name from DeviceType", Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            comboBox1.DataSource = dataTable;
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "ID_DeviceType";
        }
        private void comboBoxStatusLoad()
        {
            sqlDataAdapter = new SqlDataAdapter("Select ID_DeviceStatus, Name from DeviceStatus", Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            comboBox2.DataSource = dataTable;
            comboBox2.DisplayMember = "Name";
            comboBox2.ValueMember = "ID_DeviceStatus";
        }
        private void comboBoxEmployeeLoad()
        {
            sqlDataAdapter = new SqlDataAdapter("Select ID_Employee, Surname+' '+Name+' '+Patronymic Сотрудник from Employee", Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            comboBox3.DataSource = dataTable;
            comboBox3.DisplayMember = "Сотрудник";
            comboBox3.ValueMember = "ID_Employee";
        }

        private void Device_Load(object sender, EventArgs e)
        {
            label8.Visible = false;
            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "" || textBox5.Text == "")
            {
                MessageBox.Show("Все поля должны быть заполнены!", "Сообщение");
            }
            else
            {
                foreach (DataGridViewRow rw in this.dataGridView1.Rows)
                {
                    if (rw.Cells[1].Value == null || rw.Cells[1].Value == DBNull.Value || String.IsNullOrWhiteSpace(rw.Cells[1].Value.ToString()))
                    {
                        if (rw.Cells[2].Value == null || rw.Cells[2].Value == DBNull.Value || String.IsNullOrWhiteSpace(rw.Cells[2].Value.ToString()))
                        {
                            continue;
                        }
                        else
                        {
                            MessageBox.Show("Заполните оба столбца характеристик!", "Сообщение");
                            return;
                        }
                    }
                    else
                    {
                        if (rw.Cells[2].Value == null || rw.Cells[2].Value == DBNull.Value || String.IsNullOrWhiteSpace(rw.Cells[2].Value.ToString()))
                        {
                            MessageBox.Show("Заполните оба столбца характеристик!", "Сообщение");
                            return;
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
                DialogResult result = MessageBox.Show("Сохранить изменения?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        if (view != "change")
                        {
                            sqlDataAdapter = new SqlDataAdapter("select Id_Device from Device where InventoryNumber = '" + textBox2.Text + "'", Connection.sqlConnection);
                            DataTable dataTable = new DataTable();
                            sqlDataAdapter.Fill(dataTable);
                            int device = Convert.ToInt32(dataTable.Rows[0][0]);
                            MessageBox.Show("Устройство с указанным инвентарным номером уже существует!", "Сообщение");
                        }
                        else
                        {
                            sqlDataAdapter = new SqlDataAdapter("select Id_Device from Device where InventoryNumber = '" + textBox2.Text + "'", Connection.sqlConnection);
                            DataTable dataTable = new DataTable();
                            sqlDataAdapter.Fill(dataTable);
                            int device = Convert.ToInt32(dataTable.Rows[0][0]);
                            if (device != id)
                            {
                                MessageBox.Show("Устройство с указанным инвентарным номером уже существует!", "Сообщение");
                            }
                            else
                            {
                                changeDevice();
                                changeCharacteristics();
                            }
                        }
                    }
                    catch
                    {
                        if (id == 0)
                        {
                            addDevice();
                            addCharacteristics();
                        }
                        else
                        {
                            changeDevice();
                            changeCharacteristics();
                        }
                    }
                }
            }
        }

        private void changeCharacteristics()
        {
            foreach (DataGridViewRow rw in dataGridView1.Rows)
            {
                if (rw.Cells[0].Value == null || rw.Cells[0].Value == DBNull.Value || String.IsNullOrWhiteSpace(rw.Cells[0].Value.ToString()))
                {
                    if (rw.Cells[2].Value == null || rw.Cells[2].Value == DBNull.Value || String.IsNullOrWhiteSpace(rw.Cells[2].Value.ToString()))
                    {
                        continue;
                    }
                    else
                    {
                        Connection.connectOpen();
                        sqlDataAdapter = new SqlDataAdapter("select * from Characteristic", Connection.sqlConnection);
                        sqlDataAdapter.InsertCommand = new SqlCommand("Insert into Characteristic (Name,Value,Device) values(@name,@value,@device)", Connection.sqlConnection);
                        sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@name", SqlDbType.VarChar));
                        sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@value", SqlDbType.VarChar));
                        sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@device", SqlDbType.Int));
                        sqlDataAdapter.InsertCommand.Parameters["@name"].Value = rw.Cells[1].Value.ToString();
                        sqlDataAdapter.InsertCommand.Parameters["@value"].Value = rw.Cells[2].Value.ToString();
                        sqlDataAdapter.InsertCommand.Parameters["@device"].Value = id;
                        sqlDataAdapter.InsertCommand.ExecuteNonQuery();
                    }
                }
                else
                {
                    if (rw.Cells[2].Value == null || rw.Cells[2].Value == DBNull.Value || String.IsNullOrWhiteSpace(rw.Cells[2].Value.ToString()))
                    {
                        Connection.connectOpen();
                        sqlDataAdapter = new SqlDataAdapter("Select * from Characteristic", Connection.sqlConnection);
                        sqlDataAdapter.DeleteCommand = new SqlCommand("delete from Characteristic where ID_Characteristic=@characteristic", Connection.sqlConnection);
                        sqlDataAdapter.DeleteCommand.Parameters.Add(new SqlParameter("@characteristic", SqlDbType.Int));
                        sqlDataAdapter.DeleteCommand.Parameters["@characteristic"].Value = Convert.ToInt32(rw.Cells[0].Value);
                        sqlDataAdapter.DeleteCommand.ExecuteNonQuery();
                    }
                    else
                    {
                        for (int j = 0; j < dataTableAdded.Rows.Count; j++)
                        {
                            if (dataTableAdded.Rows[j][0].ToString() == rw.Cells[0].Value.ToString())
                            {
                                Connection.connectOpen();
                                sqlDataAdapter = new SqlDataAdapter("select * from Characteristic", Connection.sqlConnection);
                                sqlDataAdapter.UpdateCommand = new SqlCommand("Update Characteristic set Name=@name,Value=@value where ID_Characteristic = " + Convert.ToInt32(rw.Cells[0].Value), Connection.sqlConnection);
                                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@name", SqlDbType.VarChar));
                                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@value", SqlDbType.VarChar));
                                sqlDataAdapter.UpdateCommand.Parameters["@name"].Value = rw.Cells[1].Value.ToString();
                                sqlDataAdapter.UpdateCommand.Parameters["@value"].Value = rw.Cells[2].Value.ToString();
                                sqlDataAdapter.UpdateCommand.ExecuteNonQuery();
                                break;
                            }
                        }
                    }
                }
            }                        
        }

        private void loadGridCharacteristics()
        {
            string select = "";
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView1.Columns.Add("ID_Characteristic", "ID_Characteristic");
            dataGridView1.Columns.Add("Характеристика", "Характеристика");
            dataGridView1.Columns.Add("Значение", "Значение");
            dataGridView1.Columns[0].Visible = false;
            ((DataGridViewTextBoxColumn)dataGridView1.Columns[1]).MaxInputLength = 50;
            ((DataGridViewTextBoxColumn)dataGridView1.Columns[2]).MaxInputLength = 50;
            select = "SELECT ID_Characteristic, c.Name, Value, Device FROM Characteristic c join Device d on d.ID_Device = c.Device  where ID_Device = " + id;
            sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
            sqlDataAdapter.Fill(dataTableAdded);
            for (int g = 0; g < dataTableAdded.Rows.Count; g++)
            {
                dataGridView1.Rows.Add();
                rowOfCharacteristic = dataGridView1.Rows.Count - 1;
                dataGridView1.Rows[rowOfCharacteristic].Cells[0].Value = dataTableAdded.Rows[g][0].ToString();
                dataGridView1.Rows[rowOfCharacteristic].Cells[1].Value = dataTableAdded.Rows[g][1].ToString();
                dataGridView1.Rows[rowOfCharacteristic].Cells[2].Value = dataTableAdded.Rows[g][2].ToString();
            }
        }

        private void addCharacteristics()
        {
            if (idDevice == 0)
            {
                sqlDataAdapter = new SqlDataAdapter("SELECT TOP 1 ID_Device FROM Device ORDER BY ID_Device DESC", Connection.sqlConnection);
                DataTable dataTable = new DataTable();
                sqlDataAdapter.Fill(dataTable);
                idDevice = Convert.ToInt16(dataTable.Rows[0][0]);
            }
            foreach (DataGridViewRow rw in this.dataGridView1.Rows)
            {
                if (rw.Cells[1].Value == null || rw.Cells[1].Value == DBNull.Value || String.IsNullOrWhiteSpace(rw.Cells[1].Value.ToString()))
                {
                    if (rw.Cells[2].Value == null || rw.Cells[2].Value == DBNull.Value || String.IsNullOrWhiteSpace(rw.Cells[2].Value.ToString()))
                    {
                        continue;
                    }
                }
                else
                {
                    Connection.connectOpen();
                    sqlDataAdapter = new SqlDataAdapter("select * from Characteristic", Connection.sqlConnection);
                    sqlDataAdapter.InsertCommand = new SqlCommand("Insert into Characteristic (Name,Value,Device) values(@name,@value,@device)", Connection.sqlConnection);
                    sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@name", SqlDbType.VarChar));
                    sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@value", SqlDbType.VarChar));
                    sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@device", SqlDbType.Int));
                    sqlDataAdapter.InsertCommand.Parameters["@name"].Value = rw.Cells[1].Value.ToString();
                    sqlDataAdapter.InsertCommand.Parameters["@value"].Value = rw.Cells[2].Value.ToString();
                    sqlDataAdapter.InsertCommand.Parameters["@device"].Value = idDevice;
                    sqlDataAdapter.InsertCommand.ExecuteNonQuery();
                }
            }
        }

        private void changeDevice()
        {
            try
            {
                try
                {
                    sqlDataAdapter = new SqlDataAdapter("select Id_Room from Room where name = '" + textBox3.Text + "'", Connection.sqlConnection);
                    DataTable dataTable = new DataTable();
                    sqlDataAdapter.Fill(dataTable);
                    room = Convert.ToInt32(dataTable.Rows[0][0]);
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
                    room = Convert.ToInt32(dataTable.Rows[0][0]);
                }
                Connection.connectOpen();
                sqlDataAdapter = new SqlDataAdapter("select * from Device", Connection.sqlConnection);
                sqlDataAdapter.UpdateCommand = new SqlCommand("Update Device set Name=@name,InventoryNumber=@inventorynumber,DeviceType=@devicetype,DeviceStatus=@devicestatus,Room=@room,DateOfPurchase=@dateofpurchase,GuarantyPeriod=@guarantyperiod,Price=@price,ResponsiblEemployee=@responsibleemployee where ID_Device= " + id, Connection.sqlConnection);
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@name", SqlDbType.VarChar));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@inventorynumber", SqlDbType.VarChar));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@devicetype", SqlDbType.Int));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@devicestatus", SqlDbType.Int));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@room", SqlDbType.Int));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@dateofpurchase", SqlDbType.DateTime));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@guarantyperiod", SqlDbType.VarChar));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@price", SqlDbType.Float));
                sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@responsibleemployee", SqlDbType.Int));
                sqlDataAdapter.UpdateCommand.Parameters["@name"].Value = textBox1.Text;
                sqlDataAdapter.UpdateCommand.Parameters["@inventorynumber"].Value = textBox2.Text;
                sqlDataAdapter.UpdateCommand.Parameters["@devicetype"].Value = (int)comboBox1.SelectedValue;
                sqlDataAdapter.UpdateCommand.Parameters["@devicestatus"].Value = (int)comboBox2.SelectedValue;
                sqlDataAdapter.UpdateCommand.Parameters["@room"].Value = room;
                sqlDataAdapter.UpdateCommand.Parameters["@dateofpurchase"].Value = dateTimePicker1.Value.ToShortDateString();
                sqlDataAdapter.UpdateCommand.Parameters["@guarantyperiod"].Value = textBox4.Text;
                sqlDataAdapter.UpdateCommand.Parameters["@price"].Value = Convert.ToSingle(textBox5.Text);
                sqlDataAdapter.UpdateCommand.Parameters["@responsibleemployee"].Value = (int)comboBox3.SelectedValue;
                sqlDataAdapter.UpdateCommand.ExecuteNonQuery();
                MessageBox.Show("Запись была успешно изменена!", "Сообщение");
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при изменении записи!", "Сообщение");
            }
        }

        private void addDevice()
        {
            try
            {
                try
                {
                    sqlDataAdapter = new SqlDataAdapter("select Id_Room from Room where name = '" + textBox3.Text + "'", Connection.sqlConnection);
                    DataTable dataTable = new DataTable();
                    sqlDataAdapter.Fill(dataTable);
                    room = Convert.ToInt32(dataTable.Rows[0][0]);
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
                    room = Convert.ToInt32(dataTable.Rows[0][0]);
                }
                Connection.connectOpen();
                sqlDataAdapter = new SqlDataAdapter("select * from Device", Connection.sqlConnection);
                sqlDataAdapter.InsertCommand = new SqlCommand("Insert into Device (Name,InventoryNumber,DeviceType,DeviceStatus,Room,DateOfPurchase,GuarantyPeriod,Price,ResponsibleEmployee) values(@name,@inventorynumber,@devicetype,@devicestatus,@room,@dateofpurchase,@guarantyperiod,@price,@responsibleemployee)", Connection.sqlConnection);
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@name", SqlDbType.VarChar));
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@inventorynumber", SqlDbType.VarChar));
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@devicetype", SqlDbType.Int));
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@devicestatus", SqlDbType.Int));
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@room", SqlDbType.Int));
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@dateofpurchase", SqlDbType.DateTime));
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@guarantyperiod", SqlDbType.VarChar));
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@price", SqlDbType.Float));
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@responsibleemployee", SqlDbType.Int));
                sqlDataAdapter.InsertCommand.Parameters["@name"].Value = textBox1.Text;
                sqlDataAdapter.InsertCommand.Parameters["@inventorynumber"].Value = textBox2.Text;
                sqlDataAdapter.InsertCommand.Parameters["@devicetype"].Value = (int)comboBox1.SelectedValue;
                sqlDataAdapter.InsertCommand.Parameters["@devicestatus"].Value = (int)comboBox2.SelectedValue;
                sqlDataAdapter.InsertCommand.Parameters["@room"].Value = room;
                sqlDataAdapter.InsertCommand.Parameters["@dateofpurchase"].Value = dateTimePicker1.Value.ToShortDateString();
                sqlDataAdapter.InsertCommand.Parameters["@guarantyperiod"].Value = textBox4.Text;
                sqlDataAdapter.InsertCommand.Parameters["@price"].Value = Convert.ToSingle(textBox5.Text);
                sqlDataAdapter.InsertCommand.Parameters["@responsibleemployee"].Value = (int)comboBox3.SelectedValue;
                sqlDataAdapter.InsertCommand.ExecuteNonQuery();
                MessageBox.Show("Запись была успешно добавлена!", "Сообщение");
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при добавлении записи!", "Сообщение");
            }
        }

        private void textBoxLoad()
        {
            sqlDataAdapter = new SqlDataAdapter("Select d.ID_Device,d.Name,d.InventoryNumber,d.DateOfPurchase,d.GuarantyPeriod,d.Price,r.Name,ds.Name,dt.Name,e.Surname+' '+e.Name+' '+e.Patronymic from Device d join DeviceStatus ds on ds.ID_DeviceStatus = d.DeviceStatus join DeviceType dt on dt.ID_DeviceType = d.DeviceType join Employee e on e.ID_Employee = d.ResponsibleEmployee join Room r on r.ID_Room = d.Room where d.ID_Device = " + id, Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            textBox1.Text = Convert.ToString(dataTable.Rows[0][1]);
            textBox2.Text = Convert.ToString(dataTable.Rows[0][2]);
            textBox3.Text = Convert.ToString(dataTable.Rows[0][6]);
            textBox4.Text = Convert.ToString(dataTable.Rows[0][4]);
            textBox5.Text = Convert.ToString(dataTable.Rows[0][5]);
            comboBox1.Text = Convert.ToString(dataTable.Rows[0][8]);
            comboBox2.Text = Convert.ToString(dataTable.Rows[0][7]);
            comboBox3.Text = Convert.ToString(dataTable.Rows[0][9]);
            dateTimePicker1.Text = Convert.ToString(dataTable.Rows[0][3]);
            guaranteeCheck();
        }
        private void guaranteeCheck()
        {
            DateTime date = dateTimePicker1.Value;
            if ((textBox4.Text != "") && (date.AddMonths(Convert.ToInt32(textBox4.Text)) < DateTime.Now.Date))
            {
                label8.Visible = true;
            }
            else
            {
                label8.Visible = false;
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            guaranteeCheck();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            guaranteeCheck();
        }

        private void textBox4_Leave(object sender, EventArgs e)
        {
            guaranteeCheck();
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if ((c < 'А' || c > 'я') && (c < 'A' || c > 'z') && c != '\b' && !Char.IsDigit(c) && c != '-' && c != '/' && c != '"' && c != '.' && c != ',' && c != '(' && c != ')' && c != ':' && c != ';' && c != '%' && c != '*' && c != 32)
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if (!Char.IsDigit(c) && c != '\b')
            {
                e.Handled = true;
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
            if (!Char.IsDigit(c) && c != '\b')
            {
                e.Handled = true;
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if (!Char.IsDigit(c) && c != '\b' && c != ',' && c != '.')
            {
                e.Handled = true;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (view != "change")
            {
                dataGridLoad();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add();
        }

        private void Device_FormClosing(object sender, FormClosingEventArgs e)
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
        private void dataGridLoad()
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView1.Columns.Add("ID_Characteristic", "ID_Characteristic");
            dataGridView1.Columns.Add("Характеристика", "Характеристика");
            dataGridView1.Columns.Add("Значение", "Значение");
            dataGridView1.Columns[0].Visible = false;
            ((DataGridViewTextBoxColumn)dataGridView1.Columns[1]).MaxInputLength = 50;
            ((DataGridViewTextBoxColumn)dataGridView1.Columns[2]).MaxInputLength = 50;
            String deviceType = comboBox1.Text;
            switch (deviceType)
            {
                case "Компьютер":
                    String[] characteristisPC = { "Модель процессора", "Скорость процессора", "Количество ядер", "Объем памяти", "Объем оперативной памяти", "Неисправности" };
                    foreach (String i in characteristisPC)
                    {
                        dataGridView1.Rows.Add();
                        rowOfCharacteristic = dataGridView1.Rows.Count - 1;
                        dataGridView1.Rows[rowOfCharacteristic].Cells[1].Value = i;
                    }
                    break;
                case "Монитор":
                    String[] characteristisMonitor = { "Диагональ", "Разрешение", "Неисправности" };
                    foreach (String i in characteristisMonitor)
                    {
                        dataGridView1.Rows.Add();
                        rowOfCharacteristic = dataGridView1.Rows.Count - 1;
                        dataGridView1.Rows[rowOfCharacteristic].Cells[1].Value = i;
                    }
                    break;
                case "Принтер":
                    String[] characteristisPrinter = { "Цветность", "Скорость печати", "Формат бумаги", "Неисправности" };
                    foreach (String i in characteristisPrinter)
                    {
                        dataGridView1.Rows.Add();
                        rowOfCharacteristic = dataGridView1.Rows.Count - 1;
                        dataGridView1.Rows[rowOfCharacteristic].Cells[1].Value = i;
                    }
                    break;
                case "Моноблок":
                    String[] characteristisMonobloc = { "Модель процессора", "Скорость процессора", "Количество ядер", "Объем памяти", "Объем оперативной памяти", "Диагональ", "Разрешение", "Неисправности" };
                    foreach (String i in characteristisMonobloc)
                    {
                        dataGridView1.Rows.Add();
                        rowOfCharacteristic = dataGridView1.Rows.Count - 1;
                        dataGridView1.Rows[rowOfCharacteristic].Cells[1].Value = i;
                    }
                    break;
                default:
                    String[] defect = { "Неисправности" };
                    foreach (String i in defect)
                    {
                        dataGridView1.Rows.Add();
                        rowOfCharacteristic = dataGridView1.Rows.Count - 1;
                        dataGridView1.Rows[rowOfCharacteristic].Cells[1].Value = i;
                    }
                    break;
            }

        }
    }
}

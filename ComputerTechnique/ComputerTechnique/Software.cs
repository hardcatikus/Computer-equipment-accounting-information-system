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
    public partial class Software : Form
    {
        int idSoftware, rowOfDevice, rowOfDeviceAdd = 0;
        bool deviceAdded = false;
        string view = "";
        DataTable dataTableAdded = new DataTable();
        SqlDataAdapter sqlDataAdapter;
        public Software()
        {
            InitializeComponent();
            comboBoxLoad();
            checkBox1.Checked = true;
        }

        public Software(int idSoftware, String view)
        {
            InitializeComponent();
            this.idSoftware = idSoftware;
            this.view = view;
            comboBoxLoad();
            textBoxLoad();
            checkBoxCheck();
            if (view == "watch")
            {
                textBox1.ReadOnly = true;
                textBox2.ReadOnly = true;
                textBox5.ReadOnly = true;
                textBox4.ReadOnly = true;
                dateTimePicker1.Enabled = false;
                checkBox1.Enabled = false;
                comboBox3.Enabled = false;
                linkLabel3.Visible = false;
                button2.Visible = false;
                dataGridView1.ReadOnly = true;
                dataGridView2.ReadOnly = true;
                textBox3.Enabled = false;
                button4.Enabled = false;
            }
            else
            {
                if (checkBox1.Checked is true)
                {
                    checkBox1.Enabled = false;
                }
            }
        }

        private void Software_Load(object sender, EventArgs e)
        {
            comboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            dataGridView1.ReadOnly = true;
            dataGridView2.ReadOnly = true;
            loadGridDevice();
        }
        private void loadGridDevice()
        {
            string select = "";
            select = "Select ID_Device, InventoryNumber 'Инвентарный номер', d.Name Наименование, dt.Name Тип, ds.Name Статус, r.Name 'Текущее местоположение' from Device d join DeviceType dt on dt.ID_DeviceType = d.DeviceType join DeviceStatus ds on ds.ID_DeviceStatus = d.DeviceStatus join Room r on r.ID_Room = d.Room";
            sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            dataGridView2.DataSource = dataTable;
            dataGridView2.Columns[0].Visible = false;
            rowOfDevice = -1;
            dataGridView1.Columns.Add("ID_Device", "ID_Device");
            dataGridView1.Columns.Add("Инвентарный номер", "Инвентарный номер");
            dataGridView1.Columns.Add("Наименование", "Наименование");
            dataGridView1.Columns.Add("Тип", "Тип");
            dataGridView1.Columns.Add("Статус", "Статус");
            dataGridView1.Columns.Add("Текущее местоположение", "Текущее местоположение");
            dataGridView1.Columns[0].Visible = false;
            select = "SELECT ID_Device  FROM SoftwareInstallation si join Device d on d.ID_Device = si.Device  where Software = " + idSoftware;
            sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
            sqlDataAdapter.Fill(dataTableAdded);
            for (int g = 0; g < dataTableAdded.Rows.Count; g++)
            {
                dataGridView1.Rows.Add();
                rowOfDeviceAdd = dataGridView1.Rows.Count - 1;
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    if (dataTableAdded.Rows[g][0].ToString() == dataGridView2.Rows[i].Cells[0].Value.ToString())
                    {
                        for (int j = 0; j < dataGridView2.Columns.Count; j++)
                        {
                            dataGridView1.Rows[rowOfDeviceAdd].Cells[j].Value = dataGridView2.Rows[i].Cells[j].Value;
                        }
                        break;
                    }
                }
            }
        }
        private void comboBoxLoad()
        {
            sqlDataAdapter = new SqlDataAdapter("Select ID_SoftwareType, Name from SoftwareType", Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            comboBox3.DataSource = dataTable;
            comboBox3.DisplayMember = "Name";
            comboBox3.ValueMember = "ID_SoftwareType";
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            NewType type = new NewType("SoftwareType");
            type.ShowDialog();
            comboBoxLoad();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            checkBoxCheck();
        }

        private void checkBoxCheck()
        {
            if (checkBox1.Checked is true)
            {
                textBox4.Enabled = true;
                textBox5.Enabled = true;
                textBox2.Enabled = true;
                dateTimePicker1.Enabled = true;
            }
            else
            {
                textBox4.Enabled = false;
                textBox5.Enabled = false;
                textBox2.Enabled = false;
                dateTimePicker1.Enabled = false;
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            licenseCheck();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            licenseCheck();
        }

        private void licenseCheck()
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

        private void textBox4_Leave(object sender, EventArgs e)
        {
            licenseCheck();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Все поля должны быть заполнены!", "Сообщение");
            }
            else 
            {
                if ((textBox2.Text == "" || textBox5.Text == "" || textBox4.Text == "") && (checkBox1.Checked is true))
                {
                    MessageBox.Show("Все поля должны быть заполнены!", "Сообщение");
                }
                else
                {
                    DialogResult result = MessageBox.Show("Сохранить изменения?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        if (idSoftware == 0)
                        {
                            addSoftware();
                            addInstallation();
                        }
                        else
                        {
                            changeSoftware();
                            changeInstallation();
                        }
                    }
                }
            }
        }

        private void addInstallation()
        {
            if (idSoftware == 0)
            {
                sqlDataAdapter = new SqlDataAdapter("SELECT TOP 1 ID_Software FROM Software ORDER BY ID_Software DESC", Connection.sqlConnection);
                DataTable dataTable = new DataTable();
                sqlDataAdapter.Fill(dataTable);
                idSoftware = Convert.ToInt16(dataTable.Rows[0][0]);
            }
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                Connection.connectOpen();
                sqlDataAdapter = new SqlDataAdapter("select * from SoftwareInstallation", Connection.sqlConnection);
                sqlDataAdapter.InsertCommand = new SqlCommand("Insert into SoftwareInstallation (Device,Software) values(@device,@software)", Connection.sqlConnection);
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@device", SqlDbType.Int));
                sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@software", SqlDbType.Int));
                sqlDataAdapter.InsertCommand.Parameters["@device"].Value = Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value);
                sqlDataAdapter.InsertCommand.Parameters["@software"].Value = idSoftware;
                sqlDataAdapter.InsertCommand.ExecuteNonQuery();
            }
        }

        private void changeInstallation()
        {
            //new installations
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                bool added = false;
                for (int j = 0; j < dataTableAdded.Rows.Count; j++)
                {
                    if (dataTableAdded.Rows[j][0].ToString() != dataGridView1.Rows[i].Cells[0].Value.ToString())
                    {
                        added = false;
                    }
                    else
                    {
                        added = true;
                        break;
                    }                    
                }
                if (added is false)
                {
                    Connection.connectOpen();
                    sqlDataAdapter = new SqlDataAdapter("select * from SoftwareInstallation", Connection.sqlConnection);
                    sqlDataAdapter.InsertCommand = new SqlCommand("Insert into SoftwareInstallation (Device,Software) values(@device,@software)", Connection.sqlConnection);
                    sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@device", SqlDbType.Int));
                    sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@software", SqlDbType.Int));
                    sqlDataAdapter.InsertCommand.Parameters["@device"].Value = Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value);
                    sqlDataAdapter.InsertCommand.Parameters["@software"].Value = idSoftware;
                    sqlDataAdapter.InsertCommand.ExecuteNonQuery();
                }
            }
            //nonexisting installations
            for (int j = 0; j < dataTableAdded.Rows.Count; j++)
            {
                bool added = false;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataTableAdded.Rows[j][0].ToString() != dataGridView1.Rows[i].Cells[0].Value.ToString())
                    {
                        added = false;
                    }
                    else
                    {
                        added = true;
                        break;
                    }
                }
                if (added is false)
                {
                    Connection.connectOpen();
                    sqlDataAdapter = new SqlDataAdapter("Select * from SoftwareInstallation", Connection.sqlConnection);
                    sqlDataAdapter.DeleteCommand = new SqlCommand("delete from SoftwareInstallation where Software=@software and Device=@device", Connection.sqlConnection);
                    sqlDataAdapter.DeleteCommand.Parameters.Add(new SqlParameter("@software", SqlDbType.Int));
                    sqlDataAdapter.DeleteCommand.Parameters.Add(new SqlParameter("@device", SqlDbType.Int));
                    sqlDataAdapter.DeleteCommand.Parameters["@software"].Value = idSoftware;
                    sqlDataAdapter.DeleteCommand.Parameters["@device"].Value = Convert.ToInt32(dataTableAdded.Rows[j][0]);
                    sqlDataAdapter.DeleteCommand.ExecuteNonQuery();
                }
            }            
        }
        private void addSoftware()
        {
            try
            {
                if (checkBox1.Checked is true)
                {
                    Connection.connectOpen();
                    sqlDataAdapter = new SqlDataAdapter("select * from Software", Connection.sqlConnection);
                    sqlDataAdapter.InsertCommand = new SqlCommand("Insert into Software (Name,LicenseKey,LicenseKeyDuration,DateOfPurchase,Price,SoftwareType,KeyNeed) values(@name,@licensekey,@licensekeyduration,@dateofpurchase,@price,@softwaretype,@keyneed)", Connection.sqlConnection);
                    sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@name", SqlDbType.VarChar));
                    sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@licensekey", SqlDbType.VarChar));
                    sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@licensekeyduration", SqlDbType.Int));
                    sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@dateofpurchase", SqlDbType.DateTime));
                    sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@price", SqlDbType.Float));
                    sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@softwaretype", SqlDbType.Int));
                    sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@keyneed", SqlDbType.Bit));
                    sqlDataAdapter.InsertCommand.Parameters["@name"].Value = textBox1.Text;
                    sqlDataAdapter.InsertCommand.Parameters["@licensekey"].Value = textBox2.Text;
                    sqlDataAdapter.InsertCommand.Parameters["@licensekeyduration"].Value = textBox4.Text;
                    sqlDataAdapter.InsertCommand.Parameters["@dateofpurchase"].Value = dateTimePicker1.Value.ToShortDateString();
                    sqlDataAdapter.InsertCommand.Parameters["@price"].Value = textBox5.Text;
                    sqlDataAdapter.InsertCommand.Parameters["@softwaretype"].Value = (int)comboBox3.SelectedValue;
                    sqlDataAdapter.InsertCommand.Parameters["@keyneed"].Value = true;
                    sqlDataAdapter.InsertCommand.ExecuteNonQuery();
                    MessageBox.Show("Запись была успешно добавлена!", "Сообщение");
                }
                else
                {
                    Connection.connectOpen();
                    sqlDataAdapter = new SqlDataAdapter("select * from Software", Connection.sqlConnection);
                    sqlDataAdapter.InsertCommand = new SqlCommand("Insert into Software (Name,SoftwareType,KeyNeed) values(@name,@softwaretype,@keyneed)", Connection.sqlConnection);
                    sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@name", SqlDbType.VarChar));
                    sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@softwaretype", SqlDbType.Int));
                    sqlDataAdapter.InsertCommand.Parameters.Add(new SqlParameter("@keyneed", SqlDbType.Bit));
                    sqlDataAdapter.InsertCommand.Parameters["@name"].Value = textBox1.Text;
                    sqlDataAdapter.InsertCommand.Parameters["@softwaretype"].Value = (int)comboBox3.SelectedValue;
                    sqlDataAdapter.InsertCommand.Parameters["@keyneed"].Value = false;
                    sqlDataAdapter.InsertCommand.ExecuteNonQuery();
                    MessageBox.Show("Запись была успешно добавлена!", "Сообщение");
                }
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при добавлении записи!", "Сообщение");
            }
        }

        private void changeSoftware()
        {
            try
            {
                if (checkBox1.Checked is true)
                {
                    Connection.connectOpen();
                    sqlDataAdapter = new SqlDataAdapter("select * from Software", Connection.sqlConnection);
                    sqlDataAdapter.UpdateCommand = new SqlCommand("Update Software set Name=@name,LicenseKey=@licensekey,LicenseKeyDuration=@licensekeyduration,DateOfPurchase=@dateofpurchase,Price=@price,SoftwareType=@softwaretype,KeyNeed=@keyneed where ID_Software = " + idSoftware, Connection.sqlConnection);
                    sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@name", SqlDbType.VarChar));
                    sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@licensekey", SqlDbType.VarChar));
                    sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@licensekeyduration", SqlDbType.Int));
                    sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@dateofpurchase", SqlDbType.DateTime));
                    sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@price", SqlDbType.Float));
                    sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@softwaretype", SqlDbType.Int));
                    sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@keyneed", SqlDbType.Bit));
                    sqlDataAdapter.UpdateCommand.Parameters["@name"].Value = textBox1.Text;
                    sqlDataAdapter.UpdateCommand.Parameters["@licensekey"].Value = textBox2.Text;
                    sqlDataAdapter.UpdateCommand.Parameters["@licensekeyduration"].Value = textBox4.Text;
                    sqlDataAdapter.UpdateCommand.Parameters["@dateofpurchase"].Value = dateTimePicker1.Value.ToShortDateString();
                    sqlDataAdapter.UpdateCommand.Parameters["@price"].Value = textBox5.Text;
                    sqlDataAdapter.UpdateCommand.Parameters["@softwaretype"].Value = (int)comboBox3.SelectedValue;
                    sqlDataAdapter.UpdateCommand.Parameters["@keyneed"].Value = true;
                    sqlDataAdapter.UpdateCommand.ExecuteNonQuery();
                    MessageBox.Show("Запись была успешно изменена!", "Сообщение");
                }
                else
                {
                    Connection.connectOpen();
                    sqlDataAdapter = new SqlDataAdapter("select * from Software", Connection.sqlConnection);
                    sqlDataAdapter.UpdateCommand = new SqlCommand("Update Software set Name=@name,SoftwareType=@softwaretype,KeyNeed=@keyneed where ID_Software = " + idSoftware, Connection.sqlConnection);
                    sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@name", SqlDbType.VarChar));
                    sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@softwaretype", SqlDbType.Int));
                    sqlDataAdapter.UpdateCommand.Parameters.Add(new SqlParameter("@keyneed", SqlDbType.Bit));
                    sqlDataAdapter.UpdateCommand.Parameters["@name"].Value = textBox1.Text;
                    sqlDataAdapter.UpdateCommand.Parameters["@softwaretype"].Value = (int)comboBox3.SelectedValue;
                    sqlDataAdapter.UpdateCommand.Parameters["@keyneed"].Value = false;
                    sqlDataAdapter.UpdateCommand.ExecuteNonQuery();
                    MessageBox.Show("Запись была успешно изменена!", "Сообщение");
                }
            }
            catch
            {
                MessageBox.Show("Произошла ошибка при изменении записи!", "Сообщение");
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            rowOfDevice = e.RowIndex;
            deviceAdded = false;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells[1].Value == dataGridView2.Rows[rowOfDevice].Cells[1].Value)
                {
                    MessageBox.Show("Выбранное устройство уже добавлено!", "Сообщение");
                    deviceAdded = true;
                    break;
                }
            }
            if (deviceAdded is false)
            {
                dataGridView1.Rows.Add();
                rowOfDeviceAdd = dataGridView1.Rows.Count - 1;
                for (int i = 0; i < dataGridView2.Columns.Count; i++)
                {
                    dataGridView1.Rows[rowOfDeviceAdd].Cells[i].Value = dataGridView2.Rows[rowOfDevice].Cells[i].Value;
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            searchDevice();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            searchDevice();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            rowOfDeviceAdd = e.RowIndex;
            if (rowOfDeviceAdd != -1)
            {
                dataGridView1.Rows.RemoveAt(rowOfDeviceAdd);
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

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if ((c < 'A' || c > 'z') && c != '\b' && !Char.IsDigit(c) && c != '-' && c != 32)
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

        private void Software_FormClosing(object sender, FormClosingEventArgs e)
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

        private void searchDevice()
        {
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                dataGridView2.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                    if (dataGridView2.Rows[i].Cells[j].Value != null)
                        if (dataGridView2.Rows[i].Cells[j].Value.ToString().Contains(textBox3.Text))
                        {
                            dataGridView2.Rows[i].Selected = true;
                            break;
                        }
            }
        }

        private void textBoxLoad()
        {
            sqlDataAdapter = new SqlDataAdapter("Select ID_Software,s.Name,LicenseKey,LicenseKeyDuration,DateOfPurchase,Price,KeyNeed, s.SoftwareType from Software s join SoftwareType st on st.ID_SoftwareType=s.SoftwareType where ID_Software = " + idSoftware, Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            textBox1.Text = Convert.ToString(dataTable.Rows[0][1]);
            textBox2.Text = Convert.ToString(dataTable.Rows[0][2]);
            textBox4.Text = Convert.ToString(dataTable.Rows[0][3]);
            dateTimePicker1.Text = Convert.ToString(dataTable.Rows[0][4]);
            textBox5.Text = Convert.ToString(dataTable.Rows[0][5]);
            comboBox3.SelectedValue = Convert.ToInt16(dataTable.Rows[0][7]);
            if (Convert.ToInt16(dataTable.Rows[0][6]) == 1)
            {
                checkBox1.Checked = true;
            }
            else
            {
                checkBox1.Checked = false;
            }
        }
    }
}

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
using System.Drawing.Printing;
using System.IO;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace ComputerTechnique
{
    public partial class Main : Form
    {
        int rowOfDevice = -1;
        String reportName = "";
        public Main()
        {
            InitializeComponent();
        }

        SqlDataAdapter sqlDataAdapter;
        Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
        Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;

        //Device page
        private void button1_Click(object sender, EventArgs e)
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
        private void Main_Load(object sender, EventArgs e)
        {
            loadGridDevice();
            checkBox1.Checked = true;
            comboBox8.Enabled = false;
            comboBox9.Enabled = false;
            loadComboBox();
            comboBox8.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox9.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            dataGridView1.ReadOnly = true;
            dataGridView2.ReadOnly = true;
            dataGridView6.ReadOnly = true;
            dataGridView5.ReadOnly = true;
            dataGridView4.ReadOnly = true;
            dataGridView3.ReadOnly = true;
        }

        private void loadComboBox()
        {
            sqlDataAdapter = new SqlDataAdapter("Select ID_DeviceType, Name from DeviceType", Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            comboBox8.DataSource = dataTable;
            comboBox8.DisplayMember = "Name";
            comboBox8.ValueMember = "ID_DeviceType";
            sqlDataAdapter = new SqlDataAdapter("Select ID_DeviceStatus, Name from DeviceStatus", Connection.sqlConnection);
            DataTable dataTable1 = new DataTable();
            sqlDataAdapter.Fill(dataTable1);
            comboBox9.DataSource = dataTable1;
            comboBox9.DisplayMember = "Name";
            comboBox9.ValueMember = "ID_DeviceStatus";
            sqlDataAdapter = new SqlDataAdapter("Select ID_Department, Name from Department", Connection.sqlConnection);
            DataTable dataTable2 = new DataTable();
            sqlDataAdapter.Fill(dataTable2);
            comboBox1.DataSource = dataTable2;
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "ID_Department";
            sqlDataAdapter = new SqlDataAdapter("Select ID_Post, Name from Post", Connection.sqlConnection);
            DataTable dataTable3 = new DataTable();
            sqlDataAdapter.Fill(dataTable3);
            comboBox2.DataSource = dataTable3;
            comboBox2.DisplayMember = "Name";
            comboBox2.ValueMember = "ID_Post";
        }

        private void loadGridDevice()
        {
            string select = "";
            select = "Select ID_Device, InventoryNumber 'Инвентарный номер', d.Name Наименование, dt.Name Тип, ds.Name Статус, r.Name Помещение from Device d join DeviceType dt on dt.ID_DeviceType = d.DeviceType join DeviceStatus ds on ds.ID_DeviceStatus = d.DeviceStatus join Room r on r.ID_Room = d.Room";
            sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable;
            dataGridView1.Columns[0].Visible = false;
            rowOfDevice = -1;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            rowOfDevice = e.RowIndex;
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult result = MessageBox.Show("Закрыть программу?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                if (e.CloseReason == CloseReason.UserClosing)
                {
                    e.Cancel = true;
                }
                else
                {
                    Connection.connectClose();
                }
            }
            if (ExcelWorkBook != null)
            {
                try
                {
                    ExcelWorkBook.Save();
                    ExcelWorkBook.Close(true, Type.Missing, Type.Missing);
                    ExcelApp.Quit();
                }
                catch
                {
                    ExcelApp.Quit();
                }
            }
            else
            {
                ExcelApp.Quit();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (rowOfDevice > -1)
            {
                Device device = new Device(Convert.ToInt16(dataGridView1.Rows[rowOfDevice].Cells[0].Value), "watch");
                device.ShowDialog();
            }
            else
            {
                MessageBox.Show("Выберите запись");
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            Device device = new Device();
            device.ShowDialog();
            checkBox1.Checked = true;
            filterDevice();
        }



        private void button4_Click(object sender, EventArgs e)
        {
            if (rowOfDevice > -1)
            {
                Device device = new Device(Convert.ToInt16(dataGridView1.Rows[rowOfDevice].Cells[0].Value), "change");
                device.ShowDialog();
            }
            else
            {
                MessageBox.Show("Выберите запись");
            }
            checkBox1.Checked = true;
            filterDevice();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (rowOfDevice > -1)
            {
                DialogResult result = MessageBox.Show("Удалить запись?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        Connection.connectOpen();
                        sqlDataAdapter = new SqlDataAdapter("Select * from Device", Connection.sqlConnection);
                        sqlDataAdapter.DeleteCommand = new SqlCommand("delete from Device where ID_Device=@code", Connection.sqlConnection);
                        sqlDataAdapter.DeleteCommand.Parameters.Add(new SqlParameter("@code", SqlDbType.Int));
                        sqlDataAdapter.DeleteCommand.Parameters["@code"].Value = dataGridView1.Rows[rowOfDevice].Cells[0].Value;
                        sqlDataAdapter.DeleteCommand.ExecuteNonQuery();
                        MessageBox.Show("Запись была успешно удалена!", "Сообщение");
                    }
                    catch
                    {
                        MessageBox.Show("Произошла ошибка при удалении записи!", "Сообщение");
                    }
                }
            }
            else
            {
                MessageBox.Show("Выберите запись");
            }
            filterDevice();
        }
        private void tabPage1_Enter(object sender, EventArgs e)
        {
            loadGridDevice();
        }

        //Employee page

        int rowOfEmployee = -1;
        private void button18_Click(object sender, EventArgs e)
        {
            if (rowOfEmployee > -1)
            {
                Employee employee = new Employee(Convert.ToInt16(dataGridView4.Rows[rowOfEmployee].Cells[0].Value), "watch");
                employee.ShowDialog();
            }
            else
            {
                MessageBox.Show("Выберите запись");
            }
        }
        private void button17_Click(object sender, EventArgs e)
        {
            searchEmployee();
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            searchEmployee();
        }
        private void searchEmployee()
        {
            for (int i = 0; i < dataGridView4.RowCount; i++)
            {
                dataGridView4.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView4.ColumnCount; j++)
                    if (dataGridView4.Rows[i].Cells[j].Value != null)
                        if (dataGridView4.Rows[i].Cells[j].Value.ToString().Contains(textBox7.Text))
                        {
                            dataGridView4.Rows[i].Selected = true;
                            break;
                        }
            }
        }
        private void tabPage5_Enter(object sender, EventArgs e)
        {
            loadEmployee();
            checkBox6.Checked = true;
            comboBox1.Enabled = false;
            comboBox2.Enabled = false;
        }

        private void loadEmployee()
        {
            string select = "Select ID_Employee, e.Surname+' '+e.Name+' '+e.Patronymic ФИО,p.name Должность, d.name Отдел, e.Email 'Электронная почта' from Employee e join Post p on p.ID_Post = e.Post join Department d on d.ID_Department=e.Department";
            sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            dataGridView4.DataSource = dataTable;
            dataGridView4.Columns[0].Visible = false;
            rowOfEmployee = -1;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            Employee type = new Employee();
            type.ShowDialog();
            checkBox6.Checked = true;
            filterEmployee();
        }

        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            rowOfEmployee = e.RowIndex;
        }

        private void button21_Click(object sender, EventArgs e)
        {
            if (rowOfEmployee > -1)
            {
                Employee employee = new Employee(Convert.ToInt16(dataGridView4.Rows[rowOfEmployee].Cells[0].Value), "change");
                employee.ShowDialog();
            }
            else
            {
                MessageBox.Show("Выберите запись");
            }
            checkBox6.Checked = true;
            filterEmployee();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if (rowOfEmployee > -1)
            {
                DialogResult result = MessageBox.Show("Удалить запись?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        Connection.connectOpen();
                        sqlDataAdapter = new SqlDataAdapter("Select * from Employee", Connection.sqlConnection);
                        sqlDataAdapter.DeleteCommand = new SqlCommand("delete from Employee where ID_Employee=@code", Connection.sqlConnection);
                        sqlDataAdapter.DeleteCommand.Parameters.Add(new SqlParameter("@code", SqlDbType.Int));
                        sqlDataAdapter.DeleteCommand.Parameters["@code"].Value = dataGridView4.Rows[rowOfEmployee].Cells[0].Value;
                        sqlDataAdapter.DeleteCommand.ExecuteNonQuery();
                        MessageBox.Show("Запись была успешно удалена!", "Сообщение");
                    }
                    catch
                    {
                        MessageBox.Show("Произошла ошибка при удалении записи!\r\nИзмените ответственного сотрудника в списке устройств перед удалением его из списка сотрудников!", "Сообщение");
                    }
                }
            }
            else
            {
                MessageBox.Show("Выберите запись");
            }
            checkBox6.Checked = true;
            filterEmployee();
        }

        //Movement page

        int rowOfMovement = -1;

        private void tabPage3_Enter(object sender, EventArgs e)
        {
            loadMovement();
        }

        private void loadMovement()
        {
            string select = "Select ID_DeviceMovement,d.Name Устройство,d.InventoryNumber 'Инвентарный номер', DateOfMovement 'Дата перемещения', r1.Name 'Прежнее место',r2.Name 'Новое место',e.Surname+' '+e.Name+' '+e.Patronymic 'Ответственный сотрудник' from DeviceMovement dm join Device d on d.ID_Device = dm.Device join Employee E on e.ID_Employee = dm.Employee join Room r1 on r1.ID_Room = dm.PreviousRoom join Room r2 on r2.ID_Room = dm.NewRoom";
            sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            dataGridView6.DataSource = dataTable;
            dataGridView6.Columns[0].Visible = false;
            rowOfMovement = -1;
        }

        private void dataGridView6_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            rowOfMovement = e.RowIndex;
        }

        private void button29_Click(object sender, EventArgs e)
        {
            DeviceMovement movement = new DeviceMovement();
            movement.ShowDialog();
            loadMovement();
        }

        private void button31_Click(object sender, EventArgs e)
        {
            if (rowOfMovement > -1)
            {
                DeviceMovement movement = new DeviceMovement(Convert.ToInt16(dataGridView6.Rows[rowOfMovement].Cells[0].Value), "change");
                movement.ShowDialog();
            }
            else
            {
                MessageBox.Show("Выберите запись");
            }
            loadMovement();
        }
        private void button28_Click(object sender, EventArgs e)
        {
            if (rowOfMovement > -1)
            {
                DeviceMovement movement = new DeviceMovement(Convert.ToInt16(dataGridView6.Rows[rowOfMovement].Cells[0].Value), "watch");
                movement.ShowDialog();
            }
            else
            {
                MessageBox.Show("Выберите запись");
            }
        }

        private void button30_Click(object sender, EventArgs e)
        {
            if (rowOfMovement > -1)
            {
                DialogResult result = MessageBox.Show("Удалить запись?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        Connection.connectOpen();
                        sqlDataAdapter = new SqlDataAdapter("Select * from DeviceMovement", Connection.sqlConnection);
                        sqlDataAdapter.DeleteCommand = new SqlCommand("delete from DeviceMovement where ID_DeviceMovement=@code", Connection.sqlConnection);
                        sqlDataAdapter.DeleteCommand.Parameters.Add(new SqlParameter("@code", SqlDbType.Int));
                        sqlDataAdapter.DeleteCommand.Parameters["@code"].Value = dataGridView6.Rows[rowOfMovement].Cells[0].Value;
                        sqlDataAdapter.DeleteCommand.ExecuteNonQuery();
                        MessageBox.Show("Запись была успешно удалена!", "Сообщение");
                    }
                    catch
                    {
                        MessageBox.Show("Произошла ошибка при удалении записи!", "Сообщение");
                    }
                }
            }
            else
            {
                MessageBox.Show("Выберите запись");
            }
            loadMovement();
        }

        private void button27_Click(object sender, EventArgs e)
        {
            searchMovement();
        }
        private void searchMovement()
        {
            for (int i = 0; i < dataGridView6.RowCount; i++)
            {
                dataGridView6.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView6.ColumnCount; j++)
                    if (dataGridView6.Rows[i].Cells[j].Value != null)
                        if (dataGridView6.Rows[i].Cells[j].Value.ToString().Contains(textBox9.Text))
                        {
                            dataGridView6.Rows[i].Selected = true;
                            break;
                        }
            }
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            searchMovement();
        }

        //repair page

        int rowOfRepair = -1;
        private void button24_Click(object sender, EventArgs e)
        {
            Repair repair = new Repair();
            repair.ShowDialog();
            loadRepair();
        }

        private void tabPage4_Enter(object sender, EventArgs e)
        {
            loadRepair();
        }
        private void loadRepair()
        {
            string select = "Select ID_RepairWork,d.InventoryNumber 'Инветарный номер', d.Name 'Название устройства',StartOfWork 'Дата начала работ', EndOfWork 'Дата окончания работ',e.Surname+' '+e.Name+' '+e.Patronymic Мастер from RepairWork rw join Employee e on e.ID_Employee = rw.Master join Device d on d.ID_Device = rw.Device";
            sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            dataGridView5.DataSource = dataTable;
            dataGridView5.Columns[0].Visible = false;
            rowOfMovement = -1;
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            searchRepair();
        }

        private void button22_Click(object sender, EventArgs e)
        {
            searchRepair();
        }
        private void searchRepair()
        {
            for (int i = 0; i < dataGridView5.RowCount; i++)
            {
                dataGridView5.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView5.ColumnCount; j++)
                    if (dataGridView5.Rows[i].Cells[j].Value != null)
                        if (dataGridView5.Rows[i].Cells[j].Value.ToString().Contains(textBox8.Text))
                        {
                            dataGridView5.Rows[i].Selected = true;
                            break;
                        }
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            if (rowOfRepair > -1)
            {
                DialogResult result = MessageBox.Show("Удалить запись?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        Connection.connectOpen();
                        sqlDataAdapter = new SqlDataAdapter("Select * from RepairWork", Connection.sqlConnection);
                        sqlDataAdapter.DeleteCommand = new SqlCommand("delete from RepairWork where ID_RepairWork=@code", Connection.sqlConnection);
                        sqlDataAdapter.DeleteCommand.Parameters.Add(new SqlParameter("@code", SqlDbType.Int));
                        sqlDataAdapter.DeleteCommand.Parameters["@code"].Value = dataGridView5.Rows[rowOfRepair].Cells[0].Value;
                        sqlDataAdapter.DeleteCommand.ExecuteNonQuery();
                        MessageBox.Show("Запись была успешно удалена!", "Сообщение");
                    }
                    catch
                    {
                        MessageBox.Show("Произошла ошибка при удалении записи!", "Сообщение");
                    }
                }
            }
            else
            {
                MessageBox.Show("Выберите запись");
            }
            loadRepair();
        }

        private void dataGridView5_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            rowOfRepair = e.RowIndex;
        }

        private void button26_Click(object sender, EventArgs e)
        {
            if (rowOfRepair > -1)
            {
                Repair repair = new Repair(Convert.ToInt16(dataGridView5.Rows[rowOfRepair].Cells[0].Value), "change");
                repair.ShowDialog();
            }
            else
            {
                MessageBox.Show("Выберите запись");
            }
            loadRepair();
        }

        private void button23_Click(object sender, EventArgs e)
        {
            if (rowOfRepair > -1)
            {
                Repair repair = new Repair(Convert.ToInt16(dataGridView5.Rows[rowOfRepair].Cells[0].Value), "watch");
                repair.ShowDialog();
            }
            else
            {
                MessageBox.Show("Выберите запись");
            }
            loadRepair();
        }

        //software page

        int rowOfSoftware = -1;
        private void tabPage2_Enter(object sender, EventArgs e)
        {
            loadSoftware();
            radioButton10.Checked = true;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            searchSoftware();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            searchSoftware();
        }
        private void searchSoftware()
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

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            rowOfSoftware = e.RowIndex;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Software software = new Software();
            software.ShowDialog();
            radioButton10.Checked = true;
            loadSoftware();
        }

        private void loadSoftware()
        {
            string select = "Select ID_Software,s.Name Название, st.Name 'Тип ПО', KeyNeed 'Необходимость лицензионного ключа', DateOfPurchase'Дата приобретения', LicenseKeyDuration 'Длительность лицензии' from Software s join SoftwareType st on st.ID_SoftwareType = s.SoftwareType";
            sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            dataGridView2.DataSource = dataTable;
            dataGridView2.Columns[0].Visible = false;
            rowOfSoftware = -1;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (rowOfSoftware > -1)
            {
                Software software = new Software(Convert.ToInt16(dataGridView2.Rows[rowOfSoftware].Cells[0].Value), "change");
                software.ShowDialog();
            }
            else
            {
                MessageBox.Show("Выберите запись");
            }
            radioButton10.Checked = true;
            loadSoftware();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (rowOfSoftware > -1)
            {
                DialogResult result = MessageBox.Show("Удалить запись?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        Connection.connectOpen();
                        sqlDataAdapter = new SqlDataAdapter("Select * from Software", Connection.sqlConnection);
                        sqlDataAdapter.DeleteCommand = new SqlCommand("delete from Software where ID_Software=@code", Connection.sqlConnection);
                        sqlDataAdapter.DeleteCommand.Parameters.Add(new SqlParameter("@code", SqlDbType.Int));
                        sqlDataAdapter.DeleteCommand.Parameters["@code"].Value = dataGridView2.Rows[rowOfSoftware].Cells[0].Value;
                        sqlDataAdapter.DeleteCommand.ExecuteNonQuery();
                        MessageBox.Show("Запись была успешно удалена!", "Сообщение");
                    }
                    catch
                    {
                        MessageBox.Show("Произошла ошибка при удалении записи!", "Сообщение");
                    }
                }
            }
            else
            {
                MessageBox.Show("Выберите запись");
            }
            radioButton10.Checked = true;
            loadSoftware();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (rowOfSoftware > -1)
            {
                Software software = new Software(Convert.ToInt16(dataGridView2.Rows[rowOfSoftware].Cells[0].Value), "watch");
                software.ShowDialog();
            }
            else
            {
                MessageBox.Show("Выберите запись");
            }
        }

        //filters
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked is true)
            {
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox2.Enabled = false;
                checkBox3.Enabled = false;
                comboBox8.Enabled = false;
                comboBox9.Enabled = false;
            }
            else
            {
                checkBox2.Enabled = true;
                checkBox3.Enabled = true;
            }
        }


        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked is true)
            {
                comboBox9.Enabled = true;
            }
            else
            {
                comboBox9.Enabled = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked is true)
            {
                comboBox8.Enabled = true;
            }
            else
            {
                comboBox8.Enabled = false;
            }
        }
        private void filterDevice()
        {
            if (checkBox1.Checked is true)
            {
                loadGridDevice();
            }
            else
            {
                if (checkBox2.Checked is true && checkBox3.Checked is true)
                {
                    string select = "";
                    select = "Select ID_Device, InventoryNumber 'Инвентарный номер', d.Name Наименование, dt.Name Тип, ds.Name Статус, r.Name Помещение from Device d join DeviceType dt on dt.ID_DeviceType = d.DeviceType join DeviceStatus ds on ds.ID_DeviceStatus = d.DeviceStatus join Room r on r.ID_Room = d.Room where d.DeviceStatus = " + comboBox9.SelectedValue + " and d.DeviceType = " + comboBox8.SelectedValue.ToString();
                    sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
                    DataTable dataTable = new DataTable();
                    sqlDataAdapter.Fill(dataTable);
                    dataGridView1.DataSource = dataTable;
                    dataGridView1.Columns[0].Visible = false;
                    rowOfDevice = -1;
                }
                else
                {
                    if (checkBox2.Checked is true)
                    {
                        string select = "";
                        select = "Select ID_Device, InventoryNumber 'Инвентарный номер', d.Name Наименование, dt.Name Тип, ds.Name Статус, r.Name Помещение from Device d join DeviceType dt on dt.ID_DeviceType = d.DeviceType join DeviceStatus ds on ds.ID_DeviceStatus = d.DeviceStatus join Room r on r.ID_Room = d.Room where d.DeviceStatus = " + comboBox9.SelectedValue.ToString();
                        sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
                        DataTable dataTable = new DataTable();
                        sqlDataAdapter.Fill(dataTable);
                        dataGridView1.DataSource = dataTable;
                        dataGridView1.Columns[0].Visible = false;
                        rowOfDevice = -1;
                    }
                    else
                    {
                        if (checkBox3.Checked is true)
                        {
                            string select = "";
                            select = "Select ID_Device, InventoryNumber 'Инвентарный номер', d.Name Наименование, dt.Name Тип, ds.Name Статус, r.Name Помещение from Device d join DeviceType dt on dt.ID_DeviceType = d.DeviceType join DeviceStatus ds on ds.ID_DeviceStatus = d.DeviceStatus join Room r on r.ID_Room = d.Room where d.DeviceType = " + comboBox8.SelectedValue.ToString();
                            sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
                            DataTable dataTable = new DataTable();
                            sqlDataAdapter.Fill(dataTable);
                            dataGridView1.DataSource = dataTable;
                            dataGridView1.Columns[0].Visible = false;
                            rowOfDevice = -1;
                        }
                        else
                        {
                            dataGridView1.DataSource = null;
                            rowOfDevice = -1;
                        }
                    }
                }
            }
        }
        private void button7_Click(object sender, EventArgs e)
        {
            filterDevice();
        }

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton10.Checked is true)
            {
                loadSoftware();
            }
        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {
            string select = "Select  ID_Software,s.Name Название, st.Name 'Тип ПО', KeyNeed 'Необходимость лицензионного ключа', DateOfPurchase'Дата приобретения', LicenseKeyDuration 'Длительность лицензии' from Software s join SoftwareType st on st.ID_SoftwareType = s.SoftwareType join SoftwareInstallation si on s.ID_Software = si.Software  group by ID_Software,s.Name, st.Name, KeyNeed, DateOfPurchase, LicenseKeyDuration";
            sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            dataGridView2.DataSource = dataTable;
            dataGridView2.Columns[0].Visible = false;
            rowOfSoftware = -1;
        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            string select = "Select  ID_Software,s.Name Название, st.Name 'Тип ПО', KeyNeed 'Необходимость лицензионного ключа', DateOfPurchase'Дата приобретения', LicenseKeyDuration 'Длительность лицензии' from Software s join SoftwareType st on st.ID_SoftwareType = s.SoftwareType  left join SoftwareInstallation si on s.ID_Software = si.Software  WHERE si.Software is NULL  group by ID_Software, s.Name, st.Name, KeyNeed, DateOfPurchase, LicenseKeyDuration";
            sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            dataGridView2.DataSource = dataTable;
            dataGridView2.Columns[0].Visible = false;
            rowOfSoftware = -1;
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            string select = "Select  ID_Software,s.Name Название, st.Name 'Тип ПО', KeyNeed 'Необходимость лицензионного ключа', DateOfPurchase'Дата приобретения', LicenseKeyDuration 'Длительность лицензии' from Software s join SoftwareType st on st.ID_SoftwareType = s.SoftwareType where DATEADD(month, LicenseKeyDuration, DateOfPurchase) < GETDATE()";
            sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            dataGridView2.DataSource = dataTable;
            dataGridView2.Columns[0].Visible = false;
            rowOfSoftware = -1;
        }

        //print and excel

        private void button11_Click(object sender, EventArgs e)
        {
            printGrid();
        }
        private void printDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            if (tabControl1.SelectedTab == tabPage1)
            {
                Bitmap bmp = new Bitmap(dataGridView1.Size.Width, dataGridView1.Size.Height);
                dataGridView1.DrawToBitmap(bmp, dataGridView1.Bounds);
                e.Graphics.DrawImage(bmp, 0, 0);
            }
            else
            {
                if (tabControl1.SelectedTab == tabPage2)
                {
                    Bitmap bmp = new Bitmap(dataGridView2.Size.Width, dataGridView2.Size.Height);
                    dataGridView2.DrawToBitmap(bmp, dataGridView2.Bounds);
                    e.Graphics.DrawImage(bmp, 0, 0);
                }
                else
                {
                    if (tabControl1.SelectedTab == tabPage3)
                    {
                        Bitmap bmp = new Bitmap(dataGridView6.Size.Width, dataGridView6.Size.Height);
                        dataGridView6.DrawToBitmap(bmp, dataGridView6.Bounds);
                        e.Graphics.DrawImage(bmp, 0, 0);
                    }
                    else
                    {
                        if (tabControl1.SelectedTab == tabPage4)
                        {
                            Bitmap bmp = new Bitmap(dataGridView5.Size.Width, dataGridView5.Size.Height);
                            dataGridView5.DrawToBitmap(bmp, dataGridView5.Bounds);
                            e.Graphics.DrawImage(bmp, 0, 0);
                        }
                        else
                        {
                            if (tabControl1.SelectedTab == tabPage5)
                            {
                                Bitmap bmp = new Bitmap(dataGridView4.Size.Width, dataGridView4.Size.Height);
                                dataGridView4.DrawToBitmap(bmp, dataGridView4.Bounds);
                                e.Graphics.DrawImage(bmp, 0, 0);
                            }
                            else
                            {
                                if (tabControl1.SelectedTab == tabPage7)
                                {
                                    Bitmap bmp = new Bitmap(dataGridView3.Size.Width, dataGridView3.Size.Height);
                                    dataGridView3.DrawToBitmap(bmp, dataGridView3.Bounds);
                                    e.Graphics.DrawImage(bmp, 0, 0);
                                }
                            }
                        }
                    }
                }
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            for (int j = 1; j < dataGridView1.ColumnCount; j++)
            {
                ExcelApp.Cells[1, j] = dataGridView1.Columns[j].HeaderText.ToString();
            }
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 1; j < dataGridView1.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            excelTableBorders((dataGridView1.Rows.Count + 1).ToString(), (dataGridView1.ColumnCount-1).ToString());
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }
        private void printGrid()
        {
            PrintDocument printDocument = new PrintDocument();
            printDocument.PrintPage += new PrintPageEventHandler(printDocument_PrintPage);
            PrintDialog printDialog = new PrintDialog();
            printDialog.Document = printDocument;
            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                printDialog.Document.Print();
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            printGrid();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            for (int j = 1; j < dataGridView2.ColumnCount; j++)
            {
                ExcelApp.Cells[1, j] = dataGridView2.Columns[j].HeaderText.ToString();
            }
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                for (int j = 1; j < dataGridView2.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j] = dataGridView2.Rows[i].Cells[j].Value;
                }
            }
            excelTableBorders((dataGridView2.Rows.Count + 1).ToString(), (dataGridView2.ColumnCount - 1).ToString());
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void button32_Click(object sender, EventArgs e)
        {
            printGrid();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            for (int j = 1; j < dataGridView6.ColumnCount; j++)
            {
                ExcelApp.Cells[1, j] = dataGridView6.Columns[j].HeaderText.ToString();
            }
            for (int i = 0; i < dataGridView6.Rows.Count; i++)
            {
                for (int j = 1; j < dataGridView6.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j] = dataGridView6.Rows[i].Cells[j].Value;
                }
            }
            excelTableBorders((dataGridView6.Rows.Count + 1).ToString(), (dataGridView6.ColumnCount - 1).ToString());
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void button34_Click(object sender, EventArgs e)
        {
            printGrid();
        }

        private void button33_Click(object sender, EventArgs e)
        {
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            for (int j = 1; j < dataGridView5.ColumnCount; j++)
            {
                ExcelApp.Cells[1, j] = dataGridView5.Columns[j].HeaderText.ToString();
            }
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                for (int j = 1; j < dataGridView5.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j] = dataGridView5.Rows[i].Cells[j].Value;
                }
            }
            excelTableBorders((dataGridView5.Rows.Count + 1).ToString(), (dataGridView5.ColumnCount - 1).ToString());
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void button36_Click(object sender, EventArgs e)
        {
            printGrid();
        }

        private void button35_Click(object sender, EventArgs e)
        {
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            for (int j = 1; j < dataGridView4.ColumnCount; j++)
            {
                ExcelApp.Cells[1, j] = dataGridView4.Columns[j].HeaderText.ToString();
            }
            for (int i = 0; i < dataGridView4.Rows.Count; i++)
            {
                for (int j = 1; j < dataGridView4.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j] = dataGridView4.Rows[i].Cells[j].Value;
                }
            }
            excelTableBorders((dataGridView4.Rows.Count + 1).ToString(), (dataGridView4.ColumnCount - 1).ToString());
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void button37_Click(object sender, EventArgs e)
        {
            string select = "";            
            DataTable dataTableAdded = new DataTable();
            int rowOfSoftwareAdd = 0;
            dataGridView3.Rows.Clear();
            dataGridView3.Columns.Clear();
            dataGridView3.Columns.Add("Программное обеспечение", "Программное обеспечение");
            dataGridView3.Columns.Add("Тип ПО", "Тип ПО");
            dataGridView3.Columns.Add("Инвентарный номер устройства", "Инвентарный номер устройства");
            dataGridView3.Columns.Add("Название устройства", "Название устройства");
            dataGridView3.Columns.Add("Тип устройства", "Тип устройства");
            select = " Select s.Name 'Программное обеспечение', st.Name 'Тип ПО',d.InventoryNumber 'Инвентарный номер устройства', d.Name 'Название устройства', dt.Name 'Тип устройства'  from SoftwareInstallation si  right join Software s on s.ID_Software = si.Software  left join Device d on d.ID_Device = si.Device  left join SoftwareType st on st.ID_SoftwareType = s.SoftwareType  left join DeviceType dt on dt.ID_DeviceType = d.DeviceType";
            sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
            sqlDataAdapter.Fill(dataTableAdded);
            for (int g = 0; g < dataTableAdded.Rows.Count; g++)
            {
                dataGridView3.Rows.Add();
                rowOfSoftwareAdd = dataGridView3.Rows.Count - 1;
                for (int i = 0; i < dataGridView3.Columns.Count; i++)
                {
                    dataGridView3.Rows[rowOfSoftwareAdd].Cells[i].Value = dataTableAdded.Rows[g][i].ToString();
                }
            }
            for (int i = dataGridView3.Rows.Count-1; i > 0; i--)
            {                
                if (dataGridView3.Rows[i].Cells[0].Value.ToString() == dataGridView3.Rows[i-1].Cells[0].Value.ToString())
                {
                    dataGridView3.Rows[i].Cells[0].Value = "";
                    dataGridView3.Rows[i].Cells[1].Value = "";
                }
            }
            reportName = "SoftwareAndDevices";
        }

        private void button39_Click(object sender, EventArgs e)
        {
            printGrid();
        }

        private void button38_Click(object sender, EventArgs e)
        {
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            for (int j = 0; j < dataGridView3.ColumnCount; j++)
            {
                ExcelApp.Cells[1, j+1] = dataGridView3.Columns[j].HeaderText.ToString();
            }
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView3.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j+1] = dataGridView3.Rows[i].Cells[j].Value;
                }
            }
            excelTableBorders((dataGridView3.Rows.Count+1).ToString(), dataGridView3.ColumnCount.ToString());
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
            if (reportName == "SoftwareAndDevices")
            {
                ExcelWorkSheet.get_Range("C2", "E" + (dataGridView3.Rows.Count + 1).ToString()).Interior.Color = Color.LightGray;
            }
            else
            {
                ExcelWorkSheet.get_Range("D2", "E" + (dataGridView3.Rows.Count + 1).ToString()).Interior.Color = Color.LightGray;
            }
        }
        private void excelTableBorders(String rowsCount, String columnCount)
        {
            String columnName = "";
            switch (columnCount)
            {
                case "3":
                    columnName = "C";
                    break;
                case "4":
                    columnName = "D";
                    break;
                case "5":
                    columnName = "E";
                    break;
                case "6":
                    columnName = "F";
                    break;
                default:
                    columnName = "G";
                    break;
            }
            ExcelWorkSheet.Columns.ColumnWidth = 30;
            var cells = ExcelWorkSheet.get_Range("A1", columnName + rowsCount);
            cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            cells.HorizontalAlignment = -4131;
            ExcelWorkSheet.get_Range("A1", columnName + "1").HorizontalAlignment = -4108;
            ExcelWorkSheet.get_Range("A1", columnName + "1").Interior.Color = Color.DarkGray;
            ExcelWorkSheet.get_Range("A1", columnName + "1").Font.Bold = true;
        }

        private void button40_Click(object sender, EventArgs e)
        {
            string select = "";
            DataTable dataTableAdded = new DataTable();
            int rowOfSoftwareAdd = 0;
            dataGridView3.Rows.Clear();
            dataGridView3.Columns.Clear();
            dataGridView3.Columns.Add("Инвентарный номер устройства", "Инвентарный номер устройства");
            dataGridView3.Columns.Add("Название устройства", "Название устройства");
            dataGridView3.Columns.Add("Тип устройства", "Тип устройства");
            dataGridView3.Columns.Add("Программное обеспечение", "Программное обеспечение");
            dataGridView3.Columns.Add("Тип ПО", "Тип ПО");
            select = " Select d.InventoryNumber 'Инвентарный номер устройства', d.Name 'Название устройства', dt.Name 'Тип устройства', s.Name 'Программное обеспечение', st.Name 'Тип ПО' from SoftwareInstallation si   right join Device d on d.ID_Device = si.Device  left join DeviceType dt on dt.ID_DeviceType = d.DeviceType  left join Software s on s.ID_Software = si.Software  left join SoftwareType st on st.ID_SoftwareType = s.SoftwareType  ";
            sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
            sqlDataAdapter.Fill(dataTableAdded);
            for (int g = 0; g < dataTableAdded.Rows.Count; g++)
            {
                dataGridView3.Rows.Add();
                rowOfSoftwareAdd = dataGridView3.Rows.Count - 1;
                for (int i = 0; i < dataGridView3.Columns.Count; i++)
                {
                    dataGridView3.Rows[rowOfSoftwareAdd].Cells[i].Value = dataTableAdded.Rows[g][i].ToString();
                }
            }
            for (int i = dataGridView3.Rows.Count - 1; i > 0; i--)
            {
                if (dataGridView3.Rows[i].Cells[0].Value.ToString() == dataGridView3.Rows[i - 1].Cells[0].Value.ToString())
                {
                    dataGridView3.Rows[i].Cells[0].Value = "";
                    dataGridView3.Rows[i].Cells[1].Value = "";
                    dataGridView3.Rows[i].Cells[2].Value = "";
                }
            }
            reportName = "DevicesAndSoftware";
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked is true)
            {
                checkBox5.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Enabled = false;
                checkBox4.Enabled = false;
                comboBox1.Enabled = false;
                comboBox2.Enabled = false;
            }
            else
            {
                checkBox5.Enabled = true;
                checkBox4.Enabled = true;
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked is true)
            {
                comboBox1.Enabled = true;
            }
            else
            {
                comboBox1.Enabled = false;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked is true)
            {
                comboBox2.Enabled = true;
            }
            else
            {
                comboBox2.Enabled = false;
            }
        }

        private void button41_Click(object sender, EventArgs e)
        {
            filterEmployee();
        }

        private void filterEmployee()
        {
            if (checkBox6.Checked is true)
            {
                loadEmployee();
            }
            else
            {
                if (checkBox5.Checked is true && checkBox4.Checked is true)
                {
                    string select = "";
                    select = "Select ID_Employee, e.Surname+' '+e.Name+' '+e.Patronymic ФИО,p.name Должность, d.name Отдел, e.Email 'Электронная почта' from Employee e join Post p on p.ID_Post = e.Post join Department d on d.ID_Department=e.Department where e.Department = " + comboBox1.SelectedValue + " and e.Post = " + comboBox2.SelectedValue.ToString();
                    sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
                    DataTable dataTable = new DataTable();
                    sqlDataAdapter.Fill(dataTable);
                    dataGridView4.DataSource = dataTable;
                    dataGridView4.Columns[0].Visible = false;
                    rowOfDevice = -1;
                }
                else
                {
                    if (checkBox5.Checked is true)
                    {
                        string select = "";
                        select = "Select ID_Employee, e.Surname+' '+e.Name+' '+e.Patronymic ФИО,p.name Должность, d.name Отдел, e.Email 'Электронная почта' from Employee e join Post p on p.ID_Post = e.Post join Department d on d.ID_Department=e.Department where e.Department = " + comboBox1.SelectedValue.ToString();
                        sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
                        DataTable dataTable = new DataTable();
                        sqlDataAdapter.Fill(dataTable);
                        dataGridView4.DataSource = dataTable;
                        dataGridView4.Columns[0].Visible = false;
                        rowOfDevice = -1;
                    }
                    else
                    {
                        if (checkBox4.Checked is true)
                        {
                            string select = "";
                            select = "Select ID_Employee, e.Surname+' '+e.Name+' '+e.Patronymic ФИО,p.name Должность, d.name Отдел, e.Email 'Электронная почта' from Employee e join Post p on p.ID_Post = e.Post join Department d on d.ID_Department=e.Department where e.Post = " + comboBox2.SelectedValue.ToString();
                            sqlDataAdapter = new SqlDataAdapter(select, Connection.sqlConnection);
                            DataTable dataTable = new DataTable();
                            sqlDataAdapter.Fill(dataTable);
                            dataGridView4.DataSource = dataTable;
                            dataGridView4.Columns[0].Visible = false;
                            rowOfDevice = -1;
                        }
                        else
                        {
                            dataGridView4.DataSource = null;
                            rowOfDevice = -1;
                        }
                    }
                }
            }
        }
    }
}

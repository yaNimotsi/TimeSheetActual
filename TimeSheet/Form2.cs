using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;

namespace TimeSheet
{
    public partial class Form2 : Form
    {
        public string UId,DepId, UIdT, myDate, myTime, PId, TId, conString;
        private int UserID, TimerVal;
        public DataTable Dt = new DataTable();
        public Form2(string UId, string conString, string DepId)
        {
            this.UId = UId;
            this.DepId = DepId;
            UserID = Convert.ToInt32(UId);
            this.conString = conString;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            TreeNode treeNode = treeView2.SelectedNode;
            int level = treeNode.Level;

            string PuthToNode = treeNode.FullPath;
            string NameSelProj = "";
            string TypeBild = "";

            if (level == 1)
            {
                NameSelProj = treeNode.Text;
                TypeBild = "";
            }
            else if (level > 1)
            {
                var ArrVal = PuthToNode.Split('\\');
                NameSelProj = ArrVal[1];
                TypeBild = treeView2.SelectedNode.Text;
            }
            else
            {
                NameSelProj = "";
                TypeBild = "";
            }

            PId = "";
            PId = NameProj(treeView1);

            if (PId == "") MessageBox.Show("Не выбрана площадка или сооружение для проекта");
            else if ((TypeBild.Length != 0) && (comboBox1.Text.Length != 0) && (comboBox2.Text.Length != 0) && (textBox4.Text.Length != 0))//((textBox1.Text.Length != 0) && (comboBox2.Text.Length != 0) && (comboBox3.Text.Length != 0) && (comboBox6.Text.Length != 0))
            {
                string WorkTimeVal = textBox4.Text.ToString().Replace('.', ',');
                float FWorkTimeVal = Convert.ToSingle(WorkTimeVal);

                DateTime dtime = DateTime.Now;
                string myDate = dtime.ToString("d");
                myDate = myDate.Replace('.', '-');
                string myTime = dtime.ToLongTimeString();

                string Section = comboBox2.Text.ToString();
                string WorkTime = textBox4.Text.ToString();
                string coment = richTextBox1.Text.ToString();
                string CountSheet = textBox1.Text.ToString();
                string TypeWork = comboBox1.Text.ToString();

                SqlConnection DbConn = new SqlConnection(conString);
                SqlCommand DbComandInsert = new SqlCommand();

                DbComandInsert.CommandText = "INSERT INTO dbo.Report(Id_Project, Id_user, TypeWork, Section, TypeBild, Comment, CountSheet, TimeWork, DateEntered, PuthToNode)"
                + " VALUES"
                + " (@pid, @uid, @TWork, @Section, @TBild, @Coment, @cSheet, @WorkTime, @Dtime, @PuthToNode)";

                DbComandInsert.Connection = DbConn;

                SqlParameter pid = new SqlParameter("@pid", PId);
                SqlParameter uid = new SqlParameter("@uid", UId);
                SqlParameter typewWork = new SqlParameter("@TWork", TypeWork);
                SqlParameter section = new SqlParameter("@Section", Section);
                SqlParameter tBild = new SqlParameter("@TBild", TypeBild);
                SqlParameter Coment = new SqlParameter("@Coment", coment);
                SqlParameter cSheet = new SqlParameter("@cSheet", CountSheet);
                SqlParameter workTime = new SqlParameter("@WorkTime", WorkTime);
                SqlParameter Dtime = new SqlParameter("@Dtime", dtime);
                SqlParameter timeEntered = new SqlParameter("@timeEntered", myTime);
                SqlParameter puthToNode = new SqlParameter("@PuthToNode", PuthToNode);

                DbComandInsert.Parameters.Add(pid);
                DbComandInsert.Parameters.Add(uid);
                DbComandInsert.Parameters.Add(typewWork);
                DbComandInsert.Parameters.Add(section);
                DbComandInsert.Parameters.Add(tBild);
                DbComandInsert.Parameters.Add(Coment);
                DbComandInsert.Parameters.Add(cSheet);
                DbComandInsert.Parameters.Add(workTime);
                DbComandInsert.Parameters.Add(Dtime);
                DbComandInsert.Parameters.Add(timeEntered);
                DbComandInsert.Parameters.Add(puthToNode);

                try
                {
                    DbConn.Open();

                    DbComandInsert.ExecuteNonQuery();
                    MessageBox.Show("Запись добавленна в базу данных", "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    comboBox1.SelectedIndex = -1;
                    comboBox2.SelectedIndex = -1;
                    richTextBox1.Text = "";
                    textBox1.Text = "";
                    textBox4.Text = "";
                }
                catch (Exception ex) { MessageBox.Show("Запись добавить не удалось!!! " + ex, "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                finally { DbConn.Close(); }
            }
            else MessageBox.Show("Не заполнен один из параметров! Заполните все и повторите попытку", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        }

        public string NameProj(TreeView treeView)
        {
            string rez = "";

            TreeNode SelectNode = treeView.SelectedNode;
            string path = treeView.SelectedNode.FullPath;

            string[] ArrVal = path.Split('\\');
            if (ArrVal[1].Length != 0) rez = ArrVal[1];

            SqlConnection conn = new SqlConnection(conString);
            SqlCommand sqlCommand = new SqlCommand();
            SqlParameter Rez = new SqlParameter("@rez", rez);
            sqlCommand.Connection = conn;
            sqlCommand.Parameters.Add(Rez);
            sqlCommand.CommandText = "SELECT Id_project FROM dbo.Proekti WHERE (Name_Project = @rez)";

            try
            {
                conn.Open();
                SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
                while (sqlDataReader.Read())
                {
                    rez = sqlDataReader[0].ToString();
                }
            }
            catch (SqlException ex) { MessageBox.Show("Произошла ошибка выполнения команды: " + ex, "Ошибка выполнения сценария", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            return rez;
        }

        //Событие закрытия формы
        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                if (MessageBox.Show("Вы действительно хотите закрыть форму ? ", "Внимание", MessageBoxButtons.YesNo)==DialogResult.Yes)
                {
                    Application.Exit();
                }
            }
        }

        private void AutoSizeGridColumn(DataGridView myGrid)
        {
            myGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        //События происходящие при загрузке формы 2
        private void Form2_Load(object sender, EventArgs e)
        { 
            SqlConnection sqldbConnection = new SqlConnection(conString);
            SqlCommand sqlCommand = new SqlCommand();
            SqlCommand EnteredVal = new SqlCommand();
            
            DataGridView GridTab2 = dataGridView2;
            
            AutoSizeGridColumn(GridTab2);

            sqlCommand.Connection = sqldbConnection;
            sqlCommand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet, "
                         +" dbo.Report.TimeWork, dbo.Report.DateEntered"
            +" FROM            dbo.Report INNER JOIN"
                         +" dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                         +" dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
            + " WHERE(dbo.Users.Id_users = @userID)";

            SqlParameter userID = new SqlParameter("@userID", UserID);
            sqlCommand.Parameters.Add(userID);

            UpdateGrid(sqlCommand);

            SqlParameter userID2 = new SqlParameter("@UId", UserID);
            
            EnteredVal.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet, "
                         + " dbo.Report.TimeWork, dbo.Report.DateEntered"
            + " FROM            dbo.Report INNER JOIN"
                         + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                         + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
            + " WHERE(dbo.Users.Id_users = @UId)";

            EnteredVal.Parameters.Add(userID2);
            EnteredVal.Connection = sqldbConnection;

            UpdateGrid(EnteredVal, GridTab2);

            DataToListBox();
            
            GridHeaderName(GridTab2);

            DateTime myDate = DateTime.Today;
            
            dateTimePicker5.Value = myDate;
            dateTimePicker6.Value = myDate;
            
            checkBox6.Checked = false;
            checkBox7.Checked = false;
            
            dateTimePicker5.Enabled = false;
            dateTimePicker6.Enabled = false;

            myTreeViewWork(treeView1);
            myTreeViewWork(treeView2);
        }

        public void UpdateGrid(SqlCommand sqlCommand)
        {
            SqlConnection conn = sqlCommand.Connection;
            SqlCommand myComand = sqlCommand;
            {
                try
                {
                    conn.Open();
                    SqlDataAdapter da = new SqlDataAdapter(myComand);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView2.DataSource = dt; //имя грида
                    conn.Close();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message);
                    return;   //или не нужно
                }
            }
        }

        private void GridHeaderName(DataGridView myGrid)
        {
            /*myGrid.Columns[0].HeaderText = "Название проекта";
            myGrid.Columns[1].HeaderText = "Текст задания";
            myGrid.Columns[2].HeaderText = "Тип работы";
            myGrid.Columns[3].HeaderText = "ФИО";
            myGrid.Columns[4].HeaderText = "Затраченное время";
            myGrid.Columns[5].HeaderText = "Дата заполнения";
            myGrid.Columns[6].HeaderText = "Коментарий";
            myGrid.Columns[7].HeaderText = "Процент выполнения";*/
            myGrid.Columns[0].HeaderText = "Название проекта";
            myGrid.Columns[1].HeaderText = "Тип здания";
            myGrid.Columns[2].HeaderText = "Тип работы";
            myGrid.Columns[3].HeaderText = "Раздел";
            myGrid.Columns[4].HeaderText = "ФИО";
            myGrid.Columns[5].HeaderText = "Коментарий";
            myGrid.Columns[6].HeaderText = "Число страниц";
            myGrid.Columns[7].HeaderText = "Затраченое время";
            myGrid.Columns[8].HeaderText = "Дата заполнения";
        }

        public void DataToListBox()
        {
            SqlConnection sqldbConnection = new SqlConnection(conString);
            SqlCommand sqlCommand = new SqlCommand();
            SqlCommand sqlBilding = new SqlCommand();
            SqlCommand sqlTypeWork = new SqlCommand();
            SqlCommand sqlSection = new SqlCommand();

            sqlCommand.Connection = sqldbConnection;
            sqlCommand.CommandText = "SELECT Id_project, Name_Project"
                    + " FROM dbo.Proekti"
                + " WHERE(InVork = 1)";

            SqlParameter userid = new SqlParameter("@userid", UserID);
            SqlParameter projId = new SqlParameter("@PId", PId);
            sqlCommand.Parameters.Add(userid);

            sqlBilding.CommandText = "SELECT Surename, department FROM dbo.Users GROUP BY Surename, department HAVING(department = @DepId)";
            SqlParameter depId2 = new SqlParameter("@DepId", DepId);
            sqlBilding.Parameters.Add(depId2);
            sqlBilding.Connection = sqldbConnection;


            sqlTypeWork.CommandText = "SELECT dbo.TypeWork.TypeWork FROM dbo.TypeWork";
            sqlSection.CommandText = "SELECT dbo.Sections.Section_Name FROM dbo.Sections";
            
            sqlTypeWork.Connection = sqldbConnection;
            sqlSection.Connection = sqldbConnection;

            try
            {
                sqldbConnection.Open();
                SqlDataReader ProjReader = sqlCommand.ExecuteReader();
                while (ProjReader.Read())
                {
                    string s = ProjReader.GetString(ProjReader.GetOrdinal("Name_Project"));
                    comboBox6.Items.Add(s);
                }
                ProjReader.Close();

                /*SqlDataReader BildReader = sqlUser.ExecuteReader();
                while (UserReader.Read())
                {
                    string s = UserReader[0].ToString();
                    comboBox7.Items.Add(s);
                }
                UserReader.Close();*/
                
                SqlDataReader TypeWorkReader = sqlTypeWork.ExecuteReader();
                while (TypeWorkReader.Read())
                {
                    string s = TypeWorkReader[0].ToString();
                    comboBox1.Items.Add(s);
                }
                TypeWorkReader.Close();

                SqlDataReader SectionReader = sqlSection.ExecuteReader();
                while (SectionReader.Read())
                {
                    string s = SectionReader[0].ToString();
                    comboBox2.Items.Add(s);
                }
                SectionReader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;   //или не нужно
            }
            finally
            {
                sqldbConnection.Close();
            }
        }

        public void UpdateDataIn(string myQuary, ComboBox myObj, string ColumnName)
        {
            SqlConnection sqldbConnection = new SqlConnection(conString);
            SqlCommand sqlCommand = new SqlCommand();
            sqlCommand.CommandText = myQuary;
            sqlCommand.Connection = sqldbConnection;
            try
            {
                sqldbConnection.Open();
                SqlDataReader DbReader = sqlCommand.ExecuteReader();
                while (DbReader.Read())
                {
                    string s = DbReader.GetString(DbReader.GetOrdinal(ColumnName));
                    myObj.Items.Add(s);
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            finally
            {
                sqldbConnection.Close();
            }
        }
        
        public void UpdateGrid(SqlCommand sqlCommand, DataGridView GridName)
        {
            SqlConnection conn = sqlCommand.Connection;
            SqlCommand myComand = sqlCommand;
            {
                try
                {
                    conn.Open();
                    SqlDataAdapter da = new SqlDataAdapter(myComand);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    GridName.DataSource = dt; //имя грида
                    conn.Close();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message);
                    return;   //или не нужно
                }
            }
        }
        /*public void UpdateGrid(string QUpdate)
        {
            SqlConnection conn = new SqlConnection(conString);
            SqlCommand sqlCommand = new SqlCommand();
            {
                try
                {
                    conn.Open();
                    string AllTasks = QUpdate;
                    SqlDataAdapter da = new SqlDataAdapter(AllTasks, conn);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt; //имя грида
                    conn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;   //или не нужно
                }
            }
        }*/

        /*//обновление грида
        private void button1_Click_1(object sender, EventArgs e)
        {
            DateTime dateStart = Convert.ToDateTime(dateTimePicker3.Text.ToString());
            DateTime dateEnd = Convert.ToDateTime(dateTimePicker4.Text.ToString());
            string NameProj = comboBox1.Text.ToString();

            SqlConnection conn = new SqlConnection(conString);
            SqlCommand TaskProj = new SqlCommand();

            TaskProj.Connection = conn;

            SqlParameter nameProj = new SqlParameter("@nameProj", NameProj);
            SqlParameter userID = new SqlParameter("@userID", UserID);
            SqlParameter DateStart = new SqlParameter("@DStart", dateStart);
            SqlParameter DateEnd = new SqlParameter("@DEnd", dateEnd);
                        
            if (comboBox1.Text.Length != 0)
            {
                if (checkBox4.Checked == true)
                {
                    if (checkBox5.Checked == true)//все фильтры
                    {

                        TaskProj.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Users.Surename, dbo.Zadan.Date_start, dbo.Zadan.Date_end"
                            + " FROM dbo.Zadan INNER JOIN"
                            + " dbo.Proekti ON dbo.Zadan.Id_project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users"
                        + " WHERE(dbo.Proekti.Name_Project = @nameProj) AND(dbo.Zadan.User_Give_out = @userID) AND(dbo.Zadan.Date_end >= @DStart) AND(dbo.Zadan.Date_start <= @DEnd)";

                        TaskProj.Parameters.Add(nameProj);
                        TaskProj.Parameters.Add(userID);
                        TaskProj.Parameters.Add(DateStart);
                        TaskProj.Parameters.Add(DateEnd);

                        UpdateGrid(TaskProj);
                    }
                    else//без окончания
                    {
                        TaskProj.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Users.Surename, dbo.Zadan.Date_start, dbo.Zadan.Date_end"
                             + " FROM dbo.Zadan INNER JOIN"
                             + " dbo.Proekti ON dbo.Zadan.Id_project = dbo.Proekti.Id_project INNER JOIN"
                             + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users"
                             + " WHERE(dbo.Proekti.Name_Project = @nameProj) AND(dbo.Zadan.User_Give_out = @userID) AND(dbo.Zadan.Date_start >= @DStart)";

                        TaskProj.Parameters.Add(nameProj);
                        TaskProj.Parameters.Add(userID);
                        TaskProj.Parameters.Add(DateStart);

                        UpdateGrid(TaskProj);
                    }
                }
                else if (checkBox4.Checked == true)//без старта
                {
                    TaskProj.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Users.Surename, dbo.Zadan.Date_start, dbo.Zadan.Date_end"
                            + " FROM dbo.Zadan INNER JOIN"
                            + " dbo.Proekti ON dbo.Zadan.Id_project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users"
                            + " WHERE(dbo.Proekti.Name_Project = @nameProj) AND(dbo.Zadan.User_Give_out = @userID) AND(dbo.Zadan.Date_end <= @DEnd)";

                    TaskProj.Parameters.Add(nameProj);
                    TaskProj.Parameters.Add(userID);
                    TaskProj.Parameters.Add(DateEnd);

                    UpdateGrid(TaskProj);
                }
                else//только проект
                {
                    TaskProj.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Users.Surename, dbo.Zadan.Date_start, dbo.Zadan.Date_end"
                            + " FROM dbo.Zadan INNER JOIN"
                            + " dbo.Proekti ON dbo.Zadan.Id_project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users"
                            + " WHERE(dbo.Proekti.Name_Project = @nameProj) AND(dbo.Zadan.User_Give_out = @userID)";

                    TaskProj.Parameters.Add(nameProj);
                    TaskProj.Parameters.Add(userID);

                    UpdateGrid(TaskProj);
                }
                
            }
            else if (checkBox4.Checked == true)
            {
                if (checkBox5.Checked == true)//время без проекта
                {
                    TaskProj.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Users.Surename, dbo.Zadan.Date_start, dbo.Zadan.Date_end"
                            + " FROM dbo.Zadan INNER JOIN"
                            + " dbo.Proekti ON dbo.Zadan.Id_project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users"
                            + " WHERE(dbo.Zadan.Date_end <= @DEnd) AND(dbo.Zadan.Date_start >= @DStart)";
                    
                    TaskProj.Parameters.Add(DateStart);
                    TaskProj.Parameters.Add(DateEnd);

                    UpdateGrid(TaskProj);
                }
                else//только старт
                {
                    TaskProj.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Users.Surename, dbo.Zadan.Date_start, dbo.Zadan.Date_end"
                            + " FROM dbo.Zadan INNER JOIN"
                            + " dbo.Proekti ON dbo.Zadan.Id_project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users"
                            + " WHERE(dbo.Zadan.Date_start >= @DStart)";

                    TaskProj.Parameters.Add(DateStart);

                    UpdateGrid(TaskProj);
                }
            }
            else if (checkBox5.Checked == true)//только конец
            {
                TaskProj.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Users.Surename, dbo.Zadan.Date_start, dbo.Zadan.Date_end"
                            + " FROM dbo.Zadan INNER JOIN"
                            + " dbo.Proekti ON dbo.Zadan.Id_project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users"
                            + " WHERE(dbo.Zadan.Date_end <= @DEnd)";
                
                TaskProj.Parameters.Add(DateEnd);

                UpdateGrid(TaskProj);
            }
            else//ничего не выбранно
            {
                MessageBox.Show("Параметр(ы) для фильтрации не задан(ы)","Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        */
        //все задания
        /*private void button2_Click(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = -1;
            checkBox4.Checked = false;
            checkBox5.Checked = false;

            SqlConnection conn = new SqlConnection(conString);
            SqlCommand AllTasks = new SqlCommand();
            AllTasks.Connection = conn;
            AllTasks.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Users.Surename, dbo.Zadan.Date_start, dbo.Zadan.Date_end"
            + " FROM((dbo.Zadan INNER JOIN"
            + " dbo.Proekti ON dbo.Zadan.Id_project = dbo.Proekti.Id_project) INNER JOIN"
            + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users)"
            + " WHERE(dbo.Zadan.User_Give_out = @userID)";
            
            SqlParameter userID = new SqlParameter("@userID", UserID);
            AllTasks.Parameters.Add(userID);
            
            UpdateGrid(AllTasks);
        }*/
        //Запрет ввода опр символов
        /*private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && (e.KeyChar != 8)&&(e.KeyChar != 46)) e.Handled = true;
            else
            {
                if (e.KeyChar == 46)
                {
                    if (textBox1.Text.Length > 0)
                    {
                        if (textBox1.Text.IndexOf('.') > -1) e.Handled = true;
                    }
                    else e.Handled = true;
                }
            }
        }*/

        /*private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                dateTimePicker1.Enabled = true;
                DateTime dtime = DateTime.Today;
                //string myDate = dtime.ToString("d");
                dateTimePicker1.Value = dtime;
            }
            else
            {
                dateTimePicker1.Enabled = false;
            }
        }*/

        /*private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                dateTimePicker2.Enabled = true;
                DateTime dtime = DateTime.Today;
                //string myDate = dtime.ToString("d");
                dateTimePicker2.Value = dtime;
            }
            else
            {
                dateTimePicker2.Enabled = false;
            }
        }*/

        /*private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == false)
            {
                listBox1.Enabled = true;
                listBox2.Enabled = true;
                listBox1.Items.Clear();
                listBox2.Items.Clear();
                string TaskToProj = "SELECT dbo.Proekti.Name_Project" +
                " FROM dbo.Proekti INNER JOIN(dbo.Users INNER JOIN dbo.Zadan ON dbo.Users.Id_users = dbo.Zadan.User_Give_out) ON dbo.Proekti.Id_project = dbo.Zadan.Id_project" +
                " WHERE((Zadan.User_Give_out) = " + UId + ")"
                + " GROUP BY dbo.Proekti.Name_Project";
                SqlConnection DbConn = new SqlConnection(conString);
                SqlCommand DbComand = new SqlCommand();
                DbComand.CommandText = TaskToProj;
                DbComand.Connection = DbConn;
                try
                {
                    DbConn.Open();
                    SqlDataReader DbReader = DbComand.ExecuteReader();
                    while (DbReader.Read())
                    {
                        string s = DbReader.GetString(DbReader.GetOrdinal("Name_Project"));
                        listBox1.Items.Add(s);
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
                finally
                {
                    DbConn.Close();
                }
            }
            else
            {
                listBox1.Enabled = false;
                listBox2.Enabled = false;
            }
        }*/

        /*private void button4_Click(object sender, EventArgs e)
        {
            string FStart, FEnd, CText = "";

            FStart = dateTimePicker1.Value.ToString("MM/dd/yyyy");
            FEnd = dateTimePicker2.Value.ToString("MM/dd/yyyy");

            if (checkBox3.Checked == true)
            {
                if (checkBox1.Checked == true)
                {
                    if (checkBox2.Checked == true)
                    {
                        //Все задания по всем проектам учитывая оба фильтра по времени
                        CText = "SELECT        Users.Surename, Users.First_Name, Users.Second_name, Zadan.Name_task, Zadan.Task_text, Zadan.Date_start, Zadan.Date_end, Vremay.Comment, Vremay.WorkTime, Vremay.date_entered, Otdeli.department, Doljn.Name_Position"
                        + " FROM(((((Users INNER JOIN"
                        + " Zadan ON Users.Id_users = Zadan.User_Give_out) INNER JOIN "
                        + " Proekti ON Users.Id_users = Proekti.GIP AND Zadan.Id_project = Proekti.Id_project) INNER JOIN"
                        + " Vremay ON Zadan.Id_task = Vremay.Id_task AND Zadan.Id_task = Vremay.Id_user AND Proekti.Id_project = Vremay.Id_proj) INNER JOIN"
                        + " Otdeli ON Users.department = Otdeli.Id_department) INNER JOIN"
                        + " Doljn ON Users.Id_position = Doljn.Id_Position)"
                        + " WHERE (Zadan.Date_start > #" + FStart + "#) AND(Zadan.Date_end < #" + FEnd + "#)";
                    }
                    //Фильтр только по времени старта
                    else CText = "SELECT        Users.Surename, Users.First_Name, Users.Second_name, Zadan.Name_task, Zadan.Task_text, Zadan.Date_start, Zadan.Date_end, Vremay.Comment, Vremay.WorkTime, Vremay.date_entered, Otdeli.department, Doljn.Name_Position"
                        + " FROM(((((Users INNER JOIN"
                        + " Zadan ON Users.Id_users = Zadan.User_Give_out) INNER JOIN "
                        + " Proekti ON Users.Id_users = Proekti.GIP AND Zadan.Id_project = Proekti.Id_project) INNER JOIN"
                        + " Vremay ON Zadan.Id_task = Vremay.Id_task AND Zadan.Id_task = Vremay.Id_user AND Proekti.Id_project = Vremay.Id_proj) INNER JOIN"
                        + " Otdeli ON Users.department = Otdeli.Id_department) INNER JOIN"
                        + " Doljn ON Users.Id_position = Doljn.Id_Position)"
                        + " WHERE (Zadan.Date_start > #" + FStart + "#)";
                }
                else if (checkBox2.Checked == true)
                {
                    FEnd = dateTimePicker2.Value.ToString();
                    //Фильтр только по времени окончания
                    CText = "SELECT        Users.Surename, Users.First_Name, Users.Second_name, Zadan.Name_task, Zadan.Task_text, Zadan.Date_start, Zadan.Date_end, Vremay.Comment, Vremay.WorkTime, Vremay.date_entered, Otdeli.department, Doljn.Name_Position"
                        + " FROM(((((Users INNER JOIN"
                        + " Zadan ON Users.Id_users = Zadan.User_Give_out) INNER JOIN "
                        + " Proekti ON Users.Id_users = Proekti.GIP AND Zadan.Id_project = Proekti.Id_project) INNER JOIN"
                        + " Vremay ON Zadan.Id_task = Vremay.Id_task AND Zadan.Id_task = Vremay.Id_user AND Proekti.Id_project = Vremay.Id_proj) INNER JOIN"
                        + " Otdeli ON Users.department = Otdeli.Id_department) INNER JOIN"
                        + " Doljn ON Users.Id_position = Doljn.Id_Position)"
                        + " WHERE (Zadan.Date_end < #" + FEnd + "#)";
                }
                else
                {
                    //Все задания по всем проектам, без учета фильтров по времени
                    CText = "SELECT Users.Surename, Users.First_Name, Users.Second_name, Zadan.Name_task, Zadan.Task_text, Zadan.Date_start, Zadan.Date_end, Vremay.Comment, Vremay.WorkTime, Vremay.date_entered, Otdeli.department, Doljn.Name_Position"
                        + " FROM(((((Users INNER JOIN"
                        + " Zadan ON Users.Id_users = Zadan.User_Give_out) INNER JOIN "
                        + " Proekti ON Users.Id_users = Proekti.GIP AND Zadan.Id_project = Proekti.Id_project) INNER JOIN"
                        + " Vremay ON Zadan.Id_task = Vremay.Id_task AND Zadan.Id_task = Vremay.Id_user AND Proekti.Id_project = Vremay.Id_proj) INNER JOIN"
                        + " Otdeli ON Users.department = Otdeli.Id_department) INNER JOIN"
                        + " Doljn ON Users.Id_position = Doljn.Id_Position)"
                        + " WHERE (Zadan.User_Give_out = "+ UId + ")";
                }
                    

                //вызов метода по импорту
                CreateBook(PuthToSave(), CText);
                MessageBox.Show("Экспорт выполнен", "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            //Если выборка по 1 или нескольким проектам
            //Делать проверку на коллекцию в listbox2
            else if (listBox2.Items.Count > 0)
            {
                CText = "SELECT Users.Surename, Users.First_Name, Users.Second_name, Proekti.Name_Project, Zadan.Name_task, Zadan.Task_text, Vremay.WorkTime, Vremay.date_entered"
                        +" FROM(((Users INNER JOIN Proekti ON Users.Id_users = Proekti.GIP) INNER JOIN Zadan ON Users.Id_users = Zadan.User_Give_out AND Proekti.Id_project = Zadan.Id_project) INNER JOIN"
                        +" Vremay ON Proekti.Id_project = Vremay.Id_proj AND Zadan.Id_task = Vremay.Id_task AND Zadan.Id_task = Vremay.Id_user)"
                + " WHERE((";
                //Если выбран всего 1 эллемент
                if (listBox2.Items.Count == 1)
                {
                    CText = "SELECT Proekti.Name_Project, Zadan.Name_task, Zadan.Task_text, Users.Surename, Users.First_Name, Users.department, Vremay.WorkTime, Vremay.date_entered"
                    + " FROM Users INNER JOIN(Vremay LEFT JOIN((Zadan LEFT JOIN Users ON Zadan.User_Give_out = Users.Id_users) LEFT JOIN Proekti ON Zadan.Id_project = Proekti.Id_project) ON Vremay.Id_task = Zadan.Id_task) ON Users.Id_users = Vremay.Id_user"
                    + " WHERE(((Proekti.Name_Project) = \"" + listBox2.Items[0].ToString() + "\") AND ((Zadan.User_Give_out) = " + UId + "))";
                    //Если есть фильтры по времени
                    //Проверка по дате старта
                    if (checkBox1.Checked == true)
                    {
                        //Проверка комбинации фильтров начала/окончания
                        if (checkBox2.Checked == true)
                        {
                            CText = CText + " AND((Zadan.Date_start > #" + FStart + "#) AND((Zadan.Date_end < #" + FEnd + "#)))";
                        }
                        else
                        {
                            CText = CText + " AND(Zadan.Date_start > #" + FStart + "#)";
                        }
                    }
                    //Проверка по дате окончания
                    else if (checkBox2.Checked == true)
                    {
                        CText = CText + " AND(Zadan.Date_end < #" + FEnd + "#)";
                    }
}
                else
                {
                    int k = listBox2.Items.Count;
                    //К запросу прибавляем выбранные проекты
                    foreach (object item in listBox2.Items)
                    {
                        k--;
                        if (k==0)
                        {
                            CText = CText + "(Proekti.Name_Project) = N'" + item.ToString() + "'";
                        }
                        else CText = CText + "(Proekti.Name_Project) = N'" + item.ToString() + "' OR ";
                    }
                    //Прибавляем фильтр по пользователю
                    CText = CText + ") AND ((Zadan.User_Give_out) = " + UId + "))";
                    //Проверка фильтров на время
                    if (checkBox1.Checked == true)
                    {
                        //Проверка комбинации фильтров начала/окончания
                        if (checkBox2.Checked == true)
                        {
                            CText = CText + " AND((Zadan.Date_start > #" + FStart + "#) AND((Zadan.Date_end < #" + FEnd + "#)))";
                        }
                        else
                        {
                            CText = CText + " AND(Zadan.Date_start > #" + FStart + "#)";
                        }
                    }
                    //Проверка по дате окончания
                    else if (checkBox2.Checked == true)
                    {
                        CText = CText + " AND(Zadan.Date_end < #" + FEnd + "#)";
                    }
                }

                //вызов метода по импорту
                CreateBook(PuthToSave(), CText);
                MessageBox.Show("Экспорт выполнен", "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            //Если не выбран ни один проект
            else { MessageBox.Show("Не выбранны проекты для экспорта. Выберите и повторите попытку.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Error); }

        }
        */
        //Сохранение экспортированных данных по выбранному пути
        private string PuthToSave()
        {
            string myPath = "";
            
            if(saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                myPath = saveFileDialog1.FileName;
            }

            return myPath;
        }
        private void CreateBook(string PuthToSave, string myQuary)
        {
            Microsoft.Office.Interop.Excel.Application myApp = new Microsoft.Office.Interop.Excel.Application();
            myApp.SheetsInNewWorkbook = 1;
            Microsoft.Office.Interop.Excel.Workbook NBook = myApp.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet worksheet = NBook.ActiveSheet;
            
            NBook.SaveAs(PuthToSave, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault);
            NBook.Close(false);
        }

        //Выбор проектов для экспорта в Эксель(удаление)
        /*private void button5_Click_1(object sender, EventArgs e)
        {
            while (listBox1.SelectedItems.Count > 0)
            {
                string item = (string)listBox1.SelectedItems[0];
                listBox2.Items.Add(item);
                listBox1.Items.Remove(item);
            }
        }*/

        //Выбор проектов для экспорта в Эксель(добавление)
        /*private void button6_Click(object sender, EventArgs e)
        {
            while (listBox2.SelectedItems.Count > 0)
            {
                string item = (string)listBox2.SelectedItems[0];
                listBox1.Items.Add(item);
                listBox2.Items.Remove(item);
            }
        }*/
       
        //Добавление записи в БД
        private void InsertIntoDB(string Quary)
        {
            SqlConnection DbConn = new SqlConnection(conString);
            SqlCommand DbComand = new SqlCommand();
            DbComand.CommandText = Quary;
            DbComand.Connection = DbConn;
            try
            {
                DbConn.Open();
                //OleDbDataReader DbReader = DbComand.ExecuteReader();
                DbComand.ExecuteNonQuery();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            finally
            {
                DbConn.Close();
            }
        }
        //Получение ИД элемента
        private string SomeID(ComboBox comboBox, string myQuary)
        {
            string retId = "";

            if (comboBox.Text.Length != 0)
            {
                string FSomeId = comboBox.Text.ToString();
                SqlConnection DbConn = new SqlConnection(conString);
                SqlCommand DbSomeID = new SqlCommand();
                DbSomeID.CommandText = myQuary;
                DbSomeID.Connection = DbConn;
                try
                {
                    DbConn.Open();
                    SqlDataReader DbReaderProj = DbSomeID.ExecuteReader();

                    while (DbReaderProj.Read())
                    {
                        retId = DbReaderProj[0].ToString();
                    }
                    DbReaderProj.Close();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
                finally
                {
                    DbConn.Close();
                }
            }

            return retId;
        }
        

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked == true) dateTimePicker5.Enabled = true;
            else dateTimePicker5.Enabled = false;
        }

        private void comboBox6_DropDownClosed(object sender, EventArgs e)
        {
            if (comboBox6.Text.Length != 0)
            {
                string SelectProjName = comboBox6.Text.ToString();
                SqlConnection Conn = new SqlConnection(conString);
                SqlCommand ProjName = new SqlCommand();
                SqlCommand sqlTypeBild = new SqlCommand();
                SqlParameter selProjName = new SqlParameter("@SelProjName", SelectProjName);
                SqlParameter uID = new SqlParameter("@UId", UId);


                ProjName.CommandText = "SELECT Id_project"
                    + " FROM dbo.Proekti"
                + " WHERE(Name_Project = @SelProjName)AND(InVork = 1)";
                ProjName.Parameters.Add(selProjName);
                ProjName.Connection = Conn;

                sqlTypeBild.CommandText = "SELECT dbo.TypeBild.NameBild, dbo.ProjBild.NumOrText"
                + " FROM dbo.ProjBild INNER JOIN"
                             + " dbo.TypeBild ON dbo.ProjBild.Id_TypeBilding = dbo.TypeBild.id_TypeBild"
                + " WHERE(dbo.ProjBild.Id_proj = @PId)";
                sqlTypeBild.Connection = Conn;
                try
                {
                    Conn.Open();
                    SqlDataReader IdProjReader = ProjName.ExecuteReader();
                    while (IdProjReader.Read())
                    {
                        PId = IdProjReader[0].ToString();
                    }
                    IdProjReader.Close();

                    SqlParameter pID = new SqlParameter("@PId", PId);
                    sqlTypeBild.Parameters.Add(pID);
                    SqlDataReader TypeBildProjReader = sqlTypeBild.ExecuteReader();

                    if (TypeBildProjReader.HasRows)
                    {
                        comboBox7.Items.Clear();
                        string s = "";
                        while (TypeBildProjReader.Read())
                        {
                            if (TypeBildProjReader[1].ToString().Length != 0) s = TypeBildProjReader[0].ToString() + " " + TypeBildProjReader[1].ToString();
                            else s = TypeBildProjReader[0].ToString();
                            comboBox7.Items.Add(s);
                        }
                        TypeBildProjReader.Close();
                    }
                }
                catch (SqlException ex) { MessageBox.Show("Произошла ошибка выполнения сценария: " + ex, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Stop); }
                finally { Conn.Close(); }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            comboBox6.SelectedIndex = -1;
            comboBox7.SelectedIndex = -1;

            //treeView2.SelectedNode = treeView2.Nodes[0];
            
            treeView2.CollapseAll();
            
            checkBox6.Checked = false;
            checkBox7.Checked = false;

            SqlConnection conn = new SqlConnection(conString);
            SqlCommand EnteredVal = new SqlCommand();
            SqlParameter userID2 = new SqlParameter("@UId", UserID);

            EnteredVal.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet, "
                         + " dbo.Report.TimeWork, dbo.Report.DateEntered"
            + " FROM            dbo.Report INNER JOIN"
                         + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                         + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
            + " WHERE(dbo.Users.Id_users = @UId)";

            DataGridView myObj = dataGridView2;

            EnteredVal.Parameters.Add(userID2);
            EnteredVal.Connection = conn;

            UpdateGrid(EnteredVal, myObj);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(conString);
            SqlCommand EnteredFilter = new SqlCommand();
            EnteredFilter.Connection = conn;

            DateTime DStart = Convert.ToDateTime(dateTimePicker5.Value.ToString());
            DateTime DEnd = Convert.ToDateTime(dateTimePicker6.Value.ToString());

            TreeNode treeNode = treeView2.SelectedNode;
            int level = treeNode.Level;

            string NameProj = "";
            string TipeBild = "";

            if (level == 1)
            {
                NameProj = treeNode.Text;
                TipeBild = "";
            }
            else if (level > 1)
            {
                var ArrVal = treeNode.FullPath.Split('\\');
                NameProj = ArrVal[1];
                TipeBild = treeView2.SelectedNode.Text;
            }

            SqlParameter uID = new SqlParameter("@UId", UId);
            SqlParameter dStart = new SqlParameter("@DSTart", DStart);
            SqlParameter dEnd = new SqlParameter("@DEnd", DEnd);
            SqlParameter nameProj = new SqlParameter("@NameProj", NameProj);
            SqlParameter tipeBild = new SqlParameter("@TypeBild", TipeBild);

            DataGridView myGrid = dataGridView2;

            if (NameProj.Length != 0)
            {
                if (TipeBild.Length != 0)
                {
                    if (checkBox6.Checked == true)
                    {
                        if (checkBox7.Checked == true)//1234
                        {
                            EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet,"
                                +" dbo.Report.TimeWork, dbo.Report.DateEntered"
                            +" FROM            dbo.Report INNER JOIN"
                                +" dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                +" dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Proekti.Name_Project = @NameProj) AND(dbo.Report.TypeBild = @TypeBild) AND(dbo.Report.DateEntered >= @DSTart) AND(dbo.Report.DateEntered <= @DEnd)"
                            +" ORDER BY dbo.Report.DateEntered";

                            EnteredFilter.Parameters.Add(uID);
                            EnteredFilter.Parameters.Add(dStart);
                            EnteredFilter.Parameters.Add(dEnd);
                            EnteredFilter.Parameters.Add(nameProj);
                            EnteredFilter.Parameters.Add(tipeBild);

                            UpdateGrid(EnteredFilter, myGrid);
                        }
                        else//123
                        {
                            EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet,"
                                + " dbo.Report.TimeWork, dbo.Report.DateEntered"
                            + " FROM            dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Proekti.Name_Project = @NameProj) AND(dbo.Report.TypeBild = @TypeBild) AND(dbo.Report.DateEntered >= @DSTart)"
                            + " ORDER BY dbo.Report.DateEntered";

                            EnteredFilter.Parameters.Add(uID);
                            EnteredFilter.Parameters.Add(dStart);
                            EnteredFilter.Parameters.Add(nameProj);
                            EnteredFilter.Parameters.Add(tipeBild);

                            UpdateGrid(EnteredFilter, myGrid);
                        }
                    }
                    else if (checkBox7.Checked == true)//124
                    {
                        EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet,"
                                + " dbo.Report.TimeWork, dbo.Report.DateEntered"
                            + " FROM            dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Proekti.Name_Project = @NameProj) AND(dbo.Report.TypeBild = @TypeBild) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " ORDER BY dbo.Report.DateEntered";

                        EnteredFilter.Parameters.Add(uID);
                        EnteredFilter.Parameters.Add(dEnd);
                        EnteredFilter.Parameters.Add(nameProj);
                        EnteredFilter.Parameters.Add(tipeBild);

                        UpdateGrid(EnteredFilter, myGrid);
                    }
                    else//12
                    {
                        EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet,"
                                + " dbo.Report.TimeWork, dbo.Report.DateEntered"
                            + " FROM            dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Proekti.Name_Project = @NameProj) AND(dbo.Report.TypeBild = @TypeBild)"
                            + " ORDER BY dbo.Report.DateEntered";

                        EnteredFilter.Parameters.Add(uID);
                        EnteredFilter.Parameters.Add(nameProj);
                        EnteredFilter.Parameters.Add(tipeBild);

                        UpdateGrid(EnteredFilter, myGrid);
                    }
                }
                else if (checkBox6.Checked == true)
                {
                    if (checkBox7.Checked == true)//134
                    {
                        EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet,"
                                + " dbo.Report.TimeWork, dbo.Report.DateEntered"
                            + " FROM            dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Proekti.Name_Project = @NameProj) AND(dbo.Report.DateEntered >= @DSTart) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " ORDER BY dbo.Report.DateEntered";

                        EnteredFilter.Parameters.Add(uID);
                        EnteredFilter.Parameters.Add(dStart);
                        EnteredFilter.Parameters.Add(dEnd);
                        EnteredFilter.Parameters.Add(nameProj);

                        UpdateGrid(EnteredFilter, myGrid);
                    }
                    else//13
                    {
                        EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet,"
                                + " dbo.Report.TimeWork, dbo.Report.DateEntered"
                            + " FROM            dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Proekti.Name_Project = @NameProj)  AND(dbo.Report.DateEntered >= @DSTart)"
                            + " ORDER BY dbo.Report.DateEntered";

                        EnteredFilter.Parameters.Add(uID);
                        EnteredFilter.Parameters.Add(dStart);
                        EnteredFilter.Parameters.Add(nameProj);

                        UpdateGrid(EnteredFilter, myGrid);
                    }
                }
                else if (checkBox7.Checked == true)//14
                {
                    EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet,"
                                + " dbo.Report.TimeWork, dbo.Report.DateEntered"
                            + " FROM            dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Proekti.Name_Project = @NameProj) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " ORDER BY dbo.Report.DateEntered";

                    EnteredFilter.Parameters.Add(uID);
                    EnteredFilter.Parameters.Add(dEnd);
                    EnteredFilter.Parameters.Add(nameProj);

                    UpdateGrid(EnteredFilter, myGrid);
                }
                else//1
                {
                    EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet,"
                                + " dbo.Report.TimeWork, dbo.Report.DateEntered"
                            + " FROM            dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Proekti.Name_Project = @NameProj)"
                            + " ORDER BY dbo.Report.DateEntered";

                    EnteredFilter.Parameters.Add(uID);
                    EnteredFilter.Parameters.Add(nameProj);

                    UpdateGrid(EnteredFilter, myGrid);
                }
            }
            else if (TipeBild.Length != 0)
            {
                if (checkBox6.Checked == true)
                {
                    if (checkBox7.Checked == true)//234
                    {
                        EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet,"
                                + " dbo.Report.TimeWork, dbo.Report.DateEntered"
                            + " FROM            dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.TypeBild = @TypeBild) AND(dbo.Report.DateEntered >= @DSTart) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " ORDER BY dbo.Report.DateEntered";

                        EnteredFilter.Parameters.Add(uID);
                        EnteredFilter.Parameters.Add(dStart);
                        EnteredFilter.Parameters.Add(dEnd);
                        EnteredFilter.Parameters.Add(tipeBild);

                        UpdateGrid(EnteredFilter, myGrid);
                    }
                    else//23
                    {
                        EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet,"
                                + " dbo.Report.TimeWork, dbo.Report.DateEntered"
                            + " FROM            dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
                            + " WHERE(dbo.Users.Id_users = @UId)  AND(dbo.Report.TypeBild = @TypeBild) AND(dbo.Report.DateEntered >= @DSTart))"
                            + " ORDER BY dbo.Report.DateEntered";

                        EnteredFilter.Parameters.Add(uID);
                        EnteredFilter.Parameters.Add(dStart);
                        EnteredFilter.Parameters.Add(tipeBild);

                        UpdateGrid(EnteredFilter, myGrid);
                    }
                }
                else if (checkBox7.Checked == true)//24
                {
                    EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet,"
                                + " dbo.Report.TimeWork, dbo.Report.DateEntered"
                            + " FROM            dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.TypeBild = @TypeBild) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " ORDER BY dbo.Report.DateEntered";

                    EnteredFilter.Parameters.Add(uID);
                    EnteredFilter.Parameters.Add(dEnd);
                    EnteredFilter.Parameters.Add(tipeBild);

                    UpdateGrid(EnteredFilter, myGrid);
                }
                else//2
                {
                    EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet,"
                                + " dbo.Report.TimeWork, dbo.Report.DateEntered"
                            + " FROM            dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.TypeBild = @TypeBild)"
                            + " ORDER BY dbo.Report.DateEntered";

                    EnteredFilter.Parameters.Add(uID);
                    EnteredFilter.Parameters.Add(tipeBild);

                    UpdateGrid(EnteredFilter, myGrid);
                }
            }
            else if(checkBox6.Checked == true)
            {
                if (checkBox7.Checked == true)//34
                {
                    EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet,"
                                + " dbo.Report.TimeWork, dbo.Report.DateEntered"
                            + " FROM            dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.DateEntered >= @DSTart) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " ORDER BY dbo.Report.DateEntered";

                    EnteredFilter.Parameters.Add(uID);
                    EnteredFilter.Parameters.Add(dStart);
                    EnteredFilter.Parameters.Add(dEnd);

                    UpdateGrid(EnteredFilter, myGrid);
                }
                else//3
                {
                    EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet,"
                                + " dbo.Report.TimeWork, dbo.Report.DateEntered"
                            + " FROM            dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.DateEntered >= @DSTart)"
                            + " ORDER BY dbo.Report.DateEntered";

                    EnteredFilter.Parameters.Add(uID);
                    EnteredFilter.Parameters.Add(dStart);

                    UpdateGrid(EnteredFilter, myGrid);
                }
            }
            else if (checkBox7.Checked == true)//4
            {
                EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.Comment, dbo.Report.CountSheet,"
                                + " dbo.Report.TimeWork, dbo.Report.DateEntered"
                            + " FROM            dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " ORDER BY dbo.Report.DateEntered";

                EnteredFilter.Parameters.Add(uID);
                EnteredFilter.Parameters.Add(dEnd);

                UpdateGrid(EnteredFilter, myGrid);
            }
            else
            {
                MessageBox.Show("Параметр(ы) для фильтрации не задан(ы)", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string path = "";
            DateTime now = DateTime.Now;
            string NameFile = now.ToString("d");
            
            saveFileDialog1.FileName = NameFile;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                path = saveFileDialog1.FileName;
                if (!Directory.Exists(Path.GetDirectoryName(path))) Directory.CreateDirectory(Path.GetDirectoryName(path));
                //if (!File.Exists(path)) using (File.Create(path)) { };
                Exception result = ExportGrid(path, NameFile);
                if (result != null)
                {
                    MessageBox.Show("Процес экспорта выполнен с ошибкой:" + result, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("Процес экспорта выполнен:" + result, "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("Процес экспорта отменен.", "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private Exception ExportGrid (string Path, string NameFile)
        {
            Exception flag = null;

            Microsoft.Office.Interop.Excel.Application myApp = new Microsoft.Office.Interop.Excel.Application();
            myApp.SheetsInNewWorkbook = 1;
            Microsoft.Office.Interop.Excel.Workbook NBook = myApp.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet worksheet = NBook.ActiveSheet;
            

            try
            {   //Данные из грида в DataTable
                DataTable dataTable = (DataTable)(dataGridView2.DataSource);

                int mRow = dataTable.Rows.Count;
                int mColumn = dataTable.Columns.Count;
                //Переменная для перегона из DataTable
                object[,] dataExport = new object[mRow, mColumn];

                for (int i = 0; i < mRow; i++)
                {
                    for (int j = 0; j < mColumn; j++)
                    {
                        dataExport[i, j] = dataTable.Rows[i][j];
                    }
                }

                Microsoft.Office.Interop.Excel.Range range;
                Microsoft.Office.Interop.Excel.Range range1;
                Microsoft.Office.Interop.Excel.Range range2;
                //Заполнение диапазона в эксель из массива данных
                range1 = worksheet.Cells[1, 1] as Microsoft.Office.Interop.Excel.Range;
                range2 = worksheet.Cells[mRow, mColumn] as Microsoft.Office.Interop.Excel.Range;
                range = worksheet.get_Range(range1, range2);
                range.Value2 = dataExport;

                //Для столбца с датой устанавливаем необходимый формат
                range1 = worksheet.Cells[1, 6] as Microsoft.Office.Interop.Excel.Range;
                range2 = worksheet.Cells[mRow, 6] as Microsoft.Office.Interop.Excel.Range;

                range = worksheet.get_Range(range1, range2);
                range.EntireColumn.NumberFormat = "MM/DD/YYYY";

                //Авторазмер столбцов
                worksheet.Columns.EntireColumn.AutoFit();

                //Сохранение и закрытие книги
                NBook.SaveAs(Path, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault);
                NBook.Close(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Процес экспорта отменен.", "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                flag = ex;
            }
            //Флаг метода (успешно/нет)
            return flag;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            TimerVal += 1;
            if (TimerVal == 120)
            {
                MessageBox.Show("Вы не использовали \"Журнал учета рабочего времени\" длительное время. Приложение закроется автоматически, через 5 минут, если не зафиксирует активность! ", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            if (TimerVal == 180)
            {
                Application.Exit();
            }
        }

        private void Form2_Activated(object sender, EventArgs e)
        {
            timer1.Stop();
            TimerVal = 0;
        }

        private void Form2_Deactivate(object sender, EventArgs e)
        {
            timer1.Start();
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked == true) dateTimePicker6.Enabled = true;
            else dateTimePicker6.Enabled = false;
        }

        //Получение ИД элемента
        private string SomeID(string myQuary)
        {
            string retId = "";

                SqlConnection DbConn = new SqlConnection(conString);
                SqlCommand DbSomeID = new SqlCommand();
                DbSomeID.CommandText = myQuary;
                DbSomeID.Connection = DbConn;
                try
                {
                    DbConn.Open();
                    SqlDataReader DbReaderProj = DbSomeID.ExecuteReader();

                    while (DbReaderProj.Read())
                    {
                        retId = DbReaderProj[0].ToString();
                    }
                    DbReaderProj.Close();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
                finally
                {
                    DbConn.Close();
                }

            return retId;
        }
        
        //Проверка суммы часов
        public float ChekSumWorkTimeOneDay(DateTime date)
        {
            float SumOneDay = 0;
            DateTime myDate = DateTime.Today;
            SqlConnection conn = new SqlConnection(conString);
            SqlCommand SumTime = new SqlCommand();
            SqlParameter dateEntered = new SqlParameter("@DateEntered", myDate);
            SqlParameter uId = new SqlParameter("@UId", UId);

            SumTime.CommandText = "SELECT WorkTime"
            + " FROM dbo.Vremay"
            + " WHERE(Id_user = @UId) AND(date_entered =@DateEntered)";
            SumTime.Connection = conn;
            SumTime.Parameters.Add(dateEntered);
            SumTime.Parameters.Add(uId);

            try
            {
                conn.Open();
                SqlDataReader sumTime = SumTime.ExecuteReader();
                float s;
                while (sumTime.Read())
                {
                    string a = sumTime[0].ToString();
                    if (a.Length > 0)
                    {
                        s = Convert.ToSingle(sumTime[0].ToString());
                        SumOneDay = SumOneDay + s;
                    }
                    else s = 0;
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return SumOneDay;
            }
            finally { conn.Close(); }

            return SumOneDay;
        }

        private void comboBox3_DropDownClosed(object sender, EventArgs e)
        {/*
            if ((comboBox2.Text.Length != 0) && (comboBox3.Text.Length != 0))
            {
                PId = comboBox2.Text;
                TId = comboBox3.Text;
                string NameTask = TId;

                //SELECT Proekti.Id_project FROM Proekti WHERE(((Proekti.Name_Project) = "Проект 1"));
                SqlConnection DbConn = new SqlConnection(conString);
                SqlCommand DbComandProj = new SqlCommand();
                SqlCommand DbComandTask = new SqlCommand();

                SqlParameter uID = new SqlParameter("@UId", UId);
                SqlParameter pID = new SqlParameter("@PId", PId);
                SqlParameter uID2 = new SqlParameter("@UId2", UId);
                SqlParameter pID2 = new SqlParameter("@PId2", PId);
                SqlParameter TName = new SqlParameter("@TName", NameTask);

                DbComandProj.CommandText = "SELECT dbo.Proekti.id_project FROM dbo.Proekti WHERE(((dbo.Proekti.Name_Project) = @PId))";
                DbComandProj.Parameters.Add(pID);

                DbComandTask.CommandText = "SELECT        dbo.Zadan.Id_task"
                                            +" FROM dbo.Proekti INNER JOIN"
                                            +" dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                            +" dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users"
                                            + " WHERE(dbo.Zadan.User_Give_out = @UId2) AND(dbo.Proekti.Name_Project = @PId2) AND(dbo.Zadan.Task_text = @TName)";
                DbComandTask.Parameters.Add(uID2);
                DbComandTask.Parameters.Add(pID2);
                DbComandTask.Parameters.Add(TName);

                DbComandProj.Connection = DbConn;
                DbComandTask.Connection = DbConn;

                try
                {
                    DbConn.Open();
                    SqlDataReader DbReaderProj = DbComandProj.ExecuteReader();

                    while (DbReaderProj.Read())
                    {
                        PId = DbReaderProj[0].ToString();
                    }
                    DbReaderProj.Close();

                    SqlDataReader DbReaderTask = DbComandTask.ExecuteReader();
                    while (DbReaderTask.Read())
                    {
                        TId = DbReaderTask[0].ToString();
                    }
                    DbReaderTask.Close();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
                finally
                {
                    DbConn.Close();
                }
            }
            */
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && (e.KeyChar != 8) && (e.KeyChar != 46)) e.Handled = true;
            else
            {
                if (e.KeyChar == 46)
                {
                    if (textBox4.Text.Length > 0)
                    {
                        if (textBox4.Text.IndexOf('.') > -1) e.Handled = true;
                    }
                    else e.Handled = true;
                }
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && (e.KeyChar != 8) && (e.KeyChar != 46)) e.Handled = true;
            else
            {
                if (e.KeyChar == 46)
                {
                    if (textBox1.Text.Length > 0)
                    {
                        if (textBox1.Text.IndexOf('.') > -1) e.Handled = true;
                    }
                    else e.Handled = true;
                }
            }
        }

        private void myTreeViewWork(TreeView treeView)
        {
            //TreeView treeView = treeView1;
            TreeNode ProjectsNode = new TreeNode();
            ProjectsNode.Name = "Projects";
            ProjectsNode.Text = "Проекты";
            treeView.Nodes.Add(ProjectsNode);

            SqlConnection Conn = new SqlConnection(conString);
            SqlCommand AllProj = new SqlCommand();
            AllProj.Connection = Conn;
            AllProj.CommandText = "SELECT Id_project, Name_Project"
                    + " FROM dbo.Proekti"
                + " WHERE(InVork = 1)";

            try
            {
                Conn.Open();
                SqlDataReader DRProj = AllProj.ExecuteReader();

                List<string> ProjectsName = new List<string>();
                int i = 0;
                while (DRProj.Read())
                {
                    string s = DRProj[1].ToString();
                    ProjectsName.Add(s);
                    ProjectsNode.Nodes.Add(s, s);
                    i++;
                }
                DRProj.Close();

                int j = 0;
                foreach (string s in ProjectsName)
                {
                    SqlParameter NameProj = new SqlParameter("@NameProj", s);
                    SqlCommand sqlTypeBild = new SqlCommand();
                    sqlTypeBild.Connection = Conn;
                    sqlTypeBild.CommandText = "SELECT dbo.TypeBild.NameBild, dbo.ProjBild.NumOrText, dbo.ProjBild.NumTree"
                    + " FROM dbo.Proekti INNER JOIN"
                         + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj INNER JOIN"
                         + " dbo.TypeBild ON dbo.ProjBild.Id_TypeBilding = dbo.TypeBild.id_TypeBild"
                    + " WHERE(dbo.Proekti.Name_Project = @NameProj)";
                    sqlTypeBild.Parameters.Add(NameProj);

                    //функция обработки полученных данных
                    AddTypeBild(sqlTypeBild, treeView, j, s);

                    j++;
                }
            }
            catch (SqlException ex) { MessageBox.Show("Произошла ошибка выполнения сценария: " + ex, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally { Conn.Close(); }
        }

        private void AddTypeBild(SqlCommand sqlCommand, TreeView treeView, int i, string NameProj)
        {
            SqlCommand myCommand = sqlCommand;
            List<string[]> mylist = new List<string[]>();
            try
            {
                //Получаем массив для обработки
                string[,] myData = new string[150, 2];

                SqlDataReader sqlDataReader = myCommand.ExecuteReader();
                int j = 0;
                while (sqlDataReader.Read())
                {
                    string s = sqlDataReader[0].ToString();

                    string pref = sqlDataReader[1].ToString();
                    if (pref.Length != 0) s = s + " " + pref;

                    string NumInTree = sqlDataReader[2].ToString();

                    myData[j, 0] = s;
                    myData[j, 1] = NumInTree;

                    var ArrNum = s.Split('.');
                    mylist.Add(NumInTree.Split('.'));
                    j++;//Число строк в необработанном массиве
                }

                sqlDataReader.Close();

                string[][] UnSortArr = mylist.ToArray();    //Лист в массив
                int RowCountList = mylist.Count;    //Кол-во строк всего
                int[,] GoodArrNum = CovertArr(UnSortArr, RowCountList); //Из массива массивов в двумерный массив
                GoodArrNum = SortArray(GoodArrNum, RowCountList);   //Отсортированный список номеров по возврастанию

                string[,] FinnArr = UpdateGoodArr(GoodArrNum, myData, RowCountList);
                AddNodesTree(FinnArr, RowCountList, i, treeView);
            }
            catch (Exception ex) { MessageBox.Show("" + ex); }
        }

        private void AddNodesTree(string[,] FinnArr, int RowCount, int NumberProj, TreeView treeView)
        {
            int NumNode1 = -1;
            int NumNode2 = -1;

            TreeNode rootNode = treeView.Nodes[0].Nodes[NumberProj];


            string OldNode1 = "";
            string OldNode2 = "";
            for (int i = 0; i < RowCount; i++)
            {
                var LiltleArr = FinnArr[i, 0].Split('.');
                if (LiltleArr.Length == 1)
                {
                    if (OldNode1 != LiltleArr[0])
                    {
                        NumNode1++;
                        NumNode2 = -1;
                        OldNode1 = LiltleArr[0];
                    }
                }
                if (LiltleArr.Length == 2)
                {
                    if (OldNode2 != LiltleArr[1])
                    {
                        NumNode2++;
                        OldNode2 = LiltleArr[1];
                    }
                }
                if (LiltleArr.Length == 1) rootNode.Nodes.Add(FinnArr[i, 1], FinnArr[i, 1]);
                else if (LiltleArr.Length == 2) rootNode.Nodes[NumNode1].Nodes.Add(FinnArr[i, 1], FinnArr[i, 1]);
                else if (LiltleArr.Length == 3) rootNode.Nodes[NumNode1].Nodes[NumNode2].Nodes.Add(FinnArr[i, 1], FinnArr[i, 1]);
            }
        }

        private string[,] UpdateGoodArr(int[,] GoodArrNum, string[,] myData, int RowCount)
        {
            string[,] rez = new string[RowCount, 2];
            for (int i = 0; i < RowCount; i++)
            {
                string Val = "";
                if (GoodArrNum[i, 1] == 0) Val = GoodArrNum[i, 0].ToString();
                else if (GoodArrNum[i, 2] == 0) Val = GoodArrNum[i, 0].ToString() + "." + GoodArrNum[i, 1].ToString();
                else Val = GoodArrNum[i, 0].ToString() + "." + GoodArrNum[i, 1].ToString() + "." + GoodArrNum[i, 2].ToString();

                for (int j = 0; j < RowCount; j++)
                {
                    if (Val == myData[j, 1])
                    {
                        rez[i, 0] = Val;
                        rez[i, 1] = myData[j, 0];
                        break;
                    }
                }
            }
            return rez;
        }

        private int[,] CovertArr(string[][] UnSortArr, int RowCount)
        {
            int[,] GoodArr = new int[RowCount, 3];
            int maxRow = UnSortArr.GetLength(0);
            try
            {
                for (int i = 0; i < maxRow; i++)
                {
                    int maxColumn = UnSortArr[i].Length;
                    for (int j = 0; j < maxColumn; j++)
                    {
                        if (UnSortArr[i][j].Length == 0)
                        {
                            GoodArr[i, j] = 0;
                        }
                        else GoodArr[i, j] = Convert.ToInt32(UnSortArr[i][j]);
                    }
                    if (maxColumn == 1)
                    {
                        GoodArr[i, 1] = 0;
                        GoodArr[i, 2] = 0;
                    }
                    else if (maxColumn == 2) GoodArr[i, 2] = 0;
                }
            }
            catch (Exception ex) { MessageBox.Show("" + ex); }
            return GoodArr;
        }

        private int[,] SortArray(int[,] myArr, int RowCount)
        {
            int[,] rez = new int[RowCount, 3];

            for (int i = 0; i < RowCount - 1; i++)
            {
                if (myArr[i, 0] != null)
                {
                    int min = i;
                    for (int j = i + 1; j < RowCount; j++)
                    {
                        if (myArr[min, 0] > myArr[j, 0])
                        {
                            min = j;
                        }
                    }
                    int temp0 = myArr[min, 0];
                    int temp1 = myArr[min, 1];
                    int temp2 = myArr[min, 2];

                    myArr[min, 0] = myArr[i, 0];
                    myArr[min, 1] = myArr[i, 1];
                    myArr[min, 2] = myArr[i, 2];

                    myArr[i, 0] = temp0;
                    myArr[i, 1] = temp1;
                    myArr[i, 2] = temp2;
                }
            }
            return rez = myVoid2(myArr, RowCount); ;
        }

        public int[,] myVoid2(int[,] myArr, int RowCount)
        {
            int[,] rez = new int[RowCount, 3];

            for (int i = 0; i < RowCount - 1; i++)
            {
                int min = i;
                int firstIndex = myArr[min, 0];
                int TherdIndex = myArr[min, 2];

                for (int j = i + 1; j < RowCount; j++)
                {
                    if (myArr[j, 0] == firstIndex)
                    {
                        if (myArr[j, 2] == TherdIndex)
                        {
                            if (myArr[min, 1] > myArr[j, 1])
                            {
                                min = j;
                            }
                        }
                    }
                }
                int temp0 = myArr[min, 0];
                int temp1 = myArr[min, 1];
                int temp2 = myArr[min, 2];

                myArr[min, 0] = myArr[i, 0];
                myArr[min, 1] = myArr[i, 1];
                myArr[min, 2] = myArr[i, 2];

                myArr[i, 0] = temp0;
                myArr[i, 1] = temp1;
                myArr[i, 2] = temp2;
            }
            myVoid3(myArr, RowCount);
            return rez = myVoid3(myArr, RowCount); ;
        }

        public int[,] myVoid3(int[,] myArr, int RowCount)
        {
            int[,] rez = new int[RowCount, 3];

            for (int i = 0; i < RowCount - 1; i++)
            {
                if (myArr[i, 0] != null)
                {
                    int min = i;
                    int firstIndex = myArr[min, 0];
                    int SecondIndex = myArr[min, 1];

                    for (int j = i + 1; j < RowCount; j++)
                    {
                        if (myArr[j, 0] == firstIndex)
                        {
                            if (myArr[j, 1] == SecondIndex)
                            {
                                if (myArr[min, 2] > myArr[j, 2])
                                {
                                    min = j;
                                }
                            }
                        }
                    }
                    int temp0 = myArr[min, 0];
                    int temp1 = myArr[min, 1];
                    int temp2 = myArr[min, 2];

                    myArr[min, 0] = myArr[i, 0];
                    myArr[min, 1] = myArr[i, 1];
                    myArr[min, 2] = myArr[i, 2];

                    myArr[i, 0] = temp0;
                    myArr[i, 1] = temp1;
                    myArr[i, 2] = temp2;
                }
            }
            return rez = myArr;
        }
    }
}
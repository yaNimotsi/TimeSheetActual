using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;

namespace TimeSheet
{
    public partial class Form3 : Form
    {
        public string UId, DepId, TId, PId, conString, ProjName1, ProjName2, TypeBild1, TypeBild2, fullPath1, fullPath2;
        private int UserID, level1, level2;
        SqlDataAdapter daGrid2;
        public Form3(string UId, string DepId, string connectionString)
        {
            this.UId = UId;
            this.DepId = DepId;
            UserID = Convert.ToInt32(UId);
            this.conString = connectionString;
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(conString);
            SqlCommand AllTasks = new SqlCommand();

            DataGridView GridTab1 = dataGridView1;

            AutoSizeGridColumn(GridTab1);
            
            AllTasks.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                + " dbo.Report.Comment, dbo.Report.CountSheet, dbo.Otdeli.department"
                + " FROM dbo.Report INNER JOIN"
                         + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                         + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                         + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                         + " WHERE(dbo.Users.department = @DepId)";

            SqlParameter depId = new SqlParameter("@DepId", Convert.ToInt32(DepId.ToString()));
            AllTasks.Parameters.Add(depId);
            
            UpdateGrid(AllTasks);
            GridHeaderName();
            DataToListBox();
            DateTime myDate = DateTime.Today;
            
            dateTimePicker3.Value = myDate;
            dateTimePicker4.Value = myDate;

            checkBox4.Checked = false;
            checkBox5.Checked = false;
            
            dateTimePicker3.Enabled = false;
            dateTimePicker4.Enabled = false;

            myTreeViewWork(treeView1);
            myTreeViewWork(treeView2);

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView1.AllowUserToResizeColumns = true;
            dataGridView1.RowHeadersVisible = false;
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
                    daGrid2 = da;

                    da.Fill(dt);

                    GridName.DataSource = dt; //имя грида
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message);
                    return;   //или не нужно
                }
                finally { conn.Close(); }
            }
        }

        public void UpdateGrid(SqlCommand sqlCommand)
        {
            SqlConnection conn = new SqlConnection(conString);
            SqlCommand myComand = sqlCommand;
            myComand.Connection = conn;
            {
                try
                {
                    conn.Open();
                    SqlDataAdapter da = new SqlDataAdapter(myComand);
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
                finally { conn.Close(); }
            }
        }
        public void UpdateGrid(string QUpdate)
        {

            SqlConnection conn = new SqlConnection(conString);
            try
                {
                    conn.Open();
                    //Фильтр по заданию SELECT Zadan.Name_task, Zadan.Task_text, Zadan.User_Give_out, Zadan.Date_start, Zadan.Date_end, Users.Surename, Users.First_Name, Users.Second_name, Proekti.Name_Project FROM Proekti INNER JOIN(Users INNER JOIN Zadan ON Users.Id_users = Zadan.User_Give_out) ON Proekti.Id_project = Zadan.Id_project WHERE(((Zadan.User_Give_out) = 4) AND((Zadan.Id_project) = 1));
                    string AllTasks = QUpdate;
                    SqlDataAdapter da = new SqlDataAdapter(AllTasks, conn);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt; //имя грида
                    conn.Close();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message);
                    return;   //или не нужно
                }
            finally { conn.Close(); }
        }

        private void GridHeaderName()
        {
            dataGridView1.Columns[0].HeaderText = "Название проекта";
            dataGridView1.Columns[1].HeaderText = "Название здания/сооружения";
            dataGridView1.Columns[2].HeaderText = "Тип работы";
            dataGridView1.Columns[3].HeaderText = "Раздел";
            dataGridView1.Columns[4].HeaderText = "ФИО исполнителя";
            dataGridView1.Columns[5].HeaderText = "Затраченое время";
            dataGridView1.Columns[6].HeaderText = "Дата заполнения отчетности";
            dataGridView1.Columns[7].HeaderText = "Коментарий";
            dataGridView1.Columns[8].HeaderText = "Кол-во выполненых листов";
            dataGridView1.Columns[9].HeaderText = "Отдел";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(conString);
            treeView2.PathSeparator = "/";
            TreeNode treeNode = treeView2.SelectedNode;

            string NameProj;
            int level = level2;
            string fullPath = fullPath2;
            
            if (ProjName2 == null) NameProj = "";
            else NameProj = ProjName2;

            string TipeBild = "";
            if (TypeBild2 == null) TipeBild = "";
            else TipeBild = TypeBild2;
            
            SqlCommand TaskToProj = new SqlCommand();

            TaskToProj.Connection = conn;

            int EntUserId = UserId(comboBox1.Text.ToString());
            DateTime dateStart = Convert.ToDateTime(dateTimePicker3.Text.ToString());
            DateTime dateEnd = Convert.ToDateTime(dateTimePicker4.Text.ToString());

            SqlParameter entUserId = new SqlParameter("@EntUserId", EntUserId);
            SqlParameter ProjName = new SqlParameter("@projName", NameProj);
            SqlParameter FullPath = new SqlParameter("@fullPath", fullPath);
            SqlParameter depId = new SqlParameter("@DepID", DepId);
            SqlParameter DateStart = new SqlParameter("@DStart", dateStart);
            SqlParameter DateEnd = new SqlParameter("@DEnd", dateEnd);
            SqlParameter DeepId = new SqlParameter("@DepId", DepId);
            
            if (NameProj.Length != 0)
            {
                if (TipeBild.Length != 0)
                {
                    if (comboBox1.Text.Length != 0)
                    {
                        if (checkBox4.Checked == true)
                        {
                            if (checkBox5.Checked == true)//12345
                            {
                                if (level > 1)
                                {
                                    TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                    + "dbo.Report.Comment, dbo.Report.CountSheet"
                                    + " FROM dbo.Report INNER JOIN"
                                    + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                    + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                    + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                                + " WHERE(dbo.Users.department = @DepId) AND (dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered >= @DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                                + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                                }
                                else if (level == 1)
                                {
                                    TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                    + "dbo.Report.Comment, dbo.Report.CountSheet"
                                    + " FROM dbo.Report INNER JOIN"
                                    + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                    + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                    + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                                + " WHERE(dbo.Users.department = @DepId) AND(dbo.Proekti.Name_Project = @projName) AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered >= @DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                                + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                                }

                                TaskToProj.Parameters.Add(DeepId);
                                TaskToProj.Parameters.Add(ProjName);
                                TaskToProj.Parameters.Add(FullPath);
                                TaskToProj.Parameters.Add(entUserId);
                                TaskToProj.Parameters.Add(DateStart);
                                TaskToProj.Parameters.Add(DateEnd);

                                UpdateGrid(TaskToProj);
                            }
                            else //1234
                            {
                                if (level > 1)
                                {
                                    TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                                + " WHERE(dbo.Users.department = @DepId) AND (dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered >= @DStart)"
                                + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                                }
                                else if (level == 1)
                                {
                                    TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                                + " WHERE(dbo.Users.department = @DepId) AND(dbo.Proekti.Name_Project = @projName)  AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered >= @DStart)"
                                + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                                }

                                TaskToProj.Parameters.Add(DeepId);
                                TaskToProj.Parameters.Add(ProjName);
                                TaskToProj.Parameters.Add(FullPath);
                                TaskToProj.Parameters.Add(entUserId);
                                TaskToProj.Parameters.Add(DateStart);

                                UpdateGrid(TaskToProj);
                            }
                        }
                        else if (checkBox5.Checked == true)//1235
                        {
                            if (level > 1)
                            {
                                TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND (dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }
                            else if (level == 1)
                            {
                                TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Proekti.Name_Project = @projName) AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }
                            TaskToProj.Parameters.Add(DeepId);
                            TaskToProj.Parameters.Add(ProjName);
                            TaskToProj.Parameters.Add(FullPath);
                            TaskToProj.Parameters.Add(entUserId);
                            TaskToProj.Parameters.Add(DateEnd);

                            UpdateGrid(TaskToProj);
                        }
                        else //123
                        {
                            if (level > 1)
                            {
                                TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND (dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.Id_user = @EntUserId)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }
                            else if (level == 1)
                            {
                                TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Proekti.Name_Project = @projName) AND(dbo.Report.Id_user = @EntUserId)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }
                            TaskToProj.Parameters.Add(DeepId);
                            TaskToProj.Parameters.Add(ProjName);
                            TaskToProj.Parameters.Add(FullPath);
                            TaskToProj.Parameters.Add(entUserId);

                            UpdateGrid(TaskToProj);
                        }
                    }
                    else if(checkBox4.Checked == true)
                    {
                        if (checkBox5.Checked == true)//1245
                        {
                            if (level > 1)
                            {
                                TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND (dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered >= @DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }
                            else if (level == 1)
                            {
                                TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Proekti.Name_Project = @projName) AND(dbo.Report.DateEntered >= @DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }

                            TaskToProj.Parameters.Add(DeepId);
                            TaskToProj.Parameters.Add(ProjName);
                            TaskToProj.Parameters.Add(FullPath);
                            TaskToProj.Parameters.Add(DateStart);
                            TaskToProj.Parameters.Add(DateEnd);

                            UpdateGrid(TaskToProj);
                        }
                        else//124
                        {
                            if (level > 1)
                            {
                                TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND (dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered >= @DStart)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }
                            else if (level == 1)
                            {
                                TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Proekti.Name_Project = @projName) AND(dbo.Report.DateEntered >= @DStart)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }

                            TaskToProj.Parameters.Add(DeepId);
                            TaskToProj.Parameters.Add(ProjName);
                            TaskToProj.Parameters.Add(FullPath);
                            TaskToProj.Parameters.Add(DateStart);

                            UpdateGrid(TaskToProj);
                        }
                    }
                    else if(checkBox5.Checked == true)//125
                    {
                        if (level > 1)
                        {
                            TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }
                        else if (level == 1)
                        {
                            TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Proekti.Name_Project = @projName) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }

                        TaskToProj.Parameters.Add(DeepId);
                        TaskToProj.Parameters.Add(ProjName);
                        TaskToProj.Parameters.Add(FullPath);
                        TaskToProj.Parameters.Add(DateEnd);

                        UpdateGrid(TaskToProj);
                    }
                    else //12
                    {
                        if (level > 1)
                        {
                            TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND (dbo.Report.PuthToNode = @fullPath)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }
                        else if (level == 1)
                        {
                            TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Proekti.Name_Project = @projName)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }

                        TaskToProj.Parameters.Add(DeepId);
                        TaskToProj.Parameters.Add(ProjName);
                        TaskToProj.Parameters.Add(FullPath);

                        UpdateGrid(TaskToProj);
                    }
                }
                else if (comboBox1.Text.Length != 0)
                {
                    if (checkBox4.Checked == true)
                    {
                        if (checkBox5.Checked == true)//1345
                        {
                            if (level > 1)
                            {
                                TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered >= @DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }
                            else if (level == 1)
                            {
                                TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Proekti.Name_Project = @projName) AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered >= @DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }

                            TaskToProj.Parameters.Add(DeepId);
                            TaskToProj.Parameters.Add(ProjName); 
                            TaskToProj.Parameters.Add(FullPath);
                            TaskToProj.Parameters.Add(entUserId);
                            TaskToProj.Parameters.Add(DateStart);
                            TaskToProj.Parameters.Add(DateEnd);

                            UpdateGrid(TaskToProj);
                        }
                        else //134
                        {
                            if (level > 1)
                            {
                                TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                 + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                                 + " WHERE(dbo.Users.department = @DepId) AND(dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered >= @DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                                 + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }
                            else if (level == 1)
                            {
                                TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                 + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                                 + " WHERE(dbo.Users.department = @DepId) AND(dbo.Proekti.Name_Project = @projName) AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered >= @DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                                 + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }

                            TaskToProj.Parameters.Add(DeepId);
                            TaskToProj.Parameters.Add(ProjName);
                            TaskToProj.Parameters.Add(FullPath);
                            TaskToProj.Parameters.Add(entUserId);
                            TaskToProj.Parameters.Add(DateStart);
                            TaskToProj.Parameters.Add(DateEnd);

                            UpdateGrid(TaskToProj);
                        }
                    }
                    else if (checkBox5.Checked == true)//135
                    {
                        if (level > 1)
                        {
                            TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }
                        else if (level == 1)
                        {
                            TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Proekti.Name_Project = @projName) AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }

                        TaskToProj.Parameters.Add(DeepId);
                        TaskToProj.Parameters.Add(ProjName);
                        TaskToProj.Parameters.Add(FullPath);
                        TaskToProj.Parameters.Add(entUserId);
                        TaskToProj.Parameters.Add(DateEnd);

                        UpdateGrid(TaskToProj);
                    }
                    else //13
                    {
                        if (level > 1)
                        {
                            TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.Id_user = @EntUserId)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }
                        else if (level == 1)
                        {
                            TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Proekti.Name_Project = @projName) AND(dbo.Report.Id_user = @EntUserId)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }

                        TaskToProj.Parameters.Add(DeepId);
                        TaskToProj.Parameters.Add(ProjName);
                        TaskToProj.Parameters.Add(FullPath);
                        TaskToProj.Parameters.Add(entUserId);

                        UpdateGrid(TaskToProj);
                    }
                }
                else if (checkBox4.Checked == true)
                {
                    if (checkBox5.Checked == true)//145
                    {
                        if (level > 1)
                        {
                            TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                         + " WHERE(dbo.Users.department = @DepId) AND(dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered >= @DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                         + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }
                        else if (level == 1)
                        {
                            TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                         + " WHERE(dbo.Users.department = @DepId) AND(dbo.Proekti.Name_Project = @projName) AND(dbo.Report.DateEntered >= @DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                         + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }

                        TaskToProj.Parameters.Add(DeepId);
                        TaskToProj.Parameters.Add(ProjName);
                        TaskToProj.Parameters.Add(FullPath);
                        TaskToProj.Parameters.Add(DateStart);
                        TaskToProj.Parameters.Add(DateEnd);
                    }
                    else//14
                    {
                        if (level > 1)
                        {
                            TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                          + "dbo.Report.Comment, dbo.Report.CountSheet"
                                 + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                         + " WHERE(dbo.Users.department = @DepId) AND(dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered >= @DStart)"
                         + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }
                        else if (level == 1)
                        {
                            TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                          + "dbo.Report.Comment, dbo.Report.CountSheet"
                                 + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                         + " WHERE(dbo.Users.department = @DepId) AND(dbo.Proekti.Name_Project = @projName) AND(dbo.Report.DateEntered >= @DStart)"
                         + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }

                        TaskToProj.Parameters.Add(DeepId);
                        TaskToProj.Parameters.Add(ProjName);
                        TaskToProj.Parameters.Add(FullPath);
                        TaskToProj.Parameters.Add(DateStart);
                    }
                }
                else if (checkBox5.Checked == true)//15
                {
                    if (level > 1)
                    {
                        TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                         + " WHERE(dbo.Users.department = @DepId) AND(dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered <= @DEnd)"
                         + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    }
                    else if (level == 1)
                    {
                        TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                         + " WHERE(dbo.Users.department = @DepId) AND(dbo.Proekti.Name_Project = @projName) AND(dbo.Report.DateEntered <= @DEnd)"
                         + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    }

                    TaskToProj.Parameters.Add(DeepId);
                    TaskToProj.Parameters.Add(ProjName);
                    TaskToProj.Parameters.Add(FullPath);
                    TaskToProj.Parameters.Add(DateEnd);
                }
                else //1
                {
                    if (level > 1)
                    {
                        TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                         + " WHERE(dbo.Users.department = @DepId) AND(dbo.Report.PuthToNode = @fullPath)"
                         + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    }
                    else if (level == 1)
                    {
                        TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                         + " WHERE(dbo.Users.department = @DepId) AND(dbo.Proekti.Name_Project = @projName)"
                         + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    }

                    TaskToProj.Parameters.Add(DeepId);
                    TaskToProj.Parameters.Add(ProjName);
                    TaskToProj.Parameters.Add(FullPath);
                }
            }
            else if (TipeBild.Length != 0)
            {
                if (comboBox1.Text.Length != 0)
                {
                    if (checkBox4.Checked == true)
                    {
                        if(checkBox5.Checked == true)//2345
                        {
                            TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND (dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered >= @DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            TaskToProj.Parameters.Add(DeepId);
                            TaskToProj.Parameters.Add(FullPath);
                            TaskToProj.Parameters.Add(entUserId);
                            TaskToProj.Parameters.Add(DateStart);
                            TaskToProj.Parameters.Add(DateEnd);

                            UpdateGrid(TaskToProj);
                        }
                        else//234
                        {
                            TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND (dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered >= @DStart) "
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            TaskToProj.Parameters.Add(DeepId);
                            TaskToProj.Parameters.Add(FullPath);
                            TaskToProj.Parameters.Add(entUserId);
                            TaskToProj.Parameters.Add(DateStart);

                            UpdateGrid(TaskToProj);
                        }
                    }
                    else if (checkBox5.Checked == true)//235
                    {
                        TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND (dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        TaskToProj.Parameters.Add(DeepId);
                        TaskToProj.Parameters.Add(FullPath);
                        TaskToProj.Parameters.Add(entUserId);
                        TaskToProj.Parameters.Add(DateEnd);

                        UpdateGrid(TaskToProj);
                    }
                    else//23
                    {
                        TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND (dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.Id_user = @EntUserId)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        TaskToProj.Parameters.Add(DeepId);
                        TaskToProj.Parameters.Add(FullPath);
                        TaskToProj.Parameters.Add(entUserId);

                        UpdateGrid(TaskToProj);
                    }
                }
                else if (checkBox4.Checked == true)
                {
                    if (checkBox5.Checked == true)//245
                    {
                        TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND (dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered >= @DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        TaskToProj.Parameters.Add(DeepId);
                        TaskToProj.Parameters.Add(FullPath);
                        TaskToProj.Parameters.Add(DateStart);
                        TaskToProj.Parameters.Add(DateEnd);

                        UpdateGrid(TaskToProj);
                    }
                    else//24
                    {
                        TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND (dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered >= @DStart)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        TaskToProj.Parameters.Add(DeepId);
                        TaskToProj.Parameters.Add(FullPath);
                        TaskToProj.Parameters.Add(DateStart);

                        UpdateGrid(TaskToProj);
                    }
                }
                else if (checkBox5.Checked == true)//25
                {
                    TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND (dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    TaskToProj.Parameters.Add(DeepId);
                    TaskToProj.Parameters.Add(FullPath);
                    TaskToProj.Parameters.Add(DateEnd);

                    UpdateGrid(TaskToProj);
                }
                else//2
                {
                    TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND (dbo.Report.PuthToNode = @fullPath)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    TaskToProj.Parameters.Add(DeepId);
                    TaskToProj.Parameters.Add(FullPath);

                    UpdateGrid(TaskToProj);
                }
            }
            else if (comboBox1.Text.Length != 0)
            {
                if (checkBox4.Checked == true)
                {
                    if (checkBox5.Checked == true)//345
                    {
                        TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered >= @DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        TaskToProj.Parameters.Add(DeepId);
                        TaskToProj.Parameters.Add(entUserId);
                        TaskToProj.Parameters.Add(DateStart);
                        TaskToProj.Parameters.Add(DateEnd);

                        UpdateGrid(TaskToProj);
                    }
                    else//34
                    {
                        TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered >= @DStart)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        TaskToProj.Parameters.Add(DeepId);
                        TaskToProj.Parameters.Add(entUserId);
                        TaskToProj.Parameters.Add(DateStart);

                        UpdateGrid(TaskToProj);
                    }
                }
                else if (checkBox5.Checked == true)//35
                {
                    TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Report.Id_user = @EntUserId) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    TaskToProj.Parameters.Add(DeepId);
                    TaskToProj.Parameters.Add(entUserId);
                    TaskToProj.Parameters.Add(DateEnd);

                    UpdateGrid(TaskToProj);
                }
                else//3
                {
                    TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Report.Id_user = @EntUserId)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    TaskToProj.Parameters.Add(DeepId);
                    TaskToProj.Parameters.Add(entUserId);

                    UpdateGrid(TaskToProj);
                }
            }
            else if (checkBox4.Checked == true)
            {
                if (checkBox5.Checked == true)//45
                {
                    TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Report.DateEntered >= @DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    TaskToProj.Parameters.Add(DeepId);
                    TaskToProj.Parameters.Add(DateStart);
                    TaskToProj.Parameters.Add(DateEnd);

                    UpdateGrid(TaskToProj);
                }
                else//4
                {
                    TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Report.DateEntered >= @DStart)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    TaskToProj.Parameters.Add(DeepId);
                    TaskToProj.Parameters.Add(DateStart);

                    UpdateGrid(TaskToProj);
                }
            }
            else if (checkBox5.Checked == true)//5
            {
                TaskToProj.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE(dbo.Users.department = @DepId) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                TaskToProj.Parameters.Add(DeepId);
                TaskToProj.Parameters.Add(DateEnd);

                UpdateGrid(TaskToProj);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = -1;

            ProjName2 = "";
            TypeBild2 = "";
            fullPath2 = "";
            
            treeView2.CollapseAll();

            checkBox4.Checked = false;
            checkBox5.Checked = false;

            SqlConnection conn = new SqlConnection(conString);
            SqlCommand AllTasks = new SqlCommand();
            SqlParameter depId = new SqlParameter("@DepId", DepId);

            AllTasks.Connection = conn;
            AllTasks.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                         + " FROM            dbo.Proekti INNER JOIN"
                         + " dbo.Report ON dbo.Proekti.Id_project = dbo.Report.Id_Project INNER JOIN"
                         + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                         + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                         + " WHERE(dbo.Users.department = @DepId)";

            AllTasks.Parameters.Add(depId);

            UpdateGrid(AllTasks);
        }

        public void DataToListBox()
        {
            SqlConnection sqldbConnection = new SqlConnection(conString);
            SqlCommand sqlCommand = new SqlCommand();
            SqlCommand sqlUser = new SqlCommand();
            SqlCommand sqlTypeWork = new SqlCommand();
            SqlCommand sqlSection = new SqlCommand();

            sqlCommand.Connection = sqldbConnection;
            sqlCommand.CommandText = "SELECT Id_project, Name_Project"
                    + " FROM dbo.Proekti"
                + " WHERE(InVork = 1)";

            SqlParameter userid = new SqlParameter("@userid", UserID);
            SqlParameter projId = new SqlParameter("@PId", PId);

            sqlUser.CommandText = "SELECT Surename FROM dbo.Users WHERE(department = @DepId)"
                + " ORDER BY dbo.Users.Surename";
            SqlParameter depId2 = new SqlParameter("@DepId", DepId);
            sqlUser.Parameters.Add(depId2);
            sqlUser.Connection = sqldbConnection;


            sqlTypeWork.CommandText = "SELECT dbo.TypeWork.TypeWork FROM dbo.TypeWork";
            sqlSection.CommandText = "SELECT dbo.Sections.Section_Name FROM dbo.Sections";

            sqlTypeWork.Connection = sqldbConnection;
            sqlSection.Connection = sqldbConnection;

            try
            {
                sqldbConnection.Open();
                /*SqlDataReader ProjReader = sqlCommand.ExecuteReader();
                while (ProjReader.Read())
                {
                    string s = ProjReader.GetString(ProjReader.GetOrdinal("Name_Project"));
                    comboBox1.Items.Add(s);
                    //comboBox2.Items.Add(s);
                }
                ProjReader.Close();*/

                SqlDataReader UserReader = sqlUser.ExecuteReader();
                while (UserReader.Read())
                {
                    string s = UserReader[0].ToString();
                    comboBox1.Items.Add(s);
                }
                UserReader.Close();

                SqlDataReader TypeWorkReader = sqlTypeWork.ExecuteReader();
                while (TypeWorkReader.Read())
                {
                    string s = TypeWorkReader[0].ToString();
                    comboBox6.Items.Add(s);
                }
                TypeWorkReader.Close();

                SqlDataReader SectionReader = sqlSection.ExecuteReader();
                while (SectionReader.Read())
                {
                    string s = SectionReader[0].ToString();
                    comboBox7.Items.Add(s);
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

        private string PuthToSave()
        {
            string myPath = "";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
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

            ExportDataToExcel(worksheet, myQuary);

            NBook.SaveAs(PuthToSave, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault);
            NBook.Close(false);
        }

        //Экспорт данных в эксель
        private void ExportDataToExcel(Microsoft.Office.Interop.Excel.Worksheet workshee, string myQuary)
        {
            SqlConnection DbConn = new SqlConnection(conString);
            SqlCommand DbSomeID = new SqlCommand();
            DbSomeID.CommandText = myQuary;
            DbSomeID.Connection = DbConn;
            SqlDataReader oleDbDataReader;
            try
            {
                DbConn.Open();
                oleDbDataReader = DbSomeID.ExecuteReader();
                int row = 1;

                workshee.Cells[row, "A"] = oleDbDataReader.GetName(0).ToString();
                workshee.Cells[row, "B"] = oleDbDataReader.GetName(1).ToString();
                workshee.Cells[row, "C"] = oleDbDataReader.GetName(2).ToString();
                workshee.Cells[row, "D"] = oleDbDataReader.GetName(3).ToString();
                workshee.Cells[row, "E"] = oleDbDataReader.GetName(4).ToString();
                workshee.Cells[row, "F"] = oleDbDataReader.GetName(5).ToString();
                workshee.Cells[row, "G"] = oleDbDataReader.GetName(6).ToString();
                workshee.Cells[row, "H"] = oleDbDataReader.GetName(7).ToString();
                workshee.Cells[row, "I"] = oleDbDataReader.GetName(8).ToString();
                workshee.Cells[row, "J"] = oleDbDataReader.GetName(9).ToString();
                workshee.Cells[row, "K"] = oleDbDataReader.GetName(10).ToString();
                workshee.Cells[row, "L"] = oleDbDataReader.GetName(11).ToString();

                row++;

                while (oleDbDataReader.Read())
                {
                    workshee.Cells[row, "A"] = oleDbDataReader.GetValue(0).ToString();
                    workshee.Cells[row, "B"] = oleDbDataReader.GetValue(1).ToString();
                    workshee.Cells[row, "C"] = oleDbDataReader.GetValue(2).ToString();
                    workshee.Cells[row, "D"] = oleDbDataReader.GetValue(3).ToString();
                    workshee.Cells[row, "E"] = oleDbDataReader.GetValue(4).ToString();
                    workshee.Cells[row, "F"] = oleDbDataReader.GetValue(5).ToString();
                    workshee.Cells[row, "G"] = oleDbDataReader.GetValue(6).ToString();
                    workshee.Cells[row, "H"] = oleDbDataReader.GetValue(7).ToString();
                    workshee.Cells[row, "I"] = oleDbDataReader.GetValue(8).ToString();
                    workshee.Cells[row, "J"] = oleDbDataReader.GetValue(9).ToString();
                    workshee.Cells[row, "K"] = oleDbDataReader.GetValue(10).ToString();
                    workshee.Cells[row, "L"] = oleDbDataReader.GetValue(11).ToString();
                    row++;
                }
                oleDbDataReader.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            finally
            {
                DbConn.Close();
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox6.Text.Length != 0) richTextBox1.Enabled = true;
            else richTextBox1.Enabled = false;
        }

        /*private void comboBox2_DropDown(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            SqlConnection DbConn = new SqlConnection(conString);
            SqlCommand DbComand = new SqlCommand();
            SqlParameter uId = new SqlParameter("@UId", UId);
            DbComand.CommandText = " SELECT dbo.Proekti.Name_Project"
            + " FROM(dbo.Users INNER JOIN (dbo.Proekti INNER JOIN dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project) ON dbo.Users.Id_users = dbo.Zadan.User_Give_out)"
            + " WHERE(((dbo.Zadan.User_Give_out) =@Uid))"
            + " GROUP BY dbo.Proekti.Name_Project";
            DbComand.Parameters.Add(uId);
            DbComand.Connection = DbConn;

            try
            {
                DbConn.Open();
                SqlDataReader DbReader = DbComand.ExecuteReader();
                while (DbReader.Read())
                {
                    string s = DbReader.GetString(DbReader.GetOrdinal("Name_Project"));
                    comboBox2.Items.Add(s);
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            finally
            {
                DbConn.Close();
            }
        }*/

        /*private void comboBox2_DropDownClosed(object sender, EventArgs e)
        {
            comboBox3.Items.Clear();
            if (comboBox2.Text.Length != 0)
            {
                string SelectProjName = comboBox2.SelectedItem.ToString();
                SqlConnection DbConn = new SqlConnection(conString);
                SqlCommand DbComand = new SqlCommand();
                SqlParameter selProjName = new SqlParameter("@SelProjName", SelectProjName);
                SqlParameter uID = new SqlParameter("@UId", UId);

                DbComand.CommandText = "SELECT dbo.Zadan.Task_text"
                + " FROM(dbo.Zadan INNER JOIN"
                + " dbo.Proekti ON dbo.Zadan.Id_project = dbo.Proekti.Id_project)"
                + " WHERE(dbo.Proekti.Name_Project = @SelProjName) AND(dbo.Zadan.User_Give_out = @UId)";

                DbComand.Parameters.Add(selProjName);
                DbComand.Parameters.Add(uID);

                DbComand.Connection = DbConn;
                try
                {
                    DbConn.Open();
                    SqlDataReader DbReader = DbComand.ExecuteReader();
                    while (DbReader.Read())
                    {
                        string s = DbReader.GetString(DbReader.GetOrdinal("Task_text"));
                        comboBox3.Items.Add(s);
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
                finally { DbConn.Close(); }
                comboBox3.Enabled = true;
            }
        }*/

        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                if (MessageBox.Show("Вы действительно хотите закрыть форму ? ", "Внимание", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    Application.Exit();
                }
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true) dateTimePicker3.Enabled = true;
            else dateTimePicker3.Enabled = false;
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true) dateTimePicker4.Enabled = true;
            else dateTimePicker4.Enabled = false;
        }

        /*private void comboBox3_DropDownClosed(object sender, EventArgs e)
        {
            if ((comboBox2.Text.Length != 0) && (comboBox3.Text.Length != 0))
            {
                PId = comboBox2.Text;
                TId = comboBox3.Text;
                string NameTask = TId;
                
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
                                            + " FROM dbo.Proekti INNER JOIN"
                                            + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                            + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users"
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
        }*/

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

        public string NameProj(TreeView treeView)
        {
            string rez = "";

            TreeNode SelectNode = treeView.SelectedNode;
            string path = treeView.SelectedNode.FullPath;

            string[] ArrVal = path.Split('/');
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
                while(sqlDataReader.Read())
                {
                    rez = sqlDataReader[0].ToString();
                }
            }
            catch(SqlException ex) { MessageBox.Show("Произошла ошибка выполнения команды: " + ex, "Ошибка выполнения сценария", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally { conn.Close(); }
            return rez;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int level = level1;
            string fullPath = fullPath1;

            string PuthToNode = fullPath1;
            string NameSelProj = ProjName1;
            string TypeBild = TypeBild1;

            //if ((NameSelProj == "") || (NameSelProj == null) || (TypeBild == "") || (TypeBild == null) || (NameSelProj == TypeBild)) MessageBox.Show("Не выбрана площадка или сооружение. Сделайте выбор и повторите попытку.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            if ((NameSelProj == "") || (NameSelProj == null) || (NameSelProj == TypeBild)) MessageBox.Show("Не выбрана площадка или сооружение. Сделайте выбор и повторите попытку.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else if ((TypeBild.Length != 0) && (comboBox6.Text.Length != 0) && (comboBox7.Text.Length != 0) && (textBox1.Text.Length != 0))
            {
                string WorkTimeVal = textBox1.Text.ToString().Replace('.', ',');
                float FWorkTimeVal = Convert.ToSingle(WorkTimeVal);

                DateTime dtime = DateTime.Now;
                string myDate = dtime.ToString("d");
                myDate = myDate.Replace('.', '-');
                string myTime = dtime.ToLongTimeString();

                string Section = comboBox7.Text.ToString();
                string WorkTime = textBox1.Text.ToString();
                string coment = richTextBox1.Text.ToString();
                string CountSheet = textBox2.Text.ToString();
                string TypeWork = comboBox6.Text.ToString();

                SqlConnection DbConn = new SqlConnection(conString);
                SqlCommand DbComandInsert = new SqlCommand();

                DbComandInsert.CommandText = "INSERT INTO dbo.Report(Id_Project, Id_user, TypeWork, Section, TypeBild, Comment, CountSheet, TimeWork, DateEntered, TimeEntered1,PuthToNode)"
                + " VALUES"
                + " (@pid, @uid, @TWork, @Section, @TBild, @Coment, @cSheet, @WorkTime, @Dtime, @timeEntered, @PuthToNode)";

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
                SqlParameter timeEntered = new SqlParameter("@timeEntered", dtime);
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
                    
                    comboBox6.SelectedIndex = -1;
                    comboBox7.SelectedIndex = -1;
                    richTextBox1.Text = "";
                    textBox1.Text = "";
                    textBox2.Text = "";
                }
                catch (Exception ex) { MessageBox.Show("Запись добавить не удалось!!! " + ex, "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                finally { DbConn.Close(); }
            }
            else MessageBox.Show("Не заполнен один из параметров! Заполните все и повторите попытку", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string path = "";
            DateTime now = DateTime.Now;
            string NameFile = now.ToString("d");

            saveFileDialog1.FileName = NameFile;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                path = saveFileDialog1.FileName;
                if (!Directory.Exists(Path.GetDirectoryName(path))) Directory.CreateDirectory(Path.GetDirectoryName(path));
                
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

        public int UserId(string NameUser)
        {
            int EntUseId = -1;
            SqlConnection DbConn = new SqlConnection(conString);
            SqlCommand DbComand = new SqlCommand();
            SqlParameter EnterUser = new SqlParameter("@EntUser", NameUser);
            DbComand.CommandText = "SELECT Id_users FROM Users WHERE  (Surename = @EntUser)";
            DbComand.Parameters.Add(EnterUser);
            DbComand.Connection = DbConn;
            try
            {
                DbConn.Open();
                SqlDataReader DbReader = DbComand.ExecuteReader();
                while (DbReader.Read())
                {
                    EntUseId = Convert.ToInt32(DbReader[0]);
                    //EntUseId = DbReader.GetString(DbReader.GetOrdinal("Id_users"));
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            finally
            {
                DbConn.Close();
            }

            return EntUseId;
        }

        private void AutoSizeGridColumn(DataGridView myGrid)
        {
            myGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count != 1) return;

            splitContainer1.Panel1Collapsed = false;

            var selRow = dataGridView2.SelectedRows;
            int idRowToUpdate = 0;

            idRowToUpdate = Convert.ToInt32(selRow[0].Cells["Aproval_Id"].Value);
            label6.Text = selRow[0].Cells["Surename"].Value.ToString();

            var a = selRow[0].Cells["DateStart"].Value.ToString().Split(' ');
            label8.Text = a[0];
            a = selRow[0].Cells["DateEnd"].Value.ToString().Split(' ');
            label9.Text = a[0];

            label10.Text = selRow[0].Cells["TypeOfAbsence"].Value.ToString();
            label16.Text = idRowToUpdate.ToString();

        }

        /// <summary>
        /// Скрыть/развернуть чать сплит контейнера
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void label11_Click(object sender, EventArgs e)
        {
            if (label11.Text == "<")
            {
                label11.Text = ">";
                splitContainer1.Panel1Collapsed = true;
            }
                else
            {
                label11.Text = "<";
                splitContainer1.Panel1Collapsed = false;
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab.Name == "tabPage2")
            {
                SqlConnection conn = new SqlConnection(conString);
                SqlCommand updateGrid = new SqlCommand
                {
                    Connection = conn,
                    CommandText = "SELECT        dbo.AbsenceRequestTable.Aproval_Id, dbo.Users.Surename, dbo.AbsenceRequestTable.TypeOfAbsence, dbo.AbsenceRequestTable.DateStart,"
                    + " dbo.AbsenceRequestTable.DateEnd, dbo.AbsenceRequestTable.AprovalFlag"
                    + " FROM            dbo.AbsenceRequestTable INNER JOIN"
                    + " dbo.Users ON dbo.AbsenceRequestTable.id_EnterUser = dbo.Users.Id_users"
                    + " WHERE(dbo.AbsenceRequestTable.id_UserDepart = @UserIdDepart)"
                };

                SqlParameter UserIdDepart = new SqlParameter("@UserIdDepart", DepId);
                updateGrid.Parameters.Add(UserIdDepart);

                UpdateGrid(updateGrid, dataGridView2);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (label6.Text == "") return;

            string fio = label6.Text;
            DateTime dateAproval = DateTime.Now;
            int flag = 1;
            int aprovalId = Convert.ToInt32(label16.Text);

            SqlConnection conn = new SqlConnection(conString);
            SqlCommand updateDataBase = new SqlCommand
            {
                Connection = conn,
                CommandText = "UPDATE [dbo].[AbsenceRequestTable]"
                +" Set [id_UserAproval] = @AprovalUserID, [AprovalUserName] = @FIO, [AprovalFlag] = @AprovalFlag, [DateAproval] = @DateAproval"
                + " Where Aproval_Id = @AprovalId"
            };

            SqlParameter AprovalUserID = new SqlParameter("@AprovalUserID", UId);
            SqlParameter FIO = new SqlParameter("@FIO", fio);
            SqlParameter AprovalFlag = new SqlParameter("@AprovalFlag", flag);
            SqlParameter DateAproval = new SqlParameter("@DateAproval", dateAproval);
            SqlParameter AprovalId = new SqlParameter("@AprovalId", aprovalId);

            updateDataBase.Parameters.Add(AprovalUserID);
            updateDataBase.Parameters.Add(FIO);
            updateDataBase.Parameters.Add(AprovalFlag);
            updateDataBase.Parameters.Add(DateAproval);
            updateDataBase.Parameters.Add(AprovalId);

            try
            {
                conn.Open();
                updateDataBase.ExecuteNonQuery();
            }
            catch (Exception ex)
            { MessageBox.Show("" + ex); }
            finally { conn.Close(); }

            SqlCommand updateGrid = new SqlCommand
            {
                Connection = conn,
                CommandText = "SELECT        dbo.AbsenceRequestTable.Aproval_Id, dbo.Users.Surename, dbo.AbsenceRequestTable.TypeOfAbsence, dbo.AbsenceRequestTable.DateStart,"
                    + " dbo.AbsenceRequestTable.DateEnd, dbo.AbsenceRequestTable.AprovalFlag"
                    + " FROM            dbo.AbsenceRequestTable INNER JOIN"
                    + " dbo.Users ON dbo.AbsenceRequestTable.id_EnterUser = dbo.Users.Id_users"
                    + " WHERE(dbo.AbsenceRequestTable.id_UserDepart = @UserIdDepart)"
            };

            SqlParameter UserIdDepart = new SqlParameter("@UserIdDepart", DepId);
            updateGrid.Parameters.Add(UserIdDepart);

            UpdateGrid(updateGrid, dataGridView2);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (label6.Text == "") return;

            string nameFunction = "absence_UpdateData";

            string fio = label6.Text;
            DateTime dateAproval = DateTime.Now;
            int flag = 0;
            int aprovalId = Convert.ToInt32(label16.Text);
            
            
            using (SqlConnection conn = new SqlConnection(conString))
            {
                conn.Open();
                SqlCommand sqlCommand = new SqlCommand(nameFunction, conn)
                {
                    CommandType = System.Data.CommandType.StoredProcedure
                };
                
                SqlParameter AprovalUserID = new SqlParameter("@id_UserAproval", UId);
                SqlParameter FIO = new SqlParameter("@aprovalUserName", fio);
                SqlParameter AprovalFlag = new SqlParameter("@aprovalFlag", flag);
                SqlParameter DateAproval = new SqlParameter("@dateAproval", dateAproval);
                SqlParameter AprovalId = new SqlParameter("@rowToUpdate", aprovalId);

                sqlCommand.Parameters.Add(AprovalUserID);
                sqlCommand.Parameters.Add(FIO);
                sqlCommand.Parameters.Add(AprovalFlag);
                sqlCommand.Parameters.Add(DateAproval);
                sqlCommand.Parameters.Add(AprovalId);

                try
                {
                    sqlCommand.ExecuteNonQuery();
                }
                catch (Exception ex) { MessageBox.Show("" + ex); }
                finally { conn.Close(); }

                SqlCommand updateGrid = new SqlCommand
                {
                    Connection = conn,
                    CommandText = "SELECT        dbo.AbsenceRequestTable.Aproval_Id, dbo.Users.Surename, dbo.AbsenceRequestTable.TypeOfAbsence, dbo.AbsenceRequestTable.DateStart,"
                    + " dbo.AbsenceRequestTable.DateEnd, dbo.AbsenceRequestTable.AprovalFlag"
                    + " FROM            dbo.AbsenceRequestTable INNER JOIN"
                    + " dbo.Users ON dbo.AbsenceRequestTable.id_EnterUser = dbo.Users.Id_users"
                    + " WHERE(dbo.AbsenceRequestTable.id_UserDepart = @UserIdDepart)"
                };

                SqlParameter UserIdDepart = new SqlParameter("@UserIdDepart", DepId);
                updateGrid.Parameters.Add(UserIdDepart);

                UpdateGrid(updateGrid, dataGridView2);
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
                DataTable dataTable = (DataTable)(dataGridView1.DataSource);

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
                Microsoft.Office.Interop.Excel.Range RangeToEnter;
                
                //Заполнение диапазона в эксель из массива данных
                range1 = worksheet.Cells[1, 1] as Microsoft.Office.Interop.Excel.Range;
                range2 = worksheet.Cells[mRow, mColumn] as Microsoft.Office.Interop.Excel.Range;
                range = worksheet.get_Range(range1, range2);
                range.Value2 = dataExport;

                //Добавление строки с фильтрами
                Microsoft.Office.Interop.Excel.Range line = (Microsoft.Office.Interop.Excel.Range)worksheet.Rows[1];
                line.Insert();
                line.Insert();

                worksheet.Cells[1, 1].value2 = "Проект";
                worksheet.Cells[1, 2].value2 = "Сооружение";
                worksheet.Cells[1, 3].value2 = "Тип работы";
                worksheet.Cells[1, 4].value2 = "Раздел";
                worksheet.Cells[1, 5].value2 = "Автор отчета";
                worksheet.Cells[1, 6].value2 = "Трудозатраты";
                worksheet.Cells[1, 7].value2 = "Дата заполнения";
                worksheet.Cells[1, 8].value2 = "Комментарий";
                worksheet.Cells[1, 9].value2 = "Кол-во листов";
                
                Microsoft.Office.Interop.Excel.Range rangetofilter1 = worksheet.Cells[2, 1] as Microsoft.Office.Interop.Excel.Range;
                Microsoft.Office.Interop.Excel.Range rangetofilter2 = worksheet.Cells[2, 9] as Microsoft.Office.Interop.Excel.Range;

                RangeToEnter = worksheet.get_Range(rangetofilter1, rangetofilter2);
                RangeToEnter.AutoFilter(1);

                RangeToEnter = worksheet.Cells[2, 7] as Microsoft.Office.Interop.Excel.Range;
                RangeToEnter.EntireColumn.NumberFormat = "General";

                //Для столбца с датой устанавливаем необходимый формат
                range1 = worksheet.Cells[3, 7] as Microsoft.Office.Interop.Excel.Range;
                range2 = worksheet.Cells[mRow, 7] as Microsoft.Office.Interop.Excel.Range;

                range = worksheet.get_Range(range1, range2);
                range.EntireColumn.NumberFormat = "DD/MM/YYYY";

                //Авторазмер столбцов
                worksheet.Columns.EntireColumn.AutoFit();

                //Сохранение и закрытие книги
                NBook.SaveAs(Path, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
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

        private void treeView2_AfterSelect(object sender, TreeViewEventArgs e)
        {
            treeView2.PathSeparator = "/";
            fullPath2 = treeView2.SelectedNode.FullPath.ToString();
            var arr = fullPath2.Split('/');
            if (arr.Length >= 2)
            {
                ProjName2 = arr[1];
                TypeBild2 = treeView2.SelectedNode.Text;
                level2 = treeView2.SelectedNode.Level;
            }
            else
            {
                ProjName2 = "";
                TypeBild2 = "";
                level2 = 0;
            }
        }

        private void myTreeViewWork(TreeView treeView)
        {
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
                         +" dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj INNER JOIN"
                         +" dbo.TypeBild ON dbo.ProjBild.Id_TypeBilding = dbo.TypeBild.id_TypeBild"
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
                string[,] myData = new string[5000, 2];

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
                AddNodesTree(FinnArr, RowCountList, i , treeView);
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
            for (int i = 0; i< RowCount; i++)
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
                if (LiltleArr.Length ==2)
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

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && (e.KeyChar != 8) && (e.KeyChar != 46)) e.Handled = true;
            else
            {
                if (e.KeyChar == 46)
                {
                    if (textBox2.Text.Length > 0)
                    {
                        if (textBox2.Text.IndexOf('.') > -1) e.Handled = true;
                    }
                    else e.Handled = true;
                }
            }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            treeView1.PathSeparator = "/";
            fullPath1 = treeView1.SelectedNode.FullPath.ToString();
            var arr = fullPath1.Split('/');
            if (arr.Length >= 2)
            {
                ProjName1 = arr[1];
                TypeBild1 = treeView1.SelectedNode.Text;
                level1 = treeView1.SelectedNode.Level;

                SqlConnection conn = new SqlConnection(conString);
                SqlCommand idProj = new SqlCommand();
                SqlParameter NProj = new SqlParameter("@NProj", ProjName1);
                idProj.CommandText = "SELECT        Id_project, Name_Project FROM dbo.Proekti WHERE(Name_Project = @NProj)";
                idProj.Connection = conn;
                idProj.Parameters.Add(NProj);

                try
                {
                    conn.Open();
                    SqlDataReader readerIdProj = idProj.ExecuteReader();
                    if (readerIdProj.Read())
                    {
                        PId = readerIdProj[0].ToString();
                    }
                }
                catch (SqlException ex) { MessageBox.Show("" + ex); }
                finally { conn.Close(); }
            }
            else
            {
                ProjName1 = "";
                TypeBild1 = "";
                level1 = 0;
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
                    if (Val == myData[j,1])
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
            int[,] GoodArr = new int[RowCount,3];
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
            int[,] rez = new int[RowCount,3];

            for (int i = 0; i < RowCount - 1; i++)
            {
                //if (myArr[i,0] != null)
                {
                    int min = i;
                    for (int j = i + 1; j < RowCount; j++)
                    {
                        if (myArr[min,0] > myArr[j,0])
                        {
                            min = j;
                        }
                    }
                    int temp0 = myArr[min,0];
                    int temp1 = myArr[min,1];
                    int temp2 = myArr[min,2];

                    myArr[min,0] = myArr[i,0];
                    myArr[min,1] = myArr[i,1];
                    myArr[min,2] = myArr[i,2];

                    myArr[i,0] = temp0;
                    myArr[i,1] = temp1;
                    myArr[i,2] = temp2;
                }
            }
            return rez = myVoid2(myArr, RowCount); ;
        }
        public int[,] myVoid2(int[,] myArr, int RowCount)
        {
            int[,] rez = new int[RowCount,3];

            for (int i = 0; i < RowCount - 1; i++)
            {
                int min = i;
                int firstIndex = myArr[min,0];
                int TherdIndex = myArr[min,2];

                for (int j = i + 1; j < RowCount; j++)
                {
                    if (myArr[j,0] == firstIndex)
                    {
                        if (myArr[min, 1] > myArr[j, 1])
                        {
                            min = j;
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
            int[,] rez = new int[RowCount,3];

            for (int i = 0; i < RowCount - 1; i++)
            {
                //if (myArr[i,0] != null)
                {
                    int min = i;
                    int firstIndex = myArr[min,0];
                    int SecondIndex = myArr[min,1];

                    for (int j = i + 1; j < RowCount; j++)
                    {
                        if (myArr[j,0] == firstIndex)
                        {
                            if (myArr[j,1] == SecondIndex)
                            {
                                if (myArr[min,2] > myArr[j,2])
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

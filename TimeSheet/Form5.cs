using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace TimeSheet
{
    public partial class Form5 : Form
    {
        public string UId, DepId, TId, PId, conString;
        
        public Form5(string UId, string DepId, string conString)
        {
            this.UId = UId;
            this.DepId = DepId;
            this.conString = conString;
            InitializeComponent();
        }
        
        private void comboBox1_DropDownClosed(object sender, EventArgs e)
        {
            if (comboBox1.Text.Length != 0)
            {
                string NameProj = comboBox1.Text.ToString();
                comboBox2.Items.Clear();
                comboBox3.Items.Clear();

                SqlConnection DbConn = new SqlConnection(conString);
                SqlCommand DbComand = new SqlCommand();
                SqlCommand DbComandUser = new SqlCommand();

                SqlParameter nameProj = new SqlParameter("@NameProj", NameProj);
                SqlParameter nameProj2 = new SqlParameter("@NameProj2", NameProj);

                DbComand.CommandText = "SELECT dbo.Otdeli.department"
                + " FROM            dbo.Proekti INNER JOIN"
                         + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                         + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                         + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                + " GROUP BY dbo.Otdeli.department, dbo.Proekti.Name_Project"
                + " HAVING(dbo.Proekti.Name_Project = @NameProj)";

                DbComandUser.CommandText = "SELECT dbo.Users.Surename"
                + " FROM            dbo.Proekti INNER JOIN"
                         + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                         + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users"
                + " GROUP BY dbo.Proekti.Name_Project, dbo.Users.Surename"
                + " HAVING(dbo.Proekti.Name_Project = @NameProj2)";

                DbComand.Connection = DbConn;
                DbComand.Parameters.Add(nameProj);

                DbComandUser.Connection = DbConn;
                DbComandUser.Parameters.Add(nameProj2);
                try
                {
                    DbConn.Open();
                    SqlDataReader DbReader = DbComand.ExecuteReader();
                    while (DbReader.Read())
                    {
                        string s = DbReader.GetString(DbReader.GetOrdinal("department"));
                        comboBox2.Items.Add(s);
                    }
                    DbReader.Close();

                    SqlDataReader DbReaderUser = DbComandUser.ExecuteReader();
                    while (DbReaderUser.Read())
                    {
                        string s = DbReaderUser.GetString(DbReaderUser.GetOrdinal("Surename"));
                        comboBox3.Items.Add(s);
                    }
                    DbReaderUser.Close();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
                finally
                {
                    DbConn.Close();
                }
            }
        }

        private void comboBox2_DropDownClosed(object sender, EventArgs e)
        {
            if (comboBox2.Text.Length != 0)
            {
                SqlConnection DbConn = new SqlConnection(conString);
                SqlCommand DbComand = new SqlCommand();

                comboBox3.Items.Clear();
                string DepName = comboBox2.Text.ToString();

                SqlParameter depName = new SqlParameter("@DepName", DepName);

                DbComand.CommandText = "SELECT dbo.Users.Surename"
                + " FROM            dbo.Users INNER JOIN"
                         + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                + " WHERE(dbo.Otdeli.department = @DepName)";
                DbComand.Parameters.Add(depName);
                DbComand.Connection = DbConn;
                try
                {
                    DbConn.Open();
                    SqlDataReader DbReader = DbComand.ExecuteReader();
                    while (DbReader.Read())
                    {
                        string s = DbReader.GetString(DbReader.GetOrdinal("Surename"));
                        comboBox3.Items.Add(s);
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
                finally
                {
                    DbConn.Close();
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(conString);
            SqlCommand AllEntered = new SqlCommand();
            AllEntered.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, "
                         + " dbo.Vremay.Persent"
                         + " FROM dbo.Proekti INNER JOIN"
                         + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                         + " dbo.Vremay ON dbo.Proekti.Id_project = dbo.Vremay.Id_proj AND dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                         + " dbo.Users ON dbo.Vremay.Id_user = dbo.Users.Id_users";
            AllEntered.Connection = conn;

            UpdateGrid(AllEntered);
            DataToListBox();
        }

        private void Form5_FormClosing(object sender, FormClosingEventArgs e)
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

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true) dateTimePicker1.Enabled = true;
            else dateTimePicker1.Enabled = false;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true) dateTimePicker2.Enabled = true;
            else dateTimePicker2.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string PName = comboBox1.Text.ToString();
            string DName = comboBox2.Text.ToString();
            string UName = comboBox3.Text.ToString();
            DateTime DStart = Convert.ToDateTime(dateTimePicker1.Text.ToString());
            DateTime DEnd = Convert.ToDateTime(dateTimePicker2.Text.ToString());

            SqlConnection conn = new SqlConnection(conString);
            SqlCommand filterComand = new SqlCommand();

            SqlParameter ProjName = new SqlParameter("@ProjName", PName);
            SqlParameter DepName = new SqlParameter("@DepName", DName);
            SqlParameter UserName = new SqlParameter("@UserName", UName);
            SqlParameter DateStart = new SqlParameter("@DStart", DStart);
            SqlParameter DateEnd = new SqlParameter("@DEnd", DEnd);
            if (comboBox1.Text.Length != 0)
            {
                if (comboBox2.Text.Length != 0)
                {
                    if (comboBox3.Text.Length != 0)
                    {
                        if (checkBox1.Checked == true)
                        {
                            if (checkBox2.Checked == true)//все
                            {
                                filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Otdeli.department = @DepName) AND (dbo.Users.Surename = @UserName) AND(dbo.Vremay.date_entered >=@DStart) AND(dbo.Vremay.date_entered <= @DEnd)";
                                
                                filterComand.Parameters.Add(ProjName);
                                filterComand.Parameters.Add(DepName);
                                filterComand.Parameters.Add(UserName);
                                filterComand.Parameters.Add(DateStart);
                                filterComand.Parameters.Add(DateEnd);

                                UpdateGrid(filterComand);
                            }
                            else//Все кроме финиша
                            {
                                filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Otdeli.department = @DepName) AND (dbo.Users.Surename = @UserName) AND(dbo.Vremay.date_entered >=@DStart)";

                                filterComand.Parameters.Add(ProjName);
                                filterComand.Parameters.Add(DepName);
                                filterComand.Parameters.Add(UserName);
                                filterComand.Parameters.Add(DateStart);

                                UpdateGrid(filterComand);
                            }
                        }
                        else if (checkBox2.Checked == true)//Все кроме старта
                        {
                            filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Otdeli.department = @DepName) AND (dbo.Users.Surename = @UserName) AND(dbo.Vremay.date_entered <= @DEnd)";

                            filterComand.Parameters.Add(ProjName);
                            filterComand.Parameters.Add(DepName);
                            filterComand.Parameters.Add(UserName);
                            filterComand.Parameters.Add(DateEnd);

                            UpdateGrid(filterComand);
                        }
                        else////Проект, отдел и пользователь
                        {
                            filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Otdeli.department = @DepName) AND (dbo.Users.Surename = @UserName)";

                            filterComand.Parameters.Add(ProjName);
                            filterComand.Parameters.Add(DepName);
                            filterComand.Parameters.Add(UserName);

                            UpdateGrid(filterComand);
                        }
                    }
                    else if (checkBox1.Checked == true) //Проект, отдел и старт
                    {
                        filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Otdeli.department = @DepName) AND(dbo.Vremay.date_entered >=@DStart)";

                        filterComand.Parameters.Add(ProjName);
                        filterComand.Parameters.Add(DepName);
                        filterComand.Parameters.Add(DateStart);

                        UpdateGrid(filterComand);
                    }
                    else if (checkBox2.Checked == true)//Проект, отдел и финиш
                    {
                        filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Otdeli.department = @DepName) AND(dbo.Vremay.date_entered <= @DateEnd)";

                        filterComand.Parameters.Add(ProjName);
                        filterComand.Parameters.Add(DepName);
                        filterComand.Parameters.Add(DateEnd);

                        UpdateGrid(filterComand);
                    }
                    else//Проект и отдел
                    {
                        filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                               + " FROM            dbo.Proekti INNER JOIN"
                               + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                               + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                               + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                               + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Otdeli.department = @DepName)";

                        filterComand.Parameters.Add(ProjName);
                        filterComand.Parameters.Add(DepName);

                        UpdateGrid(filterComand);
                    }
                }
                else if (comboBox3.Text.Length != 0)//проект и пользователь
                {
                    if (checkBox1.Checked == true)
                    {
                        if (checkBox2.Checked == true)//проект, пользователь, страт и финиш
                        {
                            filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND (dbo.Users.Surename = @UserName) AND(dbo.Vremay.date_entered >=@DStart) AND(dbo.Vremay.date_entered <= @DEnd)";

                            filterComand.Parameters.Add(ProjName);
                            filterComand.Parameters.Add(UserName);
                            filterComand.Parameters.Add(DateStart);
                            filterComand.Parameters.Add(DateEnd);

                            UpdateGrid(filterComand);
                        }
                        else//проект, пользователь и страт
                        {
                            filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND (dbo.Users.Surename = @UserName) AND(dbo.Vremay.date_entered >=@DStart)";

                            filterComand.Parameters.Add(ProjName);
                            filterComand.Parameters.Add(UserName);
                            filterComand.Parameters.Add(DateStart);

                            UpdateGrid(filterComand);
                        }
                    }
                    else if (checkBox2.Checked == true)//проект, пользователь и финиш
                    {
                        filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND (dbo.Users.Surename = @UserName) AND(dbo.Vremay.date_entered <= @DEnd)";

                        filterComand.Parameters.Add(ProjName);
                        filterComand.Parameters.Add(UserName);
                        filterComand.Parameters.Add(DateEnd);

                        UpdateGrid(filterComand);
                    }
                    else
                    {
                        filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                         + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Users.Surename = @UserName)";
                        filterComand.Parameters.Add(ProjName);
                        filterComand.Parameters.Add(UserName);

                        UpdateGrid(filterComand);
                    }
                }
                else if (checkBox1.Checked == true)//проект и старт
                {
                    if (checkBox2.Checked == true)
                    {
                        filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Vremay.date_entered >=@DStart) AND(dbo.Vremay.date_entered <= @DEnd)";

                        filterComand.Parameters.Add(ProjName);
                        filterComand.Parameters.Add(DateStart);
                        filterComand.Parameters.Add(DateEnd);

                        UpdateGrid(filterComand);
                    }
                    else
                    {


                        filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                         + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Vremay.date_entered > @DStart)";
                        filterComand.Parameters.Add(ProjName);
                        filterComand.Parameters.Add(DateStart);

                        UpdateGrid(filterComand);
                    }
                }
                else if (checkBox2.Checked == true)//Проект и финиш
                {
                    filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                 + " FROM            dbo.Proekti INNER JOIN"
                                 + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                 + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                 + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                 + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                          + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Vremay.date_entered <= @DEnd)";
                    filterComand.Parameters.Add(ProjName);
                    filterComand.Parameters.Add(DateEnd);

                    UpdateGrid(filterComand);
                }
                else//Только проект
                {
                    filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                + " WHERE        (dbo.Proekti.Name_Project = @ProjName)";
                    filterComand.Parameters.Add(ProjName);

                    UpdateGrid(filterComand);
                }
            }
            else if (comboBox2.Text.Length != 0)//Проект не выбран
            {
                if (comboBox3.Text.Length != 0)
                {
                    if (checkBox1.Checked == true)
                    {
                        if (checkBox2.Checked == true)//все кроме проекта
                        {
                            filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Otdeli.department = @DepName) AND (dbo.Users.Surename = @UserName) AND(dbo.Vremay.date_entered >=@DStart) AND(dbo.Vremay.date_entered <= @DEnd)";

                            filterComand.Parameters.Add(DepName);
                            filterComand.Parameters.Add(UserName);
                            filterComand.Parameters.Add(DateStart);
                            filterComand.Parameters.Add(DateEnd);

                            UpdateGrid(filterComand);
                        }
                        else//234
                        {
                            filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Otdeli.department = @DepName) AND (dbo.Users.Surename = @UserName) AND(dbo.Vremay.date_entered >=@DStart)";

                            filterComand.Parameters.Add(DepName);
                            filterComand.Parameters.Add(UserName);
                            filterComand.Parameters.Add(DateStart);

                            UpdateGrid(filterComand);
                        }
                    }
                    else if (checkBox2.Checked == true)//235
                    {
                        filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Otdeli.department = @DepName) AND (dbo.Users.Surename = @UserName) AND(dbo.Vremay.date_entered <= @DEnd)";

                        filterComand.Parameters.Add(DepName);
                        filterComand.Parameters.Add(UserName);
                        filterComand.Parameters.Add(DateEnd);

                        UpdateGrid(filterComand);
                    }
                    else//23
                    {
                        filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Otdeli.department = @DepName) AND (dbo.Users.Surename = @UserName)";

                        filterComand.Parameters.Add(DepName);
                        filterComand.Parameters.Add(UserName);

                        UpdateGrid(filterComand);
                    }
                }
                else if (checkBox1.Checked == true)
                {
                    if (checkBox2.Checked == true)//245
                    {
                        filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Otdeli.department = @DepName) AND(dbo.Vremay.date_entered >=@DStart) AND(dbo.Vremay.date_entered <= @DEnd)";

                        filterComand.Parameters.Add(DepName);
                        filterComand.Parameters.Add(DateStart);
                        filterComand.Parameters.Add(DateEnd);

                        UpdateGrid(filterComand);
                    }
                    else//24
                    {
                        filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Otdeli.department = @DepName) AND(dbo.Vremay.date_entered >=@DStart)";

                        filterComand.Parameters.Add(DepName);
                        filterComand.Parameters.Add(DateStart);

                        UpdateGrid(filterComand);
                    }
                }
                else if (checkBox2.Checked == true)//25
                {
                    filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Otdeli.department = @DepName) AND(dbo.Vremay.date_entered <= @DEnd)";

                    filterComand.Parameters.Add(DepName);
                    filterComand.Parameters.Add(DateEnd);

                    UpdateGrid(filterComand);
                }
                else//2
                {
                    filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Otdeli.department = @DepName)";

                    filterComand.Parameters.Add(DepName);

                    UpdateGrid(filterComand);
                }
                    
            }
            else if (comboBox3.Text.Length != 0)
            {
                if (checkBox1.Checked == true)
                {
                    if (checkBox2.Checked == true)//345
                    {
                        filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Users.Surename = @UserName) AND(dbo.Vremay.date_entered >=@DStart) AND(dbo.Vremay.date_entered <= @DEnd)";
                        
                        filterComand.Parameters.Add(UserName);
                        filterComand.Parameters.Add(DateStart);
                        filterComand.Parameters.Add(DateEnd);

                        UpdateGrid(filterComand);
                    }
                    else//34
                    {
                        filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Users.Surename = @UserName) AND(dbo.Vremay.date_entered >=@DStart)";

                        filterComand.Parameters.Add(UserName);
                        filterComand.Parameters.Add(DateStart);

                        UpdateGrid(filterComand);
                    }
                }
                else if (checkBox2.Checked == true)//35
                {
                    filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Users.Surename = @UserName) AND(dbo.Vremay.date_entered <= @DEnd)";

                    filterComand.Parameters.Add(UserName);
                    filterComand.Parameters.Add(DateEnd);

                    UpdateGrid(filterComand);
                }
                else//3
                {
                    filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Users.Surename = @UserName)";

                    filterComand.Parameters.Add(UserName);

                    UpdateGrid(filterComand);
                }
            }
            else if (checkBox1.Checked == true)
            {
                if (checkBox2.Checked == true)//45
                {
                    filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Vremay.date_entered >=@DStart) AND(dbo.Vremay.date_entered <= @DEnd)";
                    
                    filterComand.Parameters.Add(DateStart);
                    filterComand.Parameters.Add(DateEnd);

                    UpdateGrid(filterComand);
                }
                else//4
                {
                    filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Vremay.date_entered >=@DStart)";

                    filterComand.Parameters.Add(DateStart);

                    UpdateGrid(filterComand);
                }
            }
            else if (checkBox2.Checked == true)//5
            {
                filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                                + " FROM            dbo.Proekti INNER JOIN"
                                + " dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                                + " dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                                + " dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                                + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                                 + " WHERE        (dbo.Vremay.date_entered <= @DEnd)";

                filterComand.CommandText = "SELECT dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, dbo.Vremay.Persent"
                + " FROM            dbo.Proekti INNER JOIN"
                         +" dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                         +" dbo.Vremay ON dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                         +" dbo.Users ON dbo.Zadan.User_Give_out = dbo.Users.Id_users INNER JOIN"
                         +" dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                         + " WHERE        (dbo.Vremay.date_entered = @DEnd)";
                filterComand.Parameters.Add(DateEnd);

                UpdateGrid(filterComand);
            }
            else
            {
                MessageBox.Show("Не выбран объект для фильтрации", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form5_Load(object sender, EventArgs e)
        {

            DataGridView GridTab1 = dataGridView1;

            AutoSizeGridColumn(GridTab1);
            SqlCommand AllEntered = new SqlCommand();



            AllEntered.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Zadan.Task_text, dbo.Vremay.Type_work, dbo.Users.Surename, dbo.Vremay.WorkTime, dbo.Vremay.date_entered, dbo.Vremay.Comment, "
                         + " dbo.Vremay.Persent"
                         +" FROM dbo.Proekti INNER JOIN"
                         +" dbo.Zadan ON dbo.Proekti.Id_project = dbo.Zadan.Id_project INNER JOIN"
                         +" dbo.Vremay ON dbo.Proekti.Id_project = dbo.Vremay.Id_proj AND dbo.Zadan.Id_task = dbo.Vremay.Id_task INNER JOIN"
                         +" dbo.Users ON dbo.Vremay.Id_user = dbo.Users.Id_users";

            DateTime myDate = DateTime.Today;
            dateTimePicker1.Value = myDate;
            dateTimePicker2.Value = myDate;

            checkBox1.Checked = false;
            checkBox2.Checked = false;

            dateTimePicker1.Enabled = false;
            dateTimePicker2.Enabled = false;
            
            UpdateGrid(AllEntered);
            GridHeaderName(GridTab1);
            DataToListBox();
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
            //}
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
            }
        }

        private void button3_Click(object sender, EventArgs e)
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

        public void DataToListBox()
        {
            comboBox1.Items.Clear();
            comboBox1.Text = "";

            comboBox2.Items.Clear();
            comboBox2.Text = "";

            comboBox3.Items.Clear();
            comboBox3.Text = "";

            SqlConnection DbConn = new SqlConnection(conString);
            SqlCommand AllProj = new SqlCommand();
            SqlCommand AllDep = new SqlCommand();
            SqlCommand AllUsers = new SqlCommand();
            string TaskToProj = "SELECT Name_Project FROM dbo.Proekti"
            + " GROUP BY Proekti.Name_Project";
            AllProj.CommandText = TaskToProj;
            AllProj.Connection = DbConn;

            AllDep.CommandText = "SELECT department FROM dbo.Otdeli";
            AllDep.Connection = DbConn;

            AllUsers.CommandText = "SELECT Surename FROM dbo.Users";
            AllUsers.Connection = DbConn;
            try
            {
                DbConn.Open();
                SqlDataReader Proj = AllProj.ExecuteReader();
                while (Proj.Read())
                {
                    string s = Proj.GetString(Proj.GetOrdinal("Name_Project"));
                    comboBox1.Items.Add(s);
                }
                Proj.Close();

                SqlDataReader Dep = AllDep.ExecuteReader();
                while (Dep.Read())
                {
                    string s = Dep.GetString(Dep.GetOrdinal("department"));
                    comboBox2.Items.Add(s);
                }
                Dep.Close();

                SqlDataReader myUsers = AllUsers.ExecuteReader();
                while (myUsers.Read())
                {
                    string s = myUsers.GetString(myUsers.GetOrdinal("Surename"));
                    comboBox3.Items.Add(s);
                }
                myUsers.Close();
                
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            finally
            {
                DbConn.Close();
            }
        }

        private void AutoSizeGridColumn(DataGridView myGrid)
        {
            myGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }
        private void GridHeaderName(DataGridView myGrid)
        {
            myGrid.Columns[0].HeaderText = "Название проекта";
            myGrid.Columns[1].HeaderText = "Текст задания";
            myGrid.Columns[2].HeaderText = "Тип работы";
            myGrid.Columns[3].HeaderText = "ФИО";
            myGrid.Columns[4].HeaderText = "Затраченное время";
            myGrid.Columns[5].HeaderText = "Дата заполнения";
            myGrid.Columns[6].HeaderText = "Коментарий";
            myGrid.Columns[7].HeaderText = "Процент выполнения";
        }

        private Exception ExportGrid(string Path, string NameFile)
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
    }
}

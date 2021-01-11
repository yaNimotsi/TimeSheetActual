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
        public string UId, DepId, TId, PId, conString, ProjName1, TypeBild1, fullPath1;
        public int level1, DeepId;
        public Form5(string UId, string DepId, string conString)
        {
            this.UId = UId;
            this.DepId = DepId;
            this.conString = conString;
            InitializeComponent();
        }
        
        private void comboBox2_DropDownClosed(object sender, EventArgs e)
        {
            if (comboBox2.Text.Length != 0)
            {
                SqlConnection DbConn = new SqlConnection(conString);
                SqlCommand DbComand = new SqlCommand();
                SqlCommand EnteredDepp = new SqlCommand();

                comboBox3.Items.Clear();
                string DepName = comboBox2.Text.ToString();

                SqlParameter depName = new SqlParameter("@DepName", DepName);
                SqlParameter depName2 = new SqlParameter("@DepName2", DepName);

                DbComand.CommandText = "SELECT dbo.Users.Surename"
                +" FROM            dbo.Users INNER JOIN"
                         +" dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                +" WHERE(dbo.Otdeli.department = @DepName)"
                +" ORDER BY dbo.Users.Surename";
                DbComand.Parameters.Add(depName);
                DbComand.Connection = DbConn;


                EnteredDepp.CommandText = "SELECT dbo.Otdeli.Id_department"
                + " FROM            dbo.Otdeli"
                + " WHERE(dbo.Otdeli.department = @DepName2)";
                EnteredDepp.Parameters.Add(depName2);
                EnteredDepp.Connection = DbConn;
                try
                {
                    comboBox3.Items.Clear();
                    DbConn.Open();
                    SqlDataReader DbReader = DbComand.ExecuteReader();
                    while (DbReader.Read())
                    {
                        string s = DbReader.GetString(DbReader.GetOrdinal("Surename"));
                        comboBox3.Items.Add(s);
                    }
                    DbReader.Close();

                    SqlDataReader DeppReader = EnteredDepp.ExecuteReader();
                    if (DeppReader.Read())
                    {
                        string s = DeppReader[0].ToString();
                        DepId = s;
                        DeepId = Convert.ToInt32(DepId);
                    }
                    DeppReader.Close();
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
            comboBox2.Items.Clear();
            comboBox3.Items.Clear();

            ProjName1 = "";
            TypeBild1 = "";
            fullPath1 = "";

            SqlConnection conn = new SqlConnection(conString);
            SqlCommand AllEntered = new SqlCommand();
            AllEntered.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                + " dbo.Report.Comment, dbo.Report.CountSheet, dbo.Otdeli.department"
                + " FROM dbo.Report INNER JOIN"
                         + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                         + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                         + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                         + " WHERE        (dbo.Report.DateEntered > @First)";

            DateTime myDate = DateTime.Today;
            DateTime first = new DateTime(myDate.Year, myDate.Month, 1);

            SqlParameter FirstDay = new SqlParameter("@First", first);
            AllEntered.Parameters.Add(FirstDay);
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
            treeView1.PathSeparator = "/";
            TreeNode treeNode = treeView1.SelectedNode;

            string NameProj;
            int level = level1;
            string fullPath = fullPath1;

            if (ProjName1 == null) NameProj = "";
            else NameProj = ProjName1;

            string TipeBild = "";
            if (TypeBild1 == null) TipeBild = "";
            else TipeBild = TypeBild1;
            
            string DName = comboBox2.Text.ToString();
            string UName = comboBox3.Text.ToString();
            DateTime DStart = Convert.ToDateTime(dateTimePicker1.Text.ToString());
            DateTime DEnd = Convert.ToDateTime(dateTimePicker2.Text.ToString());

            SqlConnection conn = new SqlConnection(conString);
            SqlCommand filterComand = new SqlCommand();

            SqlParameter ProjName = new SqlParameter("@ProjName", NameProj);
            SqlParameter deepId = new SqlParameter("@DeepId", DeepId);
            SqlParameter UserName = new SqlParameter("@UserName", UName);
            SqlParameter DateStart = new SqlParameter("@DStart", DStart);
            SqlParameter DateEnd = new SqlParameter("@DEnd", DEnd);
            SqlParameter FullPath = new SqlParameter("@fullPath", fullPath);

            if (NameProj.Length != 0)
            {
                if (DName.Length != 0)
                {
                    if (UName.Length != 0)
                    {
                        if (checkBox1.Checked == true)
                        {
                            if (checkBox2.Checked == true)//все
                            {
                                if (level > 1)
                                {
                                    filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                    + "dbo.Report.Comment, dbo.Report.CountSheet"
                                    + " FROM dbo.Report INNER JOIN"
                                    + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                    + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                    + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                                    + " WHERE        (dbo.Report.PuthToNode = @fullPath) AND(dbo.Users.department = @DeepId) AND (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered >=@DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                                    + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                                }
                                else if (level == 1)
                                {
                                    filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                    + "dbo.Report.Comment, dbo.Report.CountSheet"
                                    + " FROM dbo.Report INNER JOIN"
                                    + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                    + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                    + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                                    + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Users.department = @DeepId) AND (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered >=@DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                                    + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                                }

                                filterComand.Parameters.Add(ProjName);
                                filterComand.Parameters.Add(FullPath);
                                filterComand.Parameters.Add(deepId);
                                filterComand.Parameters.Add(UserName);
                                filterComand.Parameters.Add(DateStart);
                                filterComand.Parameters.Add(DateEnd);
                                

                                UpdateGrid(filterComand);
                            }
                            else//Все кроме финиша
                            {
                                if (level > 1)
                                {
                                    filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                                + " WHERE        (dbo.Report.PuthToNode = @fullPath) AND(dbo.Users.department = @DeepId) AND (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered >=@DStart)"
                                + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                                }
                                else if (level == 1)
                                {
                                    filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                                + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Users.department = @DeepId) AND (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered >=@DStart)"
                                + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                                }

                                filterComand.Parameters.Add(ProjName);
                                filterComand.Parameters.Add(FullPath);
                                filterComand.Parameters.Add(deepId);
                                filterComand.Parameters.Add(UserName);
                                filterComand.Parameters.Add(DateStart);

                                UpdateGrid(filterComand);
                            }
                        }
                        else if (checkBox2.Checked == true)//Все кроме старта
                        {
                            if (level > 1)
                            {
                                filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                    + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                    + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                    + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                                + " WHERE        (dbo.Report.PuthToNode = @fullPath) AND(dbo.Users.department = @DeepId) AND (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered <= @DEnd)"
                                + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }
                            else if (level == 1)
                            {
                                filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                    + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                    + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                    + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                                + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Users.department = @DeepId) AND (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered <= @DEnd)"
                                + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }
                            
                            filterComand.Parameters.Add(ProjName);
                            filterComand.Parameters.Add(FullPath);
                            filterComand.Parameters.Add(deepId);
                            filterComand.Parameters.Add(UserName);
                            filterComand.Parameters.Add(DateEnd);

                            UpdateGrid(filterComand);
                        }
                        else////Проект, отдел и пользователь
                        {
                            if (level > 1)
                            {
                                filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE        (dbo.Report.PuthToNode = @fullPath) AND(dbo.Users.department = @DeepId) AND (dbo.Users.Surename = @UserName)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }
                            else if (level == 1)
                            {
                                filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Users.department = @DeepId) AND (dbo.Users.Surename = @UserName)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }

                            filterComand.Parameters.Add(ProjName);
                            filterComand.Parameters.Add(FullPath);
                            filterComand.Parameters.Add(deepId);
                            filterComand.Parameters.Add(UserName);

                            UpdateGrid(filterComand);
                        }
                    }
                    else if (checkBox1.Checked == true) //Проект, отдел и старт
                    {
                        if (level > 1)
                        {
                            filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                        + " WHERE        (dbo.Report.PuthToNode = @fullPath) AND(dbo.Users.department = @DeepId) AND(dbo.Report.DateEntered >=@DStart)"
                        + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }
                        else if (level == 1)
                        {
                            filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                        + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Users.department = @DeepId) AND(dbo.Report.DateEntered >=@DStart)"
                        + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }

                        filterComand.Parameters.Add(ProjName);
                        filterComand.Parameters.Add(FullPath);
                        filterComand.Parameters.Add(deepId);
                        filterComand.Parameters.Add(DateStart);

                        UpdateGrid(filterComand);
                    }
                    else if (checkBox2.Checked == true)//Проект, отдел и финиш
                    {
                        if (level > 1)
                        {
                            filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                        + " WHERE        (dbo.Report.PuthToNode = @fullPath) AND(dbo.Users.department = @DeepId) AND(dbo.Report.DateEntered <= @DateEnd)"
                        + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }
                        else if (level == 1)
                        {
                            filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                        + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Users.department = @DeepId) AND(dbo.Report.DateEntered <= @DateEnd)"
                        + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }

                        filterComand.Parameters.Add(ProjName);
                        filterComand.Parameters.Add(FullPath);
                        filterComand.Parameters.Add(deepId);
                        filterComand.Parameters.Add(DateEnd);

                        UpdateGrid(filterComand);
                    }
                    else//Проект и отдел
                    {
                        if (level > 1)
                        {
                            filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                        + " WHERE        (dbo.Report.PuthToNode = @fullPath) AND(dbo.Users.department = @DeepId)"
                        + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }
                        else if (level == 1)
                        {
                            filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                        + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Users.department = @DeepId)"
                        + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }

                        filterComand.Parameters.Add(ProjName);
                        filterComand.Parameters.Add(FullPath);
                        filterComand.Parameters.Add(deepId);

                        UpdateGrid(filterComand);
                    }
                }
                else if (comboBox3.Text.Length != 0)//проект и пользователь
                {
                    if (checkBox1.Checked == true)
                    {
                        if (checkBox2.Checked == true)//проект, пользователь, страт и финиш
                        {
                            if (level > 1)
                            {
                                filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE        (dbo.Report.PuthToNode = @fullPath) AND (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered >=@DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }
                            else if (level == 1)
                            {
                                filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered >=@DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }

                            filterComand.Parameters.Add(ProjName);
                            filterComand.Parameters.Add(FullPath);
                            filterComand.Parameters.Add(UserName);
                            filterComand.Parameters.Add(DateStart);
                            filterComand.Parameters.Add(DateEnd);

                            UpdateGrid(filterComand);
                        }
                        else//проект, пользователь и страт
                        {
                            if (level > 1)
                            {
                                filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE        (dbo.Report.PuthToNode = @fullPath) AND (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered >=@DStart)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }
                            else if (level == 1)
                            {
                                filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered >=@DStart)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }

                            filterComand.Parameters.Add(ProjName);
                            filterComand.Parameters.Add(FullPath);
                            filterComand.Parameters.Add(UserName);
                            filterComand.Parameters.Add(DateStart);

                            UpdateGrid(filterComand);
                        }
                    }
                    else if (checkBox2.Checked == true)//проект, пользователь и финиш
                    {
                        if (level > 1)
                        {
                            filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                        + " WHERE        (dbo.Report.PuthToNode = @fullPath) AND (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered <= @DEnd)"
                        + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }
                        else if (level == 1)
                        {
                            filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                        + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered <= @DEnd)"
                        + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }

                        filterComand.Parameters.Add(ProjName);
                        filterComand.Parameters.Add(FullPath);
                        filterComand.Parameters.Add(UserName);
                        filterComand.Parameters.Add(DateEnd);

                        UpdateGrid(filterComand);
                    }
                    else
                    {
                        if (level > 1)
                        {
                            filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                        + " WHERE        (dbo.Report.PuthToNode = @fullPath) AND(dbo.Users.Surename = @UserName)"
                        + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }
                        else if (level == 1)
                        {
                            filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                        + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Users.Surename = @UserName)"
                        + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }

                        filterComand.Parameters.Add(ProjName);
                        filterComand.Parameters.Add(FullPath);
                        filterComand.Parameters.Add(UserName);

                        UpdateGrid(filterComand);
                    }
                }
                else if (checkBox1.Checked == true)//проект и старт
                {
                    if (checkBox2.Checked == true)
                    {
                        if (level > 1)
                        {
                            filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                         + " WHERE        (dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered >=@DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                         + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }
                        else if (level == 1)
                        {
                            filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                         + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Report.DateEntered >=@DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                         + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }

                        filterComand.Parameters.Add(ProjName);
                        filterComand.Parameters.Add(FullPath);
                        filterComand.Parameters.Add(DateStart);
                        filterComand.Parameters.Add(DateEnd);

                        UpdateGrid(filterComand);
                    }
                    else
                    {
                        if (level > 1)
                        {
                            filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                         + " WHERE        (dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered > @DStart)"
                         + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }
                        else if (level == 1)
                        {
                            filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                         + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Report.DateEntered > @DStart)"
                         + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }

                        filterComand.Parameters.Add(ProjName);
                        filterComand.Parameters.Add(FullPath);
                        filterComand.Parameters.Add(DateStart);

                        UpdateGrid(filterComand);
                    }
                }
                else if (checkBox2.Checked == true)//Проект и финиш
                {
                    if (level > 1)
                    {
                        filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                        + "dbo.Report.Comment, dbo.Report.CountSheet"
                        + " FROM dbo.Report INNER JOIN"
                        + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                        + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                        + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                    + " WHERE        (dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered <= @DEnd)"
                    + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    }
                    else if (level == 1)
                    {
                        filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                        + "dbo.Report.Comment, dbo.Report.CountSheet"
                        + " FROM dbo.Report INNER JOIN"
                        + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                        + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                        + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                    + " WHERE        (dbo.Proekti.Name_Project = @ProjName) AND(dbo.Report.DateEntered <= @DEnd)"
                    + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    }


                    filterComand.Parameters.Add(ProjName);
                    filterComand.Parameters.Add(FullPath);
                    filterComand.Parameters.Add(DateEnd);

                    UpdateGrid(filterComand);
                }
                else//Только проект
                {
                    if (level > 1)
                    {
                        filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                        + "dbo.Report.Comment, dbo.Report.CountSheet"
                        + " FROM dbo.Report INNER JOIN"
                        + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                        + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                        + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                    + " WHERE        (dbo.Report.PuthToNode = @fullPath)"
                    + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    }
                    else if (level == 1)
                    {
                        filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                        + "dbo.Report.Comment, dbo.Report.CountSheet"
                        + " FROM dbo.Report INNER JOIN"
                        + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                        + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                        + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                    + " WHERE        (dbo.Proekti.Name_Project = @ProjName)"
                    + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    }

                    filterComand.Parameters.Add(ProjName);
                    filterComand.Parameters.Add(FullPath);

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
                            filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                             + " WHERE        (dbo.Users.department = @DeepId) AND (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered >=@DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                             + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";


                            filterComand.Parameters.Add(deepId);
                            filterComand.Parameters.Add(UserName);
                            filterComand.Parameters.Add(DateStart);
                            filterComand.Parameters.Add(DateEnd);

                            UpdateGrid(filterComand);
                        }
                        else//234
                        {
                            filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE        (dbo.Users.department = @DeepId) AND (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered >=@DStart)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";


                            filterComand.Parameters.Add(deepId);
                            filterComand.Parameters.Add(UserName);
                            filterComand.Parameters.Add(DateStart);

                            UpdateGrid(filterComand);
                        }
                    }
                    else if (checkBox2.Checked == true)//235
                    {
                        filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                        + " WHERE        (dbo.Users.department = @DeepId) AND (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered <= @DEnd)"
                        + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";


                        filterComand.Parameters.Add(deepId);
                        filterComand.Parameters.Add(UserName);
                        filterComand.Parameters.Add(DateEnd);

                        UpdateGrid(filterComand);
                    }
                    else//23
                    {
                        filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                        + " WHERE        (dbo.Users.department = @DeepId) AND (dbo.Users.Surename = @UserName)"
                        + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";


                        filterComand.Parameters.Add(deepId);
                        filterComand.Parameters.Add(UserName);

                        UpdateGrid(filterComand);
                    }
                }
                else if (checkBox1.Checked == true)
                {
                    if (checkBox2.Checked == true)//245
                    {
                        filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                        + " WHERE        (dbo.Users.department = @DeepId) AND(dbo.Report.DateEntered >=@DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                        + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";


                        filterComand.Parameters.Add(deepId);
                        filterComand.Parameters.Add(DateStart);
                        filterComand.Parameters.Add(DateEnd);

                        UpdateGrid(filterComand);
                    }
                    else//24
                    {
                        filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                        + " WHERE        (dbo.Users.department = @DeepId) AND(dbo.Report.DateEntered >=@DStart)"
                        + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";


                        filterComand.Parameters.Add(deepId);
                        filterComand.Parameters.Add(DateStart);

                        UpdateGrid(filterComand);
                    }
                }
                else if (checkBox2.Checked == true)//25
                {
                    filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                        + "dbo.Report.Comment, dbo.Report.CountSheet"
                        + " FROM dbo.Report INNER JOIN"
                        + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                        + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                        + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                    + " WHERE        (dbo.Users.department = @DeepId) AND(dbo.Report.DateEntered <= @DEnd)"
                    + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";


                    filterComand.Parameters.Add(deepId);
                    filterComand.Parameters.Add(DateEnd);

                    UpdateGrid(filterComand);
                }
                else//2
                {
                    filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                        + "dbo.Report.Comment, dbo.Report.CountSheet"
                        + " FROM dbo.Report INNER JOIN"
                        + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                        + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                        + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                    + " WHERE        (dbo.Users.department = @DeepId)"
                    + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";


                    filterComand.Parameters.Add(deepId);

                    UpdateGrid(filterComand);
                }
                    
            }
            else if (comboBox3.Text.Length != 0)
            {
                if (checkBox1.Checked == true)
                {
                    if (checkBox2.Checked == true)//345
                    {
                        filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                         + " WHERE        (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered >=@DStart) AND(dbo.Report.DateEntered <= @DEnd)"
                         + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";


                        filterComand.Parameters.Add(UserName);
                        filterComand.Parameters.Add(DateStart);
                        filterComand.Parameters.Add(DateEnd);

                        UpdateGrid(filterComand);
                    }
                    else//34
                    {
                        filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                            + " FROM dbo.Report INNER JOIN"
                            + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                            + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                            + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                        + " WHERE        (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered >=@DStart)"
                        + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";


                        filterComand.Parameters.Add(UserName);
                        filterComand.Parameters.Add(DateStart);

                        UpdateGrid(filterComand);
                    }
                }
                else if (checkBox2.Checked == true)//35
                {
                    filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                        + "dbo.Report.Comment, dbo.Report.CountSheet"
                        + " FROM dbo.Report INNER JOIN"
                        + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                        + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                        + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                    + " WHERE        (dbo.Users.Surename = @UserName) AND(dbo.Report.DateEntered <= @DEnd)"
                    + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";


                    filterComand.Parameters.Add(UserName);
                    filterComand.Parameters.Add(DateEnd);

                    UpdateGrid(filterComand);
                }
                else//3
                {
                    filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                        + "dbo.Report.Comment, dbo.Report.CountSheet"
                        + " FROM dbo.Report INNER JOIN"
                        + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                        + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                        + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                    + " WHERE        (dbo.Users.Surename = @UserName)"
                    + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";


                    filterComand.Parameters.Add(UserName);

                    UpdateGrid(filterComand);
                }
            }
            else if (checkBox1.Checked == true)
            {
                if (checkBox2.Checked == true)//45
                {
                    filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                        + " dbo.Report.Comment, dbo.Report.CountSheet, dbo.Otdeli.department"
                        + " FROM dbo.Report INNER JOIN"
                        + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                        + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                        + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                    + " WHERE        (dbo.Report.DateEntered >=@DStart) AND(dbo.Report.DateEntered <= @DEnd)";
                    //+ " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    filterComand.Parameters.Add(DateStart);
                    filterComand.Parameters.Add(DateEnd);

                    UpdateGrid(filterComand);
                }
                else//4
                {
                    filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                        + "dbo.Report.Comment, dbo.Report.CountSheet"
                        + " FROM dbo.Report INNER JOIN"
                        + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                        + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                        + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                    + " WHERE        (dbo.Report.DateEntered >=@DStart)"
                    + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";


                    filterComand.Parameters.Add(DateStart);

                    UpdateGrid(filterComand);
                }
            }
            else if (checkBox2.Checked == true)//5
            {
                filterComand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                    + "dbo.Report.Comment, dbo.Report.CountSheet"
                    + " FROM dbo.Report INNER JOIN"
                    + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                    + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                    + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                + " WHERE        (dbo.Report.DateEntered <= @DEnd)"
                + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                filterComand.Parameters.Add(DateEnd);

                UpdateGrid(filterComand);
            }
            else
            {
                MessageBox.Show("Не выбран объект для фильтрации", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Form5_Load(object sender, EventArgs e)
        {

            DataGridView GridTab1 = dataGridView1;

            AutoSizeGridColumn(GridTab1);
            SqlCommand AllEntered = new SqlCommand();

            AllEntered.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                + " dbo.Report.Comment, dbo.Report.CountSheet, dbo.Otdeli.department"
                + " FROM dbo.Report INNER JOIN"
                         + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                         + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                         + " dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
                         + " WHERE        (dbo.Report.DateEntered > @First)";
            
            DateTime myDate = DateTime.Today;
            DateTime first = new DateTime(myDate.Year, myDate.Month, 1);

            SqlParameter FirstDay = new SqlParameter("@First", first);
            AllEntered.Parameters.Add(FirstDay);

            dateTimePicker1.Value = myDate;
            dateTimePicker2.Value = myDate;

            checkBox1.Checked = false;
            checkBox2.Checked = false;

            dateTimePicker1.Enabled = false;
            dateTimePicker2.Enabled = false;
            
            UpdateGrid(AllEntered);
            GridHeaderName(GridTab1);
            DataToListBox();

            myTreeViewWork(treeView1);

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView1.AllowUserToResizeColumns = true;
            dataGridView1.RowHeadersVisible = false;
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
                AddNodesTree(FinnArr, RowCountList, i, treeView);
            }
            catch (Exception ex) { MessageBox.Show("" + ex); }

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
                //if (myArr[i, 0] != null)
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
            int[,] rez = new int[RowCount, 3];

            for (int i = 0; i < RowCount - 1; i++)
            {
                //if (myArr[i, 0] != null)
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
                finally { conn.Close(); }
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
            }
            else
            {
                ProjName1 = "";
                TypeBild1 = "";
                level1 = 0;
            }
        }

        public void DataToListBox()
        {
            comboBox2.Items.Clear();
            comboBox2.Text = "";

            comboBox3.Items.Clear();
            comboBox3.Text = "";

            SqlConnection DbConn = new SqlConnection(conString);
            SqlCommand AllDep = new SqlCommand();
            SqlCommand AllUsers = new SqlCommand();

            AllDep.CommandText = "SELECT        TOP (100) PERCENT department, Active"
                        +" FROM dbo.Otdeli"
                        +" WHERE(Active = 1)"
                        +" ORDER BY department";
            AllDep.Connection = DbConn;

            AllUsers.CommandText = "SELECT        TOP (100) PERCENT Surename"
                        +" FROM dbo.Users"
                        +" WHERE(Active = 1)"
                        +" ORDER BY Surename";
            AllUsers.Connection = DbConn;
            try
            {
                DbConn.Open();

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
            myGrid.Columns[1].HeaderText = "Название здания/сооружения";
            myGrid.Columns[2].HeaderText = "Тип работы";
            myGrid.Columns[3].HeaderText = "Раздел";
            myGrid.Columns[4].HeaderText = "ФИО исполнителя";
            myGrid.Columns[5].HeaderText = "Затраченое время";
            myGrid.Columns[6].HeaderText = "Дата заполнения отчетности";
            myGrid.Columns[7].HeaderText = "Коментарий";
            myGrid.Columns[8].HeaderText = "Кол-во выполненых листов";
            myGrid.Columns[9].HeaderText = "Отдел";
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
    }
}

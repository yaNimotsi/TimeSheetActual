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
using System.Security.Principal;
using System.IO;

namespace TimeSheet
{
    public partial class Form1 : Form
    {
        string IniPath = @"C:\Users\Public\TimeSheet.ini";
        static string connectionString = @"Data Source=vnipipt-s-sql03.vnipipt.ru\TIMESHEET;Initial Catalog=TIMESHEET;User ID=sa;Pwd=h)SFk@j2;";
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {   //Определяем имя учетной записи
            //Добавить в таблицу Users столбец с именами учетной записи, и по нему осуществлять процедуру логирования.
            string s = Environment.UserName;

            //Определние имени активной учетной записи
            //s = WindowsIdentity.GetCurrent().Name;

            if ((textBox1.Text.Length > 0) && (textBox2.Text.Length > 0))
            {
                if(checkBox1.Checked == true)
                {
                    MyIni myIni = new MyIni(IniPath);
                    myIni.IniWriteValue("Section1", "UserName", textBox1.Text.ToString());
                    myIni.IniWriteValue("Section2", "Pass", textBox2.Text.ToString());
                }
                
                SqlConnection SQLConn = new SqlConnection(connectionString);
                SqlCommand sqlCommand = new SqlCommand();
                SqlParameter UserName = new SqlParameter("@UName", textBox1.Text.ToString());
                SqlParameter UserPass = new SqlParameter("@UPass", textBox2.Text.ToString());

                sqlCommand.CommandText = "SELECT Id_users, Id_position, department"
                + " FROM Users"
                + " WHERE(Surename = @UName) AND([Password] = @UPass)";
                sqlCommand.Connection = SQLConn;
                sqlCommand.Parameters.Add(UserName);
                sqlCommand.Parameters.Add(UserPass);
                
                try
                {
                    SQLConn.Open();
                    SqlDataReader SQLDbReade = sqlCommand.ExecuteReader();

                    /*DbConn.Open();
                    OleDbDataReader DbReader = DbComand.ExecuteReader();*/
                    string UId, PosId, DepId = "";

                    if (SQLDbReade.HasRows)
                    {
                        if (checkBox1.Checked == true)
                        {
                            Properties.Settings.Default.UserName = textBox1.Text;
                            Properties.Settings.Default.pass = textBox2.Text;
                            Properties.Settings.Default.Save();
                        }
                        while (SQLDbReade.Read())
                        {
                            UId = SQLDbReade[0].ToString();
                            PosId = SQLDbReade[1].ToString();
                            DepId = SQLDbReade[2].ToString();

                            if (PosId.Length != 0)
                            {
                                int NomPosId = Convert.ToInt32(PosId);
                                int NomDepId = Convert.ToInt32(DepId);
                                //Начальник отдела
                                /*if ((NomPosId == 13) || (NomPosId == 15) || (NomPosId == 33) || (NomPosId == 35) || (NomPosId == 36) || (NomPosId == 38) || (NomPosId == 39) || ((NomPosId >= 49) &&(NomPosId <= 55)) || (NomPosId == 57) || (NomPosId == 58) || ((NomPosId >= 60) && (NomPosId <= 62)))
                                {
                                    Form3 MainForm = new Form3(UId, DepId, connectionString);
                                    MainForm.Show();
                                    this.Hide();
                                    break;
                                }*/
                                //ГИП, Главный инженер или Исполнительный директор
                                /*else if (((NomPosId == 12) || (NomPosId == 14) || (NomPosId == 28)) || (NomPosId == 29) || (NomPosId == 31) || (NomPosId == 46))
                                {
                                    Form5 MainForm = new Form5(UId, DepId, connectionString);
                                    MainForm.Show();
                                    this.Hide();
                                    break;
                                }
                                //Исполнитель
                                //else*/
                                {
                                    Form2 MainForm = new Form2(UId, connectionString, DepId);
                                    MainForm.Show();
                                    this.Hide();
                                    break;
                                }
                            }
                        }
                        SQLDbReade.Close();
                    }
                    else
                    {
                        MessageBox.Show("Пара логин/пароль не найдена. Проверьте правильность ввода и повторите попытку");
                    }
                }
                catch (Exception ex) { MessageBox.Show("Произошла ошибка при обработке запроса к базе данных:" + ex); }
                finally { SQLConn.Close(); }
            }
            else MessageBox.Show("Необходимо заполнить поля логин/пароль", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //textBox1.Text = Properties.Settings.Default.UserName.ToString();
            //textBox2.Text = Properties.Settings.Default.pass.ToString();

            MyIni myIni = new MyIni(IniPath);

            if (myIni.IniReadValue("Section1", "UserName").Length!=0)
            {
                textBox1.Text = myIni.IniReadValue("Section1", "UserName");
            }
            if (myIni.IniReadValue("Section2", "Pass").Length != 0)
            {
                textBox2.Text = myIni.IniReadValue("Section2", "Pass");
            }
        }

        private Boolean HaveFile()
        {
            Boolean flag = false;
            string pathFile = @"C:\Users\Public\INIConfig.ini";

            if (File.Exists(pathFile)) flag = true;

            return flag;
        }
    }
}

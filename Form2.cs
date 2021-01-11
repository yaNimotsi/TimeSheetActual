using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TimeSheet
{
    public partial class Form2 : Form
    {
        public string UId,DepId, UIdT, myDate, myTime, PId, TId, conString, ProjName1, ProjName2, TypeBild1, TypeBild2, fullPath1, fullPath2;
        private int UserID, TimerVal, level1, level2;
        private string[,] enteredReportBefore;
        private TreeNode baseTreeNode;
        public DataTable Dt = new DataTable();
        public Form2(string UId, string conString, string DepId)
        {
            this.UId = UId;
            this.DepId = DepId;
            UserID = Convert.ToInt32(UId);
            this.conString = conString;
            InitializeComponent();
        }

#pragma warning disable IDE1006 // Стили именования
        private void button1_Click(object sender, EventArgs e)
#pragma warning restore IDE1006 // Стили именования
        {
            int level = level1;
            string fullPath = fullPath1;

            string PuthToNode = fullPath1;
            string NameSelProj = ProjName1;
            string TypeBild = TypeBild1;

            //if ((NameSelProj == "")||(NameSelProj == null) || (TypeBild == "") || (TypeBild  == null) || (NameSelProj == TypeBild)) MessageBox.Show("Не выбрана площадка или сооружение. Сделайте выбор и повторите попытку.","Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            if ((NameSelProj == "") || (NameSelProj == null) || (NameSelProj == TypeBild)) MessageBox.Show("Не выбрана площадка или сооружение. Сделайте выбор и повторите попытку.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else if ((comboBox1.Text.Length != 0) && (comboBox2.Text.Length != 0) && (textBox4.Text.Length != 0))
            {
                string WorkTimeVal = textBox4.Text.ToString().Replace('.', ',');
                float FWorkTimeVal = Convert.ToSingle(WorkTimeVal);
                
                DateTime dtime = DateTime.Now;

                string Section = comboBox2.Text.ToString();
                string WorkTime = textBox4.Text.ToString();
                string coment = richTextBox1.Text.ToString();
                string CountSheet = textBox1.Text.ToString();
                string TypeWork = comboBox1.Text.ToString();

                SqlConnection DbConn = new SqlConnection(conString);
                SqlCommand DbComandInsert = new SqlCommand
                {
                    CommandText = "INSERT INTO dbo.Report(Id_Project, Id_user, TypeWork, Section, TypeBild, Comment, CountSheet, TimeWork, DateEntered, TimeEntered1,PuthToNode)"
                + " VALUES"
                + " (@pid, @uid, @TWork, @Section, @TBild, @Coment, @cSheet, @WorkTime, @Dtime, @timeEntered, @PuthToNode)",

                    Connection = DbConn
                };

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

                    comboBox1.SelectedIndex = -1;
                    comboBox2.SelectedIndex = -1;
                    richTextBox1.Text = "";
                    textBox1.Text = "";
                    textBox4.Text = "";
                }
                catch (Exception ex) { MessageBox.Show("Запись добавить не удалось!!! " + ex, "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                finally { DbConn.Close(); }
            }
            else MessageBox.Show("Не заполнен один из обязательных параметров! Заполните все и повторите попытку", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);

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
                while (sqlDataReader.Read())
                {
                    rez = sqlDataReader[0].ToString();
                }
            }
            catch (SqlException ex) { MessageBox.Show("Произошла ошибка выполнения команды: " + ex, "Ошибка выполнения сценария", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally { conn.Close(); }
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

        public void UpdateGridAbsence()
        {

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
            /*sqlCommand.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Otdeli.department, dbo.Report.Comment, "
                         +" dbo.Report.CountSheet, dbo.Report.DateEntered"
+" FROM            dbo.Report INNER JOIN"
                         +" dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                         +" dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                         +" dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
        + " WHERE(dbo.Users.Id_users = @userID)";

            SqlParameter userID = new SqlParameter("@userID", UserID);
            sqlCommand.Parameters.Add(userID);

            UpdateGrid(sqlCommand);*/

            SqlParameter userID2 = new SqlParameter("@UId", UserID);

            EnteredVal.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                + " dbo.Report.Comment, dbo.Report.CountSheet, dbo.Otdeli.department"
                + " FROM dbo.Report INNER JOIN"
                         +" dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                         +" dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                         +" dbo.Otdeli ON dbo.Users.department = dbo.Otdeli.Id_department"
            + " WHERE(dbo.Users.Id_users = @UId)";
            //+ " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";
            
            EnteredVal.Parameters.Add(userID2);
            EnteredVal.Connection = sqldbConnection;

            UpdateGrid(EnteredVal, GridTab2);

            DataToListBox();

            FindTopUsed();//Определние топ используемых
            BeforeEnteredReport();

            GridHeaderName(GridTab2);

            DateTime myDate = DateTime.Today;
            
            dateTimePicker5.Value = myDate;
            dateTimePicker6.Value = myDate;
            
            checkBox6.Checked = false;
            checkBox7.Checked = false;
            
            dateTimePicker5.Enabled = false;
            dateTimePicker6.Enabled = false;

            MyTreeViewWork(treeView1);
            MyTreeViewWork(treeView2);


            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView2.AllowUserToResizeColumns = true;
            dataGridView2.RowHeadersVisible = false;

            baseTreeNode = treeView1.Nodes[0];
        }

        public void FindTopUsed()
        {
            DateTime dStart = DateTime.Now.AddDays(-30);
            SqlConnection conn = new SqlConnection(conString);
            SqlCommand tTypeWork = new SqlCommand
            {
                Connection = conn,
                CommandText = "SELECT TOP (3) TypeWork, COUNT(1) AS cnt"
            + " FROM dbo.Report"
            + " WHERE(Id_user = @uId) AND(DateEntered > @sDate)"
            + " GROUP BY TypeWork"
            + " ORDER BY cnt DESC"
            };

            SqlParameter dateStart = new SqlParameter("@sDate", dStart);
            SqlParameter uId = new SqlParameter("@uId", UId);
            tTypeWork.Parameters.Add(dateStart);
            tTypeWork.Parameters.Add(uId);


            SqlCommand tSection = new SqlCommand
            {
                Connection = conn,
                CommandText = "SELECT TOP (3) Section, COUNT(1) AS cnt"
            + " FROM dbo.Report"
            + " WHERE(Id_user = @uId1) AND(DateEntered > @sDate1)"
            + " GROUP BY Section"
            + " ORDER BY cnt DESC"
            };

            SqlParameter dateStart1= new SqlParameter("@sDate1", dStart);
            SqlParameter uId1 = new SqlParameter("@uId1", UId);
            tSection.Parameters.Add(dateStart1);
            tSection.Parameters.Add(uId1);

            try
            {
                conn.Open();

                SqlDataReader sqlDRTypeWork = tTypeWork.ExecuteReader();
                while(sqlDRTypeWork.Read())
                {
                    listBox1.Items.Add(sqlDRTypeWork[0].ToString());
                }
                sqlDRTypeWork.Close();

                SqlDataReader sqlDRSection = tSection.ExecuteReader();
                while (sqlDRSection.Read())
                {
                    listBox2.Items.Add(sqlDRSection[0].ToString());
                }
                sqlDRSection.Close();
            }
            catch (Exception) { }
            finally { conn.Close(); }
        }
        
        public void BeforeEnteredReport()
        {
            DateTime dateTime = DateTime.Now;
            dateTime = dateTime.AddDays(-7);

            SqlConnection conn = new SqlConnection(conString);

            SqlCommand lastReportsCountRow = new SqlCommand
            {
                Connection = conn,
                CommandText = "SELECT COUNT(100) AS cnt"
            + " FROM dbo.Report"
            + " WHERE(DateEntered > @Befor5Day1) AND(Id_user = @UserId1)"
            };

            SqlParameter Befor5Day1 = new SqlParameter("@Befor5Day1", dateTime);
            SqlParameter UserId1 = new SqlParameter("@UserId1", UId);

            lastReportsCountRow.Parameters.Add(Befor5Day1);
            lastReportsCountRow.Parameters.Add(UserId1);

            SqlCommand lastReports = new SqlCommand
            {
                Connection = conn,

                CommandText = "SELECT Id_Entered, Id_Project, Id_user, TypeWork, Section, TypeBild, Comment, CountSheet, TimeWork, DateEntered, PuthToNode, TimeEntered1"
            + " FROM dbo.Report"
            + " WHERE(DateEntered > @Befor5Day) AND(Id_user = @UserId)"
            };

            SqlParameter Befor5Day = new SqlParameter("@Befor5Day", dateTime);
            SqlParameter UserId = new SqlParameter("@UserId", UId);

            lastReports.Parameters.Add(Befor5Day);
            lastReports.Parameters.Add(UserId);

            int countReturnRow = 0;
            
            try
             {
                conn.Open();

                SqlDataReader lastReportsCountRowDR = lastReportsCountRow.ExecuteReader();
                if (lastReportsCountRowDR.Read())
                {
                    countReturnRow = Convert.ToInt32(lastReportsCountRowDR[0].ToString());
                }
                lastReportsCountRowDR.Close();

                if (countReturnRow > 0)
                {
                    string[,] arrReturnData = new string[countReturnRow, 12];
                    int i = 0;

                    SqlDataReader lastReportsDR = lastReports.ExecuteReader();
                    while (lastReportsDR.Read())
                    {
                        for (int j = 0; j < 12; j++)
                        {
                            arrReturnData[i, j] = lastReportsDR[j].ToString();
                        }
                        i++;
                    }
                    lastReportsDR.Close();
                    AddReportToComboBox(arrReturnData, countReturnRow);
                    enteredReportBefore = arrReturnData;
                }
            }
            catch (Exception) { }
            finally { conn.Close(); }
        }

        public void AddReportToComboBox(string[,] arrData, int countReturnRow)
        {
            string s = "";
            if (countReturnRow > 0)
            {
                for (int i = 0; i < countReturnRow; i++)
                {
                    s = arrData[i, 10].ToString() + " " +  arrData[i, 9].ToString().Substring(0,10);
                    comboBox3.Items.Add(s);
                }
            }
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
                finally { conn.Close(); }
            }
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
            SqlCommand sqlCommand = new SqlCommand
            {
                CommandText = myQuary,
                Connection = sqldbConnection
            };
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
                    long CRow = dt.Rows.Count;
                    da.Fill(dt);

                    GridName.DataSource = dt; //имя грида
                    conn.Close();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message);
                    return;   //или не нужно
                }
                finally { conn.Close(); }
            }
        }
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
            SqlCommand DbComand = new SqlCommand
            {
                CommandText = Quary,
                Connection = DbConn
            };
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
                SqlCommand DbSomeID = new SqlCommand
                {
                    CommandText = myQuary,
                    Connection = DbConn
                };
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

            ProjName2 = "";
            TypeBild2 = "";
            fullPath2 = "";

            treeView2.CollapseAll();
            
            checkBox6.Checked = false;
            checkBox7.Checked = false;

            SqlConnection conn = new SqlConnection(conString);
            SqlCommand EnteredVal = new SqlCommand();
            SqlParameter userID2 = new SqlParameter("@UId", UserID);

            EnteredVal.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
            + " WHERE(dbo.Users.Id_users = @UId)"
            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";
            
            DataGridView myObj = dataGridView2;

            EnteredVal.Parameters.Add(userID2);
            EnteredVal.Connection = conn;

            UpdateGrid(EnteredVal, myObj);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(conString);
            SqlCommand EnteredFilter = new SqlCommand
            {
                Connection = conn
            };

            DateTime DStart = Convert.ToDateTime(dateTimePicker5.Value.ToString());
            DateTime DEnd = Convert.ToDateTime(dateTimePicker6.Value.ToString());

            TreeNode treeNode = treeView2.SelectedNode;
            
            string NameProj;
            int level = level2;
            string fullPath = fullPath2;

            if (ProjName2 == null) NameProj = "";
            else NameProj = ProjName2;

            string TipeBild = "";
            if (TypeBild2 == null) TipeBild = "";
            else TipeBild = TypeBild2;

            SqlParameter uID = new SqlParameter("@UId", UId);
            SqlParameter dStart = new SqlParameter("@DSTart", DStart);
            SqlParameter dEnd = new SqlParameter("@DEnd", DEnd);
            SqlParameter nameProj = new SqlParameter("@NameProj", NameProj);
            SqlParameter tipeBild = new SqlParameter("@TypeBild", TipeBild);
            SqlParameter FullPath = new SqlParameter("@fullPath", fullPath);

            DataGridView myGrid = dataGridView2;

            if (NameProj.Length != 0)
            {
                if (TipeBild.Length != 0)
                {
                    if (checkBox6.Checked == true)
                    {
                        if (checkBox7.Checked == true)//1234
                        {
                            if (level > 1)
                            {
                                EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered >= @DSTart) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }
                            else if (level == 1)
                            {
                                EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                            + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Proekti.Name_Project = @NameProj) AND(dbo.Report.DateEntered >= @DSTart) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }
                            EnteredFilter.Parameters.Add(uID);
                            EnteredFilter.Parameters.Add(dStart);
                            EnteredFilter.Parameters.Add(dEnd);
                            EnteredFilter.Parameters.Add(nameProj);
                            EnteredFilter.Parameters.Add(tipeBild);
                            EnteredFilter.Parameters.Add(FullPath);
                            
                            UpdateGrid(EnteredFilter, myGrid);
                        }
                        else//123
                        {
                            if (level > 1)
                            {
                                EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered >= @DSTart)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }
                            else if (level == 1)
                            {
                                EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                                + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Proekti.Name_Project = @NameProj) AND(dbo.Report.DateEntered >= @DSTart)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                            }

                            EnteredFilter.Parameters.Add(uID);
                            EnteredFilter.Parameters.Add(dStart);
                            EnteredFilter.Parameters.Add(nameProj);
                            EnteredFilter.Parameters.Add(tipeBild);
                            EnteredFilter.Parameters.Add(FullPath);

                            UpdateGrid(EnteredFilter, myGrid);
                        }
                    }
                    else if (checkBox7.Checked == true)//124
                    {
                        if (level > 1)
                        {
                            EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }
                        else if (level == 1)
                        {
                            EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Proekti.Name_Project = @NameProj) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }

                        EnteredFilter.Parameters.Add(uID);
                        EnteredFilter.Parameters.Add(dEnd);
                        EnteredFilter.Parameters.Add(nameProj);
                        EnteredFilter.Parameters.Add(tipeBild);
                        EnteredFilter.Parameters.Add(FullPath);

                        UpdateGrid(EnteredFilter, myGrid);
                    }
                    else//12
                    {
                        if (level > 1)
                        {
                            EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.PuthToNode = @fullPath)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }
                        else if (level == 1)
                        {
                            EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Proekti.Name_Project = @NameProj)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }

                        EnteredFilter.Parameters.Add(uID);
                        EnteredFilter.Parameters.Add(nameProj);
                        EnteredFilter.Parameters.Add(tipeBild);
                        EnteredFilter.Parameters.Add(FullPath);

                        UpdateGrid(EnteredFilter, myGrid);
                    }
                }
                else if (checkBox6.Checked == true)
                {
                    if (checkBox7.Checked == true)//134
                    {
                        if (level > 1)
                        {
                            EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered >= @DSTart) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }
                        else if (level == 1)
                        {
                            EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Proekti.Name_Project = @NameProj) AND(dbo.Report.DateEntered >= @DSTart) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }

                        EnteredFilter.Parameters.Add(uID);
                        EnteredFilter.Parameters.Add(dStart);
                        EnteredFilter.Parameters.Add(dEnd);
                        EnteredFilter.Parameters.Add(nameProj);
                        EnteredFilter.Parameters.Add(FullPath);

                        UpdateGrid(EnteredFilter, myGrid);
                    }
                    else//13
                    {
                        if (level > 1)
                        {
                            EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                        + "dbo.Report.Comment, dbo.Report.CountSheet"
                               + " FROM dbo.Report INNER JOIN"
                               + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                               + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                               + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                           + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.PuthToNode = @fullPath)  AND(dbo.Report.DateEntered >= @DSTart)"
                           + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }
                        else if (level == 1)
                        {
                            EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                        + "dbo.Report.Comment, dbo.Report.CountSheet"
                               + " FROM dbo.Report INNER JOIN"
                               + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                               + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                               + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                           + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Proekti.Name_Project = @NameProj)  AND(dbo.Report.DateEntered >= @DSTart)"
                           + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        }

                        EnteredFilter.Parameters.Add(uID);
                        EnteredFilter.Parameters.Add(dStart);
                        EnteredFilter.Parameters.Add(nameProj);
                        EnteredFilter.Parameters.Add(FullPath);

                        UpdateGrid(EnteredFilter, myGrid);
                    }
                }
                else if (checkBox7.Checked == true)//14
                {
                    if (level > 1)
                    {
                        EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    }
                    else if (level == 1)
                    {
                        EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Proekti.Name_Project = @NameProj) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    }

                    EnteredFilter.Parameters.Add(uID);
                    EnteredFilter.Parameters.Add(dEnd);
                    EnteredFilter.Parameters.Add(nameProj);
                    EnteredFilter.Parameters.Add(FullPath);

                    UpdateGrid(EnteredFilter, myGrid);
                }
                else//1
                {
                    if (level > 1)
                    {
                        EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.PuthToNode = @fullPath)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    }
                    else if (level == 1)
                    {
                        EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Proekti.Name_Project = @NameProj)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    }

                    EnteredFilter.Parameters.Add(uID);
                    EnteredFilter.Parameters.Add(nameProj);
                    EnteredFilter.Parameters.Add(FullPath);

                    UpdateGrid(EnteredFilter, myGrid);
                }
            }
            else if (TipeBild.Length != 0)
            {
                if (checkBox6.Checked == true)
                {
                    if (checkBox7.Checked == true)//234
                    {
                        EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered >= @DSTart) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        EnteredFilter.Parameters.Add(uID);
                        EnteredFilter.Parameters.Add(dStart);
                        EnteredFilter.Parameters.Add(dEnd);
                        EnteredFilter.Parameters.Add(tipeBild);
                        EnteredFilter.Parameters.Add(FullPath);

                        UpdateGrid(EnteredFilter, myGrid);
                    }
                    else//23
                    {
                        EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId)  AND(dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered >= @DSTart))"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                        EnteredFilter.Parameters.Add(uID);
                        EnteredFilter.Parameters.Add(dStart);
                        EnteredFilter.Parameters.Add(tipeBild);
                        EnteredFilter.Parameters.Add(FullPath);

                        UpdateGrid(EnteredFilter, myGrid);
                    }
                }
                else if (checkBox7.Checked == true)//24
                {
                    EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.PuthToNode = @fullPath) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    EnteredFilter.Parameters.Add(uID);
                    EnteredFilter.Parameters.Add(dEnd);
                    EnteredFilter.Parameters.Add(tipeBild);
                    EnteredFilter.Parameters.Add(FullPath);

                    UpdateGrid(EnteredFilter, myGrid);
                }
                else//2
                {
                    EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.PuthToNode = @fullPath)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    EnteredFilter.Parameters.Add(uID);
                    EnteredFilter.Parameters.Add(tipeBild);
                    EnteredFilter.Parameters.Add(FullPath);

                    UpdateGrid(EnteredFilter, myGrid);
                }
            }
            else if(checkBox6.Checked == true)
            {
                if (checkBox7.Checked == true)//34
                {
                    EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.DateEntered >= @DSTart) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    EnteredFilter.Parameters.Add(uID);
                    EnteredFilter.Parameters.Add(dStart);
                    EnteredFilter.Parameters.Add(dEnd);

                    UpdateGrid(EnteredFilter, myGrid);
                }
                else//3
                {
                    EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.DateEntered >= @DSTart)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

                    EnteredFilter.Parameters.Add(uID);
                    EnteredFilter.Parameters.Add(dStart);

                    UpdateGrid(EnteredFilter, myGrid);
                }
            }
            else if (checkBox7.Checked == true)//4
            {
                EnteredFilter.CommandText = "SELECT        dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, "
                         + "dbo.Report.Comment, dbo.Report.CountSheet"
                                + " FROM dbo.Report INNER JOIN"
                                + " dbo.Proekti ON dbo.Report.Id_Project = dbo.Proekti.Id_project INNER JOIN"
                                + " dbo.Users ON dbo.Report.Id_user = dbo.Users.Id_users INNER JOIN"
                                + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj"
                            + " WHERE(dbo.Users.Id_users = @UId) AND(dbo.Report.DateEntered <= @DEnd)"
                            + " GROUP BY dbo.Proekti.Name_Project, dbo.Report.TypeBild, dbo.Report.TypeWork, dbo.Report.Section, dbo.Users.Surename, dbo.Report.TimeWork, dbo.Report.DateEntered, dbo.Report.Comment, dbo.Report.CountSheet, dbo.Report.Id_Entered";

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
            string NameFile = "Выгрузка " + now.ToString("d");
            
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

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            treeView1.PathSeparator = "/";
            fullPath1 = treeView1.SelectedNode.FullPath.ToString();
            richTextBox2.Text = fullPath1;
            var arr = fullPath1.Split('/');
            if (arr.Length >= 2)
            {
                ProjName1 = arr[1];
                if (treeView1.SelectedNode.Level > 1)
                {
                    TypeBild1 = treeView1.SelectedNode.Text;
                }
                else TypeBild1 = "";
                //Вместо 5 строк выше
                //TypeBild1 = treeView1.SelectedNode.Text;

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

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //string selectVal = listBox1.SelectedIndex.ToString();
            if (listBox1.SelectedItem != null)
            {
                string selectVal = listBox1.SelectedItem.ToString();
                if (comboBox1.Items.IndexOf(selectVal) != -1)
                {
                    comboBox1.SelectedIndex = comboBox1.Items.IndexOf(selectVal); //(comboBox1.FindString(selectVal));
                }
            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox2.SelectedItem != null)
            {
                string selectVal = listBox2.SelectedItem.ToString();
                if (comboBox2.Items.IndexOf(selectVal) != -1)
                {
                    comboBox2.SelectedIndex = comboBox2.Items.IndexOf(selectVal); //(comboBox1.FindString(selectVal));
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex != -1)
            {
                SelectTreeViewNode();
                if (comboBox1.Items.IndexOf(label15.Text) > 0)
                {
                    comboBox1.SelectedIndex = comboBox1.Items.IndexOf(label15.Text);
                }

                if (comboBox2.Items.IndexOf(label16.Text) > 0)
                {
                    comboBox2.SelectedIndex = comboBox2.Items.IndexOf(label16.Text);
                }

                richTextBox1.Text = label17.Text;
                textBox1.Text = label18.Text;
                textBox4.Text = label19.Text;

                fullPath1 = enteredReportBefore[comboBox3.SelectedIndex, 10];

                richTextBox2.Text = fullPath1;

                string[] arrPath = fullPath1.Split('/');
                if (arrPath.Length > 0)
                {
                    ProjName1 = arrPath[1];
                    if (arrPath.Length > 2)
                    {
                        TypeBild1 = arrPath[arrPath.Length - 1];
                    }
                    else
                    {
                        TypeBild1 = "";
                    }
                }
            }            
        }

        private TreeNode SearchNode(string SearchText, TreeNode StartNode)
        {
            TreeNode node = treeView1.Nodes[0];

            while(StartNode != null)
            {
                if (StartNode.Text.ToLower().Contains(SearchText.ToLower()))
                {
                    node = StartNode;
                    break;
                }
                if (StartNode.Nodes.Count != 0)
                {
                    node = SearchNode(SearchText, StartNode.Nodes[0]);
                    if (node != null)
                        break;
                }
                StartNode = StartNode.NextNode;
            }

            return node;
        }
        private void SelectTreeViewNode()
        {
            int selIndex = comboBox3.SelectedIndex;
            string pathNode = enteredReportBefore[selIndex, 10].ToString().Replace("\r\n", "");
            pathNode = pathNode.Replace("/", "\\");

            string[] arrPathNode = pathNode.Split(new[] { "\\" }, StringSplitOptions.RemoveEmptyEntries);
            TreeNode nextNode = treeView1.Nodes[0];

            TreeNode[] treeNode = nextNode.Nodes.Find(arrPathNode[arrPathNode.Length - 1].ToString(), true);

            /*for (int i = 0; i < arrPathNode.Length; i++)
            {
                foreach(TreeNode treeNode in nextNode.Nodes)
                {
                    if (nextNode.Nodes.Find(arrPathNode[arrPathNode.Length-1].ToString(), true)) ;
                }
            }*/

            foreach (TreeNode node in treeView1.Nodes)
            {
                if (node.FullPath == pathNode)
                {
                    treeView1.SelectedNode = node;
                }
            }
        }
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            //enteredReportBefore
            int selectIndex = comboBox3.SelectedIndex;
            label15.Text = enteredReportBefore[selectIndex, 3];
            label16.Text = enteredReportBefore[selectIndex, 4];
            label17.Text = enteredReportBefore[selectIndex, 6];
            label18.Text = enteredReportBefore[selectIndex, 7];
            label19.Text = enteredReportBefore[selectIndex, 8];

            PId = enteredReportBefore[selectIndex, 1];
            TypeBild1 = enteredReportBefore[selectIndex, 5];
        }

        /// <summary>
        /// Удаление выбранной строки из БД
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            //Удаление записи
            if (dataGridView1.SelectedRows.Count != 1) return;

            var selRow = dataGridView1.SelectedRows;
            int idRowToRemove = 0;

            idRowToRemove = Convert.ToInt32(selRow[0].Cells["Aproval_Id"].Value);
            
            SqlConnection conn = new SqlConnection(conString);
            SqlCommand dbCommandRemove = new SqlCommand
            {
                Connection = conn,
                CommandText = "DELETE FROM [dbo].[AbsenceRequestTable]"
                + " WHERE Aproval_Id = @idRowToRemove"
            };

            SqlParameter rowToDell = new SqlParameter("@idRowToRemove", idRowToRemove);
            dbCommandRemove.Parameters.Add(rowToDell);

            try
            {
                conn.Open();

                dbCommandRemove.ExecuteNonQuery();
                MessageBox.Show("Запись успешно удалена", "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show("Удалить запись не удалось!!! " + ex, "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            finally
            {
                conn.Close();
                //Обновление грида на форме
                SqlCommand updateGrid = new SqlCommand
                {
                    Connection = conn,
                    CommandText = "SELECT Aproval_Id, TypeOfAbsence, DateStart, DateEnd"
                    + " FROM dbo.AbsenceRequestTable"
                    + " WHERE(id_EnterUser = @UserId) AND(AprovalFlag IS NULL) OR"
                    + " (AprovalFlag = 0)"
                };
                SqlParameter UseIdParToUpdateGrid = new SqlParameter("@UserId", UserID);
                updateGrid.Parameters.Add(UseIdParToUpdateGrid);

                UpdateGrid(updateGrid, dataGridView1);
            }
        }

        /// <summary>
        /// Добавление запроса на согласование отсутствия сотрудника
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            if ((comboBox4.Text.Length == 0) || (monthCalendar1.SelectionRange == null)) return;

            DateTime dateStart = monthCalendar1.SelectionStart;
            DateTime dateEnd = monthCalendar1.SelectionEnd;
            string typeAbsence = comboBox4.Text.Trim();

            SqlConnection conn = new SqlConnection(conString);
            SqlCommand sqlCommand = new SqlCommand
            {
                Connection = conn,
                CommandText = "INSERT INTO [dbo].[AbsenceRequestTable]([id_EnterUser],[id_UserDepart],[TypeOfAbsence],[DateStart],[DateEnd])"
            + " VALUES (@UserId, @DepartId, @typeAbsence, @dateStart, @dateEnd)"
            };
            SqlParameter UseIdPar = new SqlParameter("@UserId", UserID);
            SqlParameter DepartIdPar = new SqlParameter("@DepartId", DepId);
            SqlParameter typeAbsencePar = new SqlParameter("@typeAbsence", typeAbsence);
            SqlParameter dateStartPar = new SqlParameter("@dateStart", dateStart);
            SqlParameter dateEndPar = new SqlParameter("@dateEnd", dateEnd);

            sqlCommand.Parameters.Add(UseIdPar);
            sqlCommand.Parameters.Add(DepartIdPar);
            sqlCommand.Parameters.Add(typeAbsencePar);
            sqlCommand.Parameters.Add(dateStartPar);
            sqlCommand.Parameters.Add(dateEndPar);

            try
            {
                conn.Open();

                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Запись добавленна в базу данных", "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex) { MessageBox.Show("Запись добавить не удалось!!! " + ex, "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally
            {
                conn.Close();

                SqlCommand updateGrid = new SqlCommand
                {
                    Connection = conn,
                    CommandText = "SELECT Aproval_Id, TypeOfAbsence, DateStart, DateEnd"
                    + " FROM dbo.AbsenceRequestTable"
                    + " WHERE(id_EnterUser = @UserId) AND(AprovalFlag IS NULL) OR"
                    + " (AprovalFlag = 0)"
                };
                SqlParameter UseIdParToUpdateGrid = new SqlParameter("@UserId", UserID);
                updateGrid.Parameters.Add(UseIdParToUpdateGrid);

                UpdateGrid(updateGrid, dataGridView1);

            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab.Name == "tabPage1")
            {
                SqlConnection conn = new SqlConnection(conString);
                SqlCommand updateGrid = new SqlCommand
                {
                    Connection = conn,
                    CommandText = "SELECT Aproval_Id, TypeOfAbsence, DateStart, DateEnd"
                   + " FROM dbo.AbsenceRequestTable"
                   + " WHERE(id_EnterUser = @UserId) AND(AprovalFlag IS NULL) OR"
                   + " (AprovalFlag = 0)"
                };
                SqlParameter UseIdParToUpdateGrid = new SqlParameter("@UserId", UserID);
                updateGrid.Parameters.Add(UseIdParToUpdateGrid);

                UpdateGrid(updateGrid, dataGridView1);
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
            SqlCommand DbSomeID = new SqlCommand
            {
                CommandText = myQuary,
                Connection = DbConn
            };
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
                SqlConnection dbConn = new SqlConnection(conString);
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

                DbComandProj.Connection = dbConn;
                DbComandTask.Connection = dbConn;

                try
                {
                    dbConn.Open();
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
                    dbConn.Close();
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

        private void MyTreeViewWork(TreeView treeView)
        {
            //TreeView treeView = treeView1;
            TreeNode ProjectsNode = new TreeNode
            {
                Name = "Projects",
                Text = "Проекты"
            };
            treeView.Nodes.Add(ProjectsNode);

            SqlConnection Conn = new SqlConnection(conString);
            SqlCommand AllProj = new SqlCommand
            {
                Connection = Conn,
                CommandText = "SELECT Id_project, Name_Project"
                    + " FROM dbo.Proekti"
                + " WHERE(InVork = 1)"
            };

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
                    SqlCommand sqlTypeBild = new SqlCommand
                    {
                        Connection = Conn,
                        CommandText = "SELECT dbo.TypeBild.NameBild, dbo.ProjBild.NumOrText, dbo.ProjBild.NumTree"
                    + " FROM dbo.Proekti INNER JOIN"
                         + " dbo.ProjBild ON dbo.Proekti.Id_project = dbo.ProjBild.Id_proj INNER JOIN"
                         + " dbo.TypeBild ON dbo.ProjBild.Id_TypeBilding = dbo.TypeBild.id_TypeBild"
                    + " WHERE(dbo.Proekti.Name_Project = @NameProj)"
                    };
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
            return rez = MyVoid2(myArr, RowCount); ;
        }

        public int[,] MyVoid2(int[,] myArr, int RowCount)
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
            MyVoid3(myArr, RowCount);
            return rez = MyVoid3(myArr, RowCount); ;
        }

        public int[,] MyVoid3(int[,] myArr, int RowCount)
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
    }
}
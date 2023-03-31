using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.Odbc;
using System.Data.OleDb;
using System.IO;
using Domino;
using System.Diagnostics;

namespace CheckPassport
{
    public partial class Form1 : Form
    {
        string sqlConnStringSQL = @"User ID=USERURO; Password=useruro2015; Initial Catalog=BD_URO; Server=GP-SQL1\SQL1";
        string FullPathImport = @"D:\PASSPORTS\";
        string fileNameBank = @"C:\Users\" + Environment.UserName + @"\РЕЕСТР БАНКРОТСТВА.xlsx";
        //string tabName = "[Check$]";
        string n = Environment.UserName;
        DataSet dsPackeg = new DataSet();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //Properties.Settings.Default.strShablon = "Employees";
            txtFms.Text = Properties.Settings.Default.strFms;
            txtServer.Text = Properties.Settings.Default.strServer;
            txtDb.Text = Properties.Settings.Default.strBase;
            txtFileBank.Text = Properties.Settings.Default.strBankr;
            //txtFileShablon.Text = Properties.Settings.Default.strShablon;

            label11.Text = "Последнее обновление: " + Convert.ToString(File.GetLastWriteTime(@"\\a100228\Общая\passports.csv"));
        }

        //выбор файла выгрузки с ФМС
        private void btFileOpen_Click(object sender, EventArgs e)
        {
            txtFms.Clear();
            txtFms.Text= OpenFileDial("Выберите файл CSV", "Файл CSV (*.csv)|*.csv");
            Properties.Settings.Default.strFms = txtFms.Text;
            Properties.Settings.Default.Save();
        }

        //загрузка CSV & LN на SQL
        private void btLoad_Click(object sender, EventArgs e)
        {
            try
            {
                Properties.Settings.Default.strServer = txtServer.Text;
                Properties.Settings.Default.strBase = txtDb.Text;
                Properties.Settings.Default.Save();

                //DataSet ds = new DataSet();
                //string strConnString = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" +
                //Path.GetDirectoryName(txtFms.Text) + ";Extensions=asc,csv,tab,txt";
                //OdbcConnection connOleDb = new OdbcConnection(strConnString);
                //connOleDb.Open();
                //string SqlSelect= "select Count(*) from  [" + Path.GetFileName(txtFms.Text) + "];";
                //OdbcDataAdapter obj_oledb_da = new OdbcDataAdapter(SqlSelect, connOleDb);
                //obj_oledb_da.Fill(ds,"PASS");
                //dataGridView1.DataSource = ds.Tables["PASS"];
                //connOleDb.Close();

                //из LN на SQL
                string serverName = txtServer.Text;
                string dbName = txtDb.Text;
                dynamic Bodys;

                NotesSession Session;
                NotesDatabase db;
                NotesView view;
                NotesDocument doc;

                Session = new NotesSession();
                Session.Initialize("");
                db = Session.GetDatabase(serverName, dbName);
                if (db.IsOpen == false)
                {
                    MessageBox.Show("Неудалось подключиться к базе данных Lotus", "Подключение к Lotus ... ");
                    return;
                }
                view = db.GetView("Все документы");
                doc = view.GetFirstDocument();
                while (doc != null)
                {
                    Bodys = doc.GetFirstItem("Body");
                    foreach (NotesEmbeddedObject fname in Bodys.EmbeddedObjects)
                    {
                        string tema = doc.GetFirstItem("Subject").Text;
                        switch (tema)
                        {
                            case "Перечень физических лиц, в отношении которых принимались меры по ПОД/ФТ":
                                fname.ExtractFile(FullPathImport + fname.Source);
                                string ListNamePod = "[Лист1$]";
                                string tableNamePOD = "PASSPORT_POD";
                                ExportExcelToSql(Path.GetFullPath(FullPathImport + fname.Source), ListNamePod, tableNamePOD);
                                break;
                            case "Список ФЛ, которым необходимо отказать в принятии на обслуживание на основани 115 ФЗ\r\n":
                                fname.ExtractFile(FullPathImport + fname.Source);
                                string ListNameFz = "[Таблица$]";
                                string tableNameFZ = "PASSPORT_FZ";
                                ExportExcelToSql(Path.GetFullPath(FullPathImport + fname.Source), ListNameFz, tableNameFZ);
                                break;
                        }
                        doc = view.GetNextDocument(doc);
                    }
                }
                StatusLabel.Text = "Готово";
                statusStrip1.Refresh();
                MessageBox.Show("Загрузка выполнена успешно!", "Информация");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Загрузка не выполнена! \n" + ex.Message, "Ошибка");
            }
        }

        //функция импорта данных Excel на SQL
        public void ExportExcelToSql(string FolderName, string ListName, string tableName)
        {
            try
            {
                string conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FolderName + ";Extended Properties=Excel 12.0";
                OleDbConnection conOledb = new OleDbConnection(conStr);
                conOledb.Open();
                OleDbDataAdapter OleDbda = new OleDbDataAdapter("select * from " + ListName, conOledb);

                //очистка таблицы
                StatusLabel.Text = "Очистка таблицы " + tableName + " ...";
                statusStrip1.Refresh();

                string strClear = "DELETE FROM " + tableName;
                SqlConnection conStrSQl = new SqlConnection(sqlConnStringSQL);
                conStrSQl.Open();
                SqlCommand cmdCl = new SqlCommand(strClear, conStrSQl);
                cmdCl.ExecuteNonQuery();
                conStrSQl.Close();

                //пакетная загрука
                StatusLabel.Text = "Загрузка файла  " + Path.GetFileName(FolderName) + " ...";
                statusStrip1.Refresh();
                OleDbCommand cmdOledb = new OleDbCommand("select * from " + ListName, conOledb);
                OleDbDataReader dr = cmdOledb.ExecuteReader();
                SqlBulkCopy BulkCopy = new SqlBulkCopy(sqlConnStringSQL);
                BulkCopy.DestinationTableName = tableName;
                BulkCopy.WriteToServer(dr);
                BulkCopy.Close();
                conOledb.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Невозможно завершить функцию импорта файла  " + Path.GetFileName(FolderName) + " \n " + ex.Message, "Внимание");
            }
        }

        //функция открытия файлов
        public string OpenFileDial(string Titl, string Filt)
        {
            string textFile = "";
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = Titl;
            dlg.Filter = Filt;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                textFile = dlg.FileName;
            }
            return textFile;
        }

        //поиск в таблице ФМС SQL
        private void btSearchFms_Click(object sender, EventArgs e)
        {
            SqlConnection strConn = new SqlConnection(sqlConnStringSQL);
            try
            {
                int resul;
                bool isIntSer = Int32.TryParse(txtFmsSer.Text, out resul);
                bool isIntNom = Int32.TryParse(txtFmsNom.Text, out resul);

                if (txtFmsNom.Text == "" || txtFmsSer.Text == "")
                {
                    MessageBox.Show("Введите серию и номер паспорта!", "Внимание");
                    return;
                }
                if ((isIntSer & isIntNom) == false)
                {
                    MessageBox.Show("Поля Серия и Номер должны иметь числовой формат.", "Внимание");
                    return;
                }
                StatusLabel.Text = "Проверка паспорта " + txtFmsSer.Text + " " + txtFmsNom.Text + "в базе ФМС...";
                statusStrip1.Refresh();
                SqlCommand cd = new SqlCommand("SELECT * FROM PASSPORT_FMS WHERE PassSeria = '" + txtFmsSer.Text + "' and PassNomber = '" + txtFmsNom.Text + "'", strConn);
                strConn.Open();
                string res = Convert.ToString(cd.ExecuteScalar());
                StatusLabel.Text = "Готово";
                statusStrip1.Refresh();
                if (res == "")
                {
                    MessageBox.Show("Данный паспорт «Среди недействительных не значится».", "Проверка ФМС", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Данный паспорт «Не действителен».", "Проверка ФМС", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                strConn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                strConn.Close();
            }
        }

        //поиск в таблицах ФЗ, ПОД SQL, Банкроты
        private void btSearchRiski_Click(object sender, EventArgs e)
        {
            try
            {
                int resul;
                bool isIntSer = Int32.TryParse(txtFmsSer.Text, out resul);
                bool isIntNom = Int32.TryParse(txtFmsNom.Text, out resul);
                string serchFio;

                if (txtFmsSer.Text != "" & txtFmsNom.Text != "")
                {
                    if ((isIntSer & isIntNom) == false)
                    {
                        MessageBox.Show("Поля Серия и Номер должны иметь числовой формат.", "Внимание");
                        return;
                    }
                }
                if ((txtFmsNom.Text == "" & txtFmsSer.Text == "") & (txtFio.Text == ""))
                {
                    MessageBox.Show("Для поиска необходимо ввести либо данные паспорта или ФИО клиента.", "Внимание");
                    return;
                }
                GridViewFZ.DataSource = null;
                GridViewPOD.DataSource = null;
                if (txtFio.Text != "")
                {
                    serchFio = txtFio.Text;
                }
                else
                {
                    serchFio = "null";
                }
                string strSqlFz = "SELECT [fio] as ФИО,[id_eq] as [ИД EQ] ,LEFT([data_br],10) as [Дата рождения],[pass] as Паспорт" +
                                          " ,[comment] as Комментарий FROM PASSPORT_FZ" +
                                          " WHERE REPLACE([fio],' ','') LIKE '%" + serchFio.Replace(" ", "") + "%' " +
                                          " and REPLACE([pass],' ','') = '" + (txtFmsSer.Text + txtFmsNom.Text).Replace(" ", "") + "' order by [fio]";

                string strSqlPod = "SELECT [fio] as ФИО,[id_eq] as [ИД EQ] ,LEFT([date_brr],10)  as [Дата рождения],[prichina_brr]  as [Причина БРР] FROM PASSPORT_POD" +
                                          " WHERE REPLACE([fio],' ','') LIKE '%" + serchFio.Replace(" ", "") + "%' order by [fio]";

                //банкротство------------------------------------------------------------------------------
                FileInfo f = new FileInfo(@"\\moscow\itfs\Банкротство физлиц\РЕЕСТР БАНКРОТСТВА.xlsx");
                
                f.CopyTo(fileNameBank,true);
                StatusLabel.Text = "Проверка по Банкротам ФЛ ...";
                statusStrip1.Refresh();
                string bank = "SELECT [ФИО клиента], [ID] as [ИД EQ], [Адрес регистрации], [наименование суда] as [Наименование суда], [Дата поступления информации] FROM " + "[Поданный заявления$]" +
                                     " WHERE (REPLACE([ФИО клиента],' ','') LIKE '%" + serchFio.Replace(" ", "") + "%') and ([ФИО клиента] is not null) order by [ФИО клиента]";
                OleDbConnection olcn = new OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + fileNameBank + " ; Extended Properties = Excel 12.0");
                olcn.Open();
                OleDbDataAdapter daBank = new OleDbDataAdapter(bank, olcn);
                DataTable dtbank = new DataTable();
                daBank.Fill(dtbank);
                GridViewBank.DataSource = dtbank.DefaultView;
                lblBank.Text = "Количество записей - " + Convert.ToString(dtbank.Rows.Count);
                dtbank.Dispose();
                olcn.Close();
                //------------------------------------------------------------------------------------------

                DataTable dtFZ = new DataTable();
                DataTable dtPOD = new DataTable();

                SqlConnection strConn = new SqlConnection(sqlConnStringSQL);
                strConn.Open();

                SqlDataAdapter daFZ = new SqlDataAdapter(strSqlFz, strConn);
                SqlDataAdapter daPOD = new SqlDataAdapter(strSqlPod, strConn);

                StatusLabel.Text = "Проверка по 115-ФЗ...";
                statusStrip1.Refresh();
                daFZ.Fill(dtFZ);

                StatusLabel.Text = "Проверка по ПОД/ФТ ...";
                statusStrip1.Refresh();
                daPOD.Fill(dtPOD);

                GridViewFZ.DataSource = dtFZ.DefaultView;
                GridViewPOD.DataSource = dtPOD.DefaultView;

                lblFz.Text = "Количество записей - " + Convert.ToString(dtFZ.Rows.Count);
                lblPOD.Text = "Количество записей - " + Convert.ToString(dtPOD.Rows.Count);

                strConn.Close();
                StatusLabel.Text = "Готово";
                statusStrip1.Refresh();
                MessageBox.Show("Выполнено.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //параметры
        private void btApply_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.strShablon = txtFileShablon.Text;
            Properties.Settings.Default.Save();
        }

        //загрузка файла открытия
        private void btOpenPackeg_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "Выберите пакетный шаблон xls";
            dlg.Filter = "Файл Excel (*.xls*)|*.xls*";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtFilePackeg.Clear();
                txtFilePackeg.Text = dlg.FileName;
            }
            StatusLabel.Text = "Загрузка шаблона: " + Path.GetFileName(dlg.FileName);
            statusStrip1.Refresh();

            string strPackeg = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source= " + Path.GetFullPath(dlg.FileName) + "; Extended Properties=Excel 12.0";
            string strBankr = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source= " + Path.GetFullPath(fileNameBank) + "; Extended Properties=Excel 12.0";
            OleDbConnection conPackeg = new OleDbConnection(strPackeg);
            OleDbConnection conBankr = new OleDbConnection(strBankr);
            try
            {
                FileInfo f = new FileInfo(@"\\moscow\itfs\Банкротство физлиц\РЕЕСТР БАНКРОТСТВА.xlsx");
                f.CopyTo(fileNameBank, true);

                SqlConnection conDel = new SqlConnection(sqlConnStringSQL);
                SqlCommand cdDel = new SqlCommand("DELETE FROM PASSOPRT_SHABLON WHERE Пользователь = '" + n + "'", conDel);
                SqlCommand cdDelBank = new SqlCommand("DELETE FROM PASSPORT_BANKROT WHERE ПользовательБД = '" + n + "'", conDel);
                conDel.Open();
                cdDel.ExecuteNonQuery();
                cdDelBank.ExecuteNonQuery();
                conDel.Close();
                //----------------------------------------------------------------------------------------------------------------
                OleDbDataAdapter daShablon = new OleDbDataAdapter("SELECT Фамилия, Имя, Отчество, [Номер документа] FROM " + "[" + txtFileShablon.Text + "$]" + " WHERE Фамилия is not null and Имя is not null and Отчество is not null order by [Фамилия]", conPackeg);
                OleDbDataAdapter daBank = new OleDbDataAdapter("SELECT [ФИО клиента], [ID] as [ИД EQ], [Адрес регистрации], [наименование суда] as [Наименование суда], [Дата поступления информации] FROM " + "[Поданный заявления$]" +
                                     " WHERE [ФИО клиента] is not null order by [ФИО клиента]",conBankr);
                conPackeg.Open();
                conBankr.Open();
                DataTable dtBank = new DataTable();
                DataTable dtShablon = new DataTable();
                daShablon.Fill(dtShablon);
                daBank.Fill(dtBank);
                dtBank.Columns.Add("ПользовательБД", typeof(string), "'" + n + "'");
                dtShablon.Columns.Add("Пользователь", typeof(string), "'" + n + "'");

                //пакетная загрука
                SqlBulkCopy BulkCopy = new SqlBulkCopy(sqlConnStringSQL);
                BulkCopy.DestinationTableName = "PASSOPRT_SHABLON";
                BulkCopy.WriteToServer(dtShablon);
                BulkCopy.Close();
                conPackeg.Close();

                SqlBulkCopy BulkCopyBank = new SqlBulkCopy(sqlConnStringSQL);
                BulkCopyBank.DestinationTableName = "PASSPORT_BANKROT";
                BulkCopyBank.WriteToServer(dtBank);
                BulkCopyBank.Close();
                conBankr.Close();
                
                SqlConnection conD = new SqlConnection(sqlConnStringSQL);
                SqlDataAdapter fullTab = new SqlDataAdapter("SELECT * FROM PASSOPRT_SHABLON WHERE Пользователь = '" + n + "'", conD);
                conD.Open();

                DataTable dtFull = new DataTable();
                fullTab.Fill(dtFull);
                GridPackeg.DataSource = dtFull.DefaultView;
                GridPackeg.Columns[4].Visible = false;

                //------------------------------------------------------------------------------------------------------


                //------------------------------------------------------------------------------------------------------

                StatusLabel.Text = "Готово";
                statusStrip1.Refresh();
                lblStatus.Text = "Количество записей - " + Convert.ToString(dtFull.Rows.Count);
                conD.Close();
            }
            catch (OleDbException)
            {
                MessageBox.Show("В файле отсутствует лист Employees или данный файл не является <Файлом открытия>.");
                StatusLabel.Text = "";
                statusStrip1.Refresh();
                conPackeg.Close();
                conBankr.Close();
            }
        }

        //пакетный поиск в БД
        private void btAllSearchPackeg_Click(object sender, EventArgs e)
        {

            if (txtFilePackeg.Text == "")
            {
                MessageBox.Show("Проверка не возможна. Выберите файл шаблона.");
                return;
            }
            SqlConnection conAllQuery = new SqlConnection(sqlConnStringSQL);
            try
            {
                string strFmsQuery = "update PASSOPRT_SHABLON   " +
                                        "set  [Результат_ФМС] = 'Не действителен'  " +
                                        "WHERE replace(Паспорт,' ','') in   " +
                                                "(select replace(Паспорт,' ','')  " +
                                                "from dbo.PASSOPRT_SHABLON inner join dbo.PASSPORT_FMS  " +
                                                "on replace(Паспорт,' ','') = PassSeria+PassNomber where [Пользователь] = '" + n + "') ";
                string strFzQuery = "update PASSOPRT_SHABLON  " +
                                        "set  Результат_ФЗ115 = 'Найден' " +
                                        "WHERE replace(Фамилия+Имя+Отчество,' ','') in  " +
                                                "(select replace(fio,' ','')  " +
                                                "from dbo.PASSOPRT_SHABLON left join dbo.PASSPORT_FZ " +
                                                "on replace(Фамилия+Имя+Отчество,' ','')  = replace(fio,' ','') " +
                                                "WHERE dbo.PASSPORT_FZ.fio is not null and [Пользователь] = '" + n + "')";
                string strPodQuery = "update PASSOPRT_SHABLON " +
                                        "set  [Результат_ПОД/ФТ] = 'Найден'  " +
                                        "WHERE replace(Фамилия+Имя+Отчество,' ','') in   " +
                                                "(select replace(fio,' ','')   " +
                                                "from dbo.PASSOPRT_SHABLON left join dbo.PASSPORT_POD  " +
                                                "on replace(Фамилия+Имя+Отчество,' ','')  = replace(fio,' ','') " +
                                                "WHERE dbo.PASSPORT_POD.fio is not null and [Пользователь] = '" + n + "')";
                string strBankQuery = "update PASSOPRT_SHABLON " +
                                        "set  [Банкротство] = 'Найден'  " +
                                        "WHERE replace(Фамилия+Имя+Отчество,' ','') in   " +
                                                "(select replace(ФИО_Клиента,' ','')   " +
                                                "from dbo.PASSOPRT_SHABLON left join dbo.PASSPORT_BANKROT  " +
                                                "on replace(Фамилия+Имя+Отчество,' ','')  = replace(ФИО_Клиента,' ','') " +
                                                "WHERE [Пользователь] = '" + n + "')";
                string[] AllQuery = { strFzQuery, strPodQuery, strFmsQuery, strBankQuery };
                string[] StatusLab = { "Проверка по ФЗ-115...", "Проверка по ПОД/ФТ...", "Проверка по ФМС...", "Проверка по Банкротству..." };
                int i = 0;
                DataTable dtAll = new DataTable();
                DataTable dtFz = new DataTable();
                DataTable dtPod = new DataTable();
                DataTable dtBank = new DataTable();

                conAllQuery.Open();
                SqlCommand cdAllQuery = new SqlCommand();
                cdAllQuery.Connection = conAllQuery;
                foreach (string v in AllQuery)
                {
                    if (v == strFmsQuery)
                    {
                        cdAllQuery.CommandTimeout = 700;
                    }
                    else
                    {
                        cdAllQuery.CommandTimeout = 30;
                    }
                    StatusLabel.Text = StatusLab[i];
                    statusStrip1.Refresh();
                    cdAllQuery.CommandText = v;
                    cdAllQuery.ExecuteNonQuery();
                    i += 1;
                }

                SqlDataAdapter daFz = new SqlDataAdapter("select PASSOPRT_SHABLON.[Фамилия],PASSOPRT_SHABLON.[Имя]  " +
                ",PASSOPRT_SHABLON.[Отчество],PASSOPRT_SHABLON.[Паспорт]  " +
                ",PASSPORT_FZ.fio as ФИО,PASSPORT_FZ.[pass] as [Данные паспорта] " +
                ",PASSPORT_FZ.[kem] as [Кем выдан],PASSPORT_FZ.[id_eq] as ID_EQ,PASSPORT_FZ.[comment] as Комментарий " +
                "from PASSOPRT_SHABLON inner join PASSPORT_FZ  " +
                "on replace(PASSOPRT_SHABLON.Фамилия+PASSOPRT_SHABLON.Имя+PASSOPRT_SHABLON.Отчество,' ','')  = replace(PASSPORT_FZ.fio,' ','')  " +
                "where [Пользователь] = '" + n + "'", conAllQuery);

                SqlDataAdapter daPod = new SqlDataAdapter("select PASSOPRT_SHABLON.[Фамилия],PASSOPRT_SHABLON.[Имя] " +
                ",PASSOPRT_SHABLON.[Отчество],PASSOPRT_SHABLON.[Паспорт] " +
                ",PASSPORT_POD.fio as ФИО,PASSPORT_POD.[id_eq] as ID_EQ,PASSPORT_POD.[prichina_brr] as [Причина БРР] " +
                "from PASSOPRT_SHABLON inner join PASSPORT_POD " +
                "on replace(PASSOPRT_SHABLON.Фамилия+PASSOPRT_SHABLON.Имя+PASSOPRT_SHABLON.Отчество,' ','')  = replace(PASSPORT_POD.fio,' ','') " +
                "where [Пользователь] = '" + n + "'", conAllQuery);

                SqlDataAdapter daBank = new SqlDataAdapter("select PASSOPRT_SHABLON.[Фамилия],PASSOPRT_SHABLON.[Имя] " +
                ",PASSOPRT_SHABLON.[Отчество],PASSOPRT_SHABLON.[Паспорт] " +
                ",PASSPORT_BANKROT.[ФИО_Клиента], PASSPORT_BANKROT.[ИД_EQ], PASSPORT_BANKROT.[Адрес_регистрации], PASSPORT_BANKROT.[Наименование_суда], PASSPORT_BANKROT.[Дата_поступления] " +
                "from PASSOPRT_SHABLON inner join PASSPORT_BANKROT " +
                "on (replace(PASSOPRT_SHABLON.Фамилия+PASSOPRT_SHABLON.Имя+PASSOPRT_SHABLON.Отчество,' ','')  = replace(PASSPORT_BANKROT.ФИО_Клиента,' ','')) and (PASSOPRT_SHABLON.[Пользователь] = PASSPORT_BANKROT.[ПользовательБД]) " +
                "where [Пользователь] = '" + n + "'", conAllQuery);

                daFz.Fill(dtFz);
                GridFZ.DataSource = dtFz.DefaultView;
                label15.Text = "Количество записей - " + Convert.ToString(dtFz.Rows.Count);

                daBank.Fill(dtBank);
                GridBankr.DataSource = dtBank.DefaultView;

                daPod.Fill(dtPod);
                GridPOD.DataSource = dtPod.DefaultView;
                label16.Text = "Количество записей - " + Convert.ToString(dtPod.Rows.Count);

                SqlDataAdapter daAllQuery = new SqlDataAdapter("SELECT * FROM PASSOPRT_SHABLON where [Пользователь] = '" + n + "'", conAllQuery);
                daAllQuery.Fill(dtAll);
                GridPackeg.DataSource = dtAll.DefaultView;

                GridPackeg.Columns[4].Visible = false;
                GridBankr.Columns[5].Visible = false;

                conAllQuery.Close();
                MessageBox.Show("Выполненно");
                StatusLabel.Text = "Готово";
                statusStrip1.Refresh();
            }
            catch (Exception ex)
            {
                conAllQuery.Close();
                MessageBox.Show(ex.Message);
            }
        }

        //права на вкладку Настройки
        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            try
            {
                string[] adminuser = { "U_M0LFM", "u_m09em" };
                int i = 0;
                foreach (string v in adminuser)
                {
                    if (n == v)
                    {
                        i = i + 1;
                    }
                }
                if ((e.TabPageIndex == 2) & (i == 0))
                {
                    MessageBox.Show("У вас нет прав доступа для владки Настройки.");
                    tabControl1.SelectTab(0);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //Получение имен файлов при перетаскивании их в целевые объекты пользовательского интерфейса
        private void txtFilePackeg_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop); //Извлекаем имя перетаскиваемого файла
            if (files.Length > 0)
            {
                txtFilePackeg.Text = files[0];
            }
        }
        private void txtFilePackeg_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true) //Разрешаем Drop только файлам
            {
                e.Effect = DragDropEffects.All;
            }
        }

        //выбор файла по банкротству фл
        private void btFileOpenBankr_Click(object sender, EventArgs e)
        {
            txtFileBank.Clear();
            txtFileBank.Text = OpenFileDial("Выберите файл Excel", "Файл Excel (*.xls*)|*.xls*");
            Properties.Settings.Default.strBankr = txtFileBank.Text;
            Properties.Settings.Default.Save();
        }


        //private void checkedListBox1_DragEnter(object sender, DragEventArgs e)
        //{
        //    //Разрешаем Drop только файлам
        //    e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ?
        //                DragDropEffects.All : DragDropEffects.None;
        //}

        //private void checkedListBox1_DragDrop(object sender, DragEventArgs e)
        //{
        //    //Извлекаем имя перетаскиваемого файла
        //    string[] strings = (string[])e.Data.GetData(DataFormats.FileDrop, true);
        //    checkedListBox1.Items.Add(strings[0]);
        //}

        //private void listBox1_DragEnter(object sender, DragEventArgs e)
        //{
        //    //Разрешаем Drop только файлам
        //    e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ?
        //                DragDropEffects.All : DragDropEffects.None;
        //}

        //private void listBox1_DragDrop(object sender, DragEventArgs e)
        //{
        //    //Извлекаем имя перетаскиваемого файла
        //    string[] strings = (string[])e.Data.GetData(DataFormats.FileDrop, true);
        //    listBox1.Items.Add(strings[0]);
        //}

        //private void comboBox1_DragEnter(object sender, DragEventArgs e)
        //{
        //    //Разрешаем Drop только файлам
        //    e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ?
        //                DragDropEffects.All : DragDropEffects.None;
        //}

        //private void comboBox1_DragDrop(object sender, DragEventArgs e)
        //{
        //    //Извлекаем имя перетаскиваемого файла
        //    string[] strings = (string[])e.Data.GetData(DataFormats.FileDrop, true);
        //    comboBox1.Items.Add(strings[0]);
        //}

        //private void textBox1_DragEnter(object sender, DragEventArgs e)
        //{
        //    //Разрешаем Drop только файлам
        //    e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ?
        //                DragDropEffects.All : DragDropEffects.None;
        //}

        //private void textBox1_DragDrop(object sender, DragEventArgs e)
        //{
        //    //Извлекаем имя перетаскиваемого файла
        //    string[] strings = (string[])e.Data.GetData(DataFormats.FileDrop, true);
        //    textBox1.Text = strings[0];
        //}

        //private void treeView1_DragEnter(object sender, DragEventArgs e)
        //{
        //    //Разрешаем Drop только файлам
        //    e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ?
        //                DragDropEffects.All : DragDropEffects.None;
        //}

        //private void treeView1_DragDrop(object sender, DragEventArgs e)
        //{
        //    //Извлекаем имена перетаскиваемых файлов
        //    string[] FileList = (string[])e.Data.GetData(DataFormats.FileDrop, true);
        //    foreach (string File in FileList)
        //        treeView1.Nodes.Add(File);
        //}


    }
}
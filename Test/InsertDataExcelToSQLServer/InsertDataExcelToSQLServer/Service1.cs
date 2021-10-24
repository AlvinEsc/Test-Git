using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Configuration;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.Timers;

namespace InsertDataExcelToSQLServer
{
    public partial class Service1 : ServiceBase
    {
        private Timer Timer = null;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Timer = new Timer();
            //1000 = 1 Detik
            this.Timer.Interval = 60000;//1 Menit
            this.Timer.Elapsed += new ElapsedEventHandler(this.Timer_Tick);
            this.Timer.Enabled = true;
        }

        private void Timer_Tick(object sender, ElapsedEventArgs e)
        {
            //Menjalankan Program ImportDataFromExcel
            InsertDataFromExcel();
            //Menjalankan Program Menghapus File
            DeleteFile();
        }

        public void InsertDataFromExcel()
        {
            string ExcelFilePath = "C:\\Users\\ALVIN ANTONIUS\\Documents\\Table_ExcelSQL.xls";
            string SQLTable = "Table_ExcelSQL";
            string MyExcelData = "select Nama,Jurusan,NIM from [Sheet1$]";
            try
            {
                //Membuat Koneksi String
                string ExcelConString = @"provider=microsoft.jet.oledb.4.0;data source=" + ExcelFilePath + ";extended properties=" + "\"excel 8.0;hdr=yes;\"";
                string SQLConString = "Data Source = LAPTOP-UTVM3T1U; initial catalog = DB_TBL_ExcelSQL; integrated security = true";

                //Program Untuk Menyalin Data Yang Ada Di Excel Ke SQL Table
                OleDbConnection OleDbcon = new OleDbConnection(ExcelConString);
                OleDbCommand OleDbcmd = new OleDbCommand(MyExcelData, OleDbcon);
                OleDbcon.Open();
                OleDbDataReader DR = OleDbcmd.ExecuteReader();
                SqlBulkCopy BulkCopy = new SqlBulkCopy(SQLConString);
                BulkCopy.DestinationTableName = SQLTable;
                while (DR.Read())
                {
                    BulkCopy.WriteToServer(DR);
                }
                DR.Close();
                OleDbcon.Close();
            }
            catch
            {

            }
        }

        public void DeleteFile()
        {
            try
            {
                string DeleteFile = "C:\\Users\\ALVIN ANTONIUS\\Documents\\Table_ExcelSQL.xls";
                File.Delete(DeleteFile);
            }
            catch
            {
            }
        }

        protected override void OnStop()
        {
        }
    }
}

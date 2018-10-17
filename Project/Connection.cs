using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
namespace Project
{
    class Connection
    {
        OleDbConnection OlConn = new OleDbConnection();
         string ConnString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Book1.xlsx;Extended Properties='Excel 12.0;HDR=YES;';";

        public Connection()
        {
            OlConn = new OleDbConnection(ConnString);
        }
        public void OpenConnection()
        {
            if (OlConn.State == ConnectionState.Closed)
            {
                OlConn.Open();
            }
        }
        public void CloseConnection()
        {
            if (OlConn.State == ConnectionState.Open)
            {
                OlConn.Close();
            }
        }
        public void ExecuteCommand(string Query)
        {
            OleDbCommand cmd = new OleDbCommand();
            cmd.CommandType = CommandType.Text;
            cmd.Connection = OlConn;
            cmd.CommandText = Query;
            OpenConnection();
            cmd.ExecuteNonQuery();
            CloseConnection();
        }
        public void ExecuteProcedure(string ProcedureName, OleDbParameter[] param)
        {
            OleDbCommand cmd = new OleDbCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = ProcedureName;
            cmd.Connection = OlConn;
            OpenConnection();
            for (int i = 0; i < param.Length; i++)
            {
                cmd.Parameters.Add(param[i]);
            }
            cmd.ExecuteNonQuery();
            CloseConnection();
        }
    }
}

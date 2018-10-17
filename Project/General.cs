using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Threading;
using System.Data.OleDb;
using System.Net;
using System.IO;
using System.Reflection;
namespace Project
{
    class General
    {
        string conlink3 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Book1.xlsx;Extended Properties='Excel 12.0;HDR=YES;';";
        string Query = "";
        private static int Log_ID = 1;
        public void EventLog(string LogData, string Status)
        {
            OleDbConnection con;
            OleDbCommand InsertBooking;
            con = new OleDbConnection(conlink3);
            con.Open();
            Query = "INSERT INTO [sheet3$] (id, Data, Status)" + "" + "VALUES(@id, @value1, @value2)";
            InsertBooking = new OleDbCommand(Query, con);
            InsertBooking.Parameters.AddWithValue("@id", Log_ID);
            InsertBooking.Parameters.AddWithValue("@value1", LogData);
            InsertBooking.Parameters.AddWithValue("@value2", Status);
            InsertBooking.ExecuteNonQuery();
            con.Close();
            Log_ID++;
        }
    }
}

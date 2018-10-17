using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
namespace Project
{
    public partial class frmReport : Form
    {
        string conlink3 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Book1.xlsx;Extended Properties='Excel 12.0;HDR=YES;';";
        public frmReport()
        {
            InitializeComponent();
        }

        private void frmReport_Load(object sender, EventArgs e)
        {
            GenerateReport();
        }
        private void GenerateReport()
        {
            OleDbConnection con;
            con = new OleDbConnection(conlink3);
            con.Open();
            string Query = "SELECT * FROM [sheet2$]";
            DataTable dt = new DataTable();
            dt.Clear();
            OleDbDataAdapter da = new OleDbDataAdapter(Query, con);
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }
    }
}

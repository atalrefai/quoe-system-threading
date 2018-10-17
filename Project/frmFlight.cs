using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Data.OleDb;
namespace Project
{
    public partial class frmFlight : Form
    {
        public Thread Thread1;
        private static Mutex mut1 = new Mutex();
        string conlink3 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Book1.xlsx;Extended Properties='Excel 12.0;HDR=YES;';";
        int Count1 = 0, Count2 = 0, Count3 = 0;
        int Number = 5;
        General gen = new General();
        public frmFlight()
        {
            InitializeComponent();
        
        }
        
        private void frmFlight_Load(object sender, EventArgs e)
        {
            this.Thread1 = new Thread(new ThreadStart(this.Thread1_Delegate));
        }
        private void Thread1_Delegate()
        {
            //Data for each person
            //string Data = "";
            try
            {
                while (true)
                {
                    Thread.Sleep(10);
                    mut1.WaitOne();
                    Count1 = listBox1.Items.Count;
                    Count2 = listBox2.Items.Count;
                    Count3 = listBox3.Items.Count;

                    mut1.ReleaseMutex();
                }
            }
            catch
            {
            }
        }
        public void Fill(string Flight_number, string Data)
        {
            if (Flight_number == textBox1.Text)
            {
                listBox1.Items.Add(Data);
                gen.EventLog(Data, "Flight");
                if (listBox1.Items.Count == Number)
                {
                    //Update Check
                    UpdateCheck(textBox1.Text);
                    listBox1.Items.Clear();

                }
            }
            else if (Flight_number == textBox2.Text)
            {
                listBox2.Items.Add(Data);
                gen.EventLog(Data, "Flight");
                if (listBox2.Items.Count == Number)
                {
                    //Update Check
                    UpdateCheck(textBox2.Text);
                    listBox2.Items.Clear();
                }
            }
            else if (Flight_number == textBox3.Text)
            {
                listBox3.Items.Add(Data);
                gen.EventLog(Data, "Flight");
                if (listBox3.Items.Count == Number)
                {
                    //Update Check
                    UpdateCheck(textBox3.Text);
                    listBox3.Items.Clear();
                }
            }
        }
        private void UpdateCheck(string FlightNumber)
        {
            OleDbConnection con1;
            OleDbCommand UpdateBooking;
            con1 = new OleDbConnection(conlink3);
            con1.Open();

            string Query = "Update [sheet2$] SET checkin = 'yes' WHere flight_no = '" + FlightNumber + "'";
            UpdateBooking = new OleDbCommand(Query, con1);
            UpdateBooking.ExecuteNonQuery();
            con1.Close();
            listBox4.Items.Add(FlightNumber + " fly" + DateTime.Now.ToShortTimeString());
        }
    }
}

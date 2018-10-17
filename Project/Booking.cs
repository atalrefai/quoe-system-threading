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
using MVCTest;
namespace Project
{
    // Form is reallly the view component which will implement the IModelObserver interface 
    // so that, it can invoke the valueIncremented function which is the implementation
    // Form also implements the IView interface to send the view changed event and to
    // set the controller associated with the view
    public partial class Booking : Form, IView, IModelObserver
    {
        IController controller1;
        IController controller2;
        public event ViewHandler<IView> changed;
        // View will set the associated controller, this is how view is linked to the controller.
        public void setController(IController cont)
        {
            controller1 = cont;
            controller2 = cont;
        }
        private Thread Thread_Queue;
        private Thread Thread_Display;
        private Thread Thread_Weight;
        private Thread Thread_Flight;
        private static Mutex mutQueue = new Mutex();
        private static Mutex mutDisplay = new Mutex();
        private static Mutex mutWeight = new Mutex();
        private static Mutex mutFlight = new Mutex();
        Queue queue = new Queue();
        Queue Save = new Queue();
        Queue SaveFlight = new Queue();
        static Queue SaveCopy = new Queue();
        string conlink3 = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=Book1.xlsx;Extended Properties='Excel 12.0;HDR=YES;';";
        string Query = "";
        int c1 = 0;
        //to store number of passenger randomized by thread
        int NumberOfPassenger = 0;
        string[] fname = { "Devora", "Fannie", "Charley", "Dorie", "Maia", "Bridget", "Johnathan", "Yer", "Annis", "Jeffry", "Karma", "Billy", "Joann", "Stacey", "Ricardo", "Brendan", "Carlena", "Maddie", "Hazel", "Marcella", "Hershel", "Tonia", "Pauline", "Ema", "Helena", "Diane", "Solomon", "Takako", "Margurite", "Wesley", "Stasia", "Joslyn", "Bulah", "Fae", "Meagan", "Maryellen", "Jennine", "Gaynell", "Madlyn", "Nathanael", "Kaci", "Sheryll", "Lyn", "Cordia", "Chi", "Thea", "Providencia", "Lyndia", "Edythe", "Phoebe" };
        string[] lname = { "Neal", "Sean", "Desmond", "Christoper", "Silas", "Reinaldo", "Markus", "Jacinto", "Gus", "Garth", "Antony", "Damion", "Curt", "Brock", "Parker", "Darell", "Thurman", "Vito", "Rogelio", "Leandro", "Vincent", "Austin", "Maximo", "Randal", "Abdul", "Zachery", "Patricia", "Angelo", "Andy", "Arden", "Virgil", "Dorian", "Jeramy", "Jonathon", "Shawn", "Salvador", "Sterling", "Dusty", "Haywood", "Porfirio", "Scott", "Blaine", "Bob", "Roscoe", "William", "Ulysses", "Hosea", "Patrick", "Michael", "Claud" };
        string[] grade_type = { "economic", "Businessmen" };
        int[] seat_lineE = { 1, 2, 3 };
        int[] seat_lineB = { 4, 5, 6, 7, 8, 9, 10 };
        string[] location_seat = { "Right", "Left" };
        string[] seat_position = { "A", "B", "C" };
        int[] flight = new int[3];
        int Passenger_Count = 0;
        int Current_Passenger = 1;
        int Passenger_ID = 1;
        double Price = 20;
        private frmQueue fq;
        private frmWeight fw;
        private frmFlight ff;
        Random rnd = new Random();
        
        int LogID = 1;
        General gen = new General();
        string result = null;
        string url = "http://alkhalilschools.com/checkupdate.txt";
        WebResponse response = null;
        StreamReader reader = null;
        public Booking()
        {
            InitializeComponent();
            updateve();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //identify threads list
            this.Thread_Queue = new Thread(new ThreadStart(this.Thread_Queue_Delegate));
            this.Thread_Display = new Thread(new ThreadStart(this.Thread_Display_Delegate));
            this.Thread_Weight = new Thread(new ThreadStart(this.Thread_Weight_Delegate));
            this.Thread_Flight = new Thread(new ThreadStart(this.Thread_Flight_Delegate));
            //list identify forms
            this.fq = new frmQueue();
            this.fq.Show();
            this.fq.Left = base.Right;
            this.fw = new frmWeight();
            this.fw.Show();
            this.fw.Left = base.Right;
            this.ff = new frmFlight();
            this.ff.Show();
            this.ff.Left = base.Left;
            //Read Flight
            ReadFlight();
            //chang header text for DGV
            dataGridView1.Columns[1].HeaderText = "Airport";
            dataGridView1.Columns[2].HeaderText = "Source";
            dataGridView1.Columns[3].HeaderText = "Destination";
            dataGridView1.Columns[4].HeaderText = "Time start";
            dataGridView1.Columns[5].HeaderText = "Time Arrive";
            dataGridView1.Columns[6].HeaderText = "Date";
            dataGridView1.Columns[7].HeaderText = "Flight NO";
            dataGridView1.Columns[8].Visible = false;
        }
        public void updateve()
        {
            //Check if update available or not
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "GET";
                response = request.GetResponse();
                reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                result = reader.ReadToEnd();
                if (result == "1")
                {
                    label8.Text = "New Update Avilable";
                    label8.ForeColor = Color.Red;
                }
                else
                {
                    label8.Text = "This is last Update";
                    label8.ForeColor = Color.Green;
                }
            }
            catch (Exception ex)
            {
                // handle error
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (response != null)
                    response.Close();
            }
        }
        private void ReadFlight()
        {
            //Connect to excel file and get data from sheet1 and store data in flight array
            OleDbConnection con;
            con = new OleDbConnection(conlink3);
            con.Open();
            Query = "SELECT * FROM [sheet1$]";
            DataTable dt = new DataTable();
            dt.Clear();
            OleDbDataAdapter da = new OleDbDataAdapter(Query, con);
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            //Check if data is not null
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < flight.Length; i++)
                {
                    flight[i] = int.Parse(dt.Rows[i][0].ToString());
                }
            }
        }
        private void btnStart_Click(object sender, EventArgs e)
        {

            Passenger_Count = int.Parse(txtPassengerCount.Text);
            panel1.Enabled = false;
            //Start first thread
            this.Thread_Queue.Start();
            btnStart.Enabled = false;
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            //Displose all forms
            this.ff.Dispose();
            this.fw.Dispose();
            this.fq.Dispose();
            this.Hide();
            //Display report form when user click exit
            frmReport f = new frmReport();
            f.ShowDialog();
            Environment.Exit(0);
        }
        private void btnIncrement_Click(object sender, EventArgs e)
        {
            // when the user clicks the button just ask the controller to increment the value
            controller1.incvalue();
        }
        private void btnDecrement_Click(object sender, EventArgs e)
        {
            // when the user clicks the button just ask the controller to decrement the value
            controller1.decvalue();
        }
        public void valueIncremented(IModel m, ModelEventArgs e)
        {
            // This event is implementation from IModelObserver which will be invoked by the
            // Model when there is a change in the value with ModelEventArgs which is derived
            // from the EventArgs. The IModel object is the one from which invoked this.
            txtSleepCounter.Text = "" + e.newval;
        }
        public void valueDecremented(IModel m, ModelEventArgs e)
        {
            // This event is implementation from IModelObserver which will be invoked by the
            // Model when there is a change in the value with ModelEventArgs which is derived
            // from the EventArgs. The IModel object is the one from which invoked this.
            txtSleepCounter.Text = "" + e.newval;
        }
        private void txtSleepCounter_Leave(object sender, EventArgs e)
        {
            //Check if user input a valid number in txtSleepCounter
            try
            {
                changed.Invoke(this, new ViewEventArgs(int.Parse(txtSleepCounter.Text)));
            }
            catch (Exception)
            {
                MessageBox.Show("Please enter a valid number");
                txtSleepCounter.Focus();
            }
        }
        private void btnInc_Click(object sender, EventArgs e)
        {
            //Increment txtkg by 1
            txtkg.Text = (int.Parse(txtkg.Text) + 1).ToString();
        }
        private void btnDec_Click(object sender, EventArgs e)
        {
            //Decrement txtkg by 1
            txtkg.Text = (int.Parse(txtkg.Text) - 1).ToString();
        }
        private void txtkg_Leave(object sender, EventArgs e)
        {
            //Check if user input a valid number in txtkg
            try
            {
                changed.Invoke(this, new ViewEventArgs(int.Parse(txtkg.Text)));
            }
            catch (Exception)
            {
                MessageBox.Show("Please enter a valid number");
                txtkg.Focus();
            }
        }
        private void btnIncPassenger_Click(object sender, EventArgs e)
        {
            //Increment txtPassengerCount by 1
            txtPassengerCount.Text = (int.Parse(txtPassengerCount.Text) + 1).ToString();
        }
        private void btnDecPassenger_Click(object sender, EventArgs e)
        {
            //Decrement txtPassengerCount by 1
            txtPassengerCount.Text = (int.Parse(txtPassengerCount.Text) - 1).ToString();
        }
        private void txtPassengerCount_Leave(object sender, EventArgs e)
        {
            //Check if user input a valid number in txtPassengerCount
            try
            {
                changed.Invoke(this, new ViewEventArgs(int.Parse(txtPassengerCount.Text)));
            }
            catch (Exception)
            {
                MessageBox.Show("Please enter a valid number");
                txtPassengerCount.Focus();
            }
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Displose all forms
            this.ff.Dispose();
            this.fw.Dispose();
            this.fq.Dispose();
            this.Hide();
            //Display report form when user click exit
            frmReport f = new frmReport();
            f.ShowDialog();
            Environment.Exit(0);
        }

        private void Thread_Queue_Delegate()
        {
            //Generate random data for each person
            int RandomNumber;
            string Data = "";
            try
            {
                while (true)
                {
                    Thread.Sleep(int.Parse(txtSleepCounter.Text));
                    mutQueue.WaitOne();
                    //Generate random first name from array
                    RandomNumber = rnd.Next(0, 50);
                    //Add it to queue
                    queue.Enqueue(fname[RandomNumber]);
                    //Add it to variable data
                    Data += fname[RandomNumber] + " - ";
                    //Generate random last name from array
                    RandomNumber = rnd.Next(0, 50);
                    //Add it to queue
                    queue.Enqueue(lname[RandomNumber]);
                    //Add it to variable data
                    Data += lname[RandomNumber] + " - ";
                    //Generate random passport number
                    RandomNumber = rnd.Next(000000000, 999999999);
                    //Add it to queue
                    queue.Enqueue(RandomNumber);
                    //Add it to variable data
                    Data += RandomNumber + " - ";
                    //Generate random grade type from array
                    RandomNumber = rnd.Next(0, 2);
                    //Add it to queue
                    queue.Enqueue(grade_type[RandomNumber]);
                    //Add it to variable data
                    Data += grade_type[RandomNumber] + " - ";
                    //Check if grade type is economic
                    if (RandomNumber == 0)
                    {
                        //Generate random seat line from array
                        RandomNumber = rnd.Next(0, 3);
                        //Add it to queue
                        queue.Enqueue(seat_lineE[RandomNumber]);
                        //Add it to variable data
                        Data += seat_lineE[RandomNumber] + " - ";
                    }
                    //Check if grade type is Businessmen
                    else
                    {
                        //Generate random seat line from array
                        RandomNumber = rnd.Next(0, 7);
                        //Add it to queue
                        queue.Enqueue(seat_lineB[RandomNumber]);
                        //Add it to variable data
                        Data += seat_lineB[RandomNumber] + " - ";
                    }
                    //Generate random location seat from array
                    RandomNumber = rnd.Next(0, 2);
                    //Add it to queue
                    queue.Enqueue(location_seat[RandomNumber]);
                    //Add it to variable data
                    Data += location_seat[RandomNumber] + " - ";
                    RandomNumber = rnd.Next(0, 3);
                    //Add it to queue
                    queue.Enqueue(seat_position[RandomNumber]);
                    //Add it to variable data
                    Data += seat_position[RandomNumber] + " - ";
                    //choose random flight number
                    RandomNumber = rnd.Next(0, flight.Length);
                    RandomNumber = flight[RandomNumber];
                    //Add data to log file with waiting status
                    gen.EventLog(Data, "Waiting");
                    OleDbConnection con;
                    con = new OleDbConnection(conlink3);
                    con.Open();
                    //Get flight details from sheet1 by flight id
                    Query = "SELECT * FROM [sheet1$] WHERE id = " + RandomNumber;
                    DataTable dt = new DataTable();
                    dt.Clear();
                    OleDbDataAdapter da = new OleDbDataAdapter(Query, con);
                    da.Fill(dt);
                    //Check if data is not null
                    if (dt.Rows.Count > 0)
                    {
                        //Add airport_name to queue
                        queue.Enqueue(dt.Rows[0][1].ToString());
                        //Add it to variable data
                        Data += dt.Rows[0][1].ToString() + " - ";
                        //Add Source_c to queue
                        queue.Enqueue(dt.Rows[0][2].ToString());
                        //Add it to variable data
                        Data += dt.Rows[0][2].ToString() + " - ";
                        //Add Dest_c to queue
                        queue.Enqueue(dt.Rows[0][3].ToString());
                        //Add it to variable data
                        Data += dt.Rows[0][3].ToString() + " - ";
                        //Add date_c to queue
                        queue.Enqueue(dt.Rows[0][4].ToString());
                        //Add it to variable data
                        Data += dt.Rows[0][4].ToString() + " - ";
                        //Add time_st to queue
                        queue.Enqueue(dt.Rows[0][5].ToString());
                        //Add it to variable data
                        Data += dt.Rows[0][5].ToString() + " - ";
                        //Add time_ar to queue
                        queue.Enqueue(dt.Rows[0][6].ToString());
                        //Add it to variable data
                        Data += dt.Rows[0][6].ToString() + " - ";
                        //Add flight_no to queue
                        queue.Enqueue(dt.Rows[0][7].ToString());
                        //Add it to variable data
                        Data += dt.Rows[0][7].ToString();
                    }
                    //Add data to listbox in frmQueue
                    this.Invoke(new MethodInvoker(delegate { this.fq.listBox1.Items.Add(Data); }));
                    //Clear data variable
                    Data = "";
                    //Get Flight number from excel file
                    string Query1 = "SELECT flight_no FROM [sheet1$]";
                    DataTable dt1 = new DataTable();
                    dt1.Clear();
                    OleDbDataAdapter da1 = new OleDbDataAdapter(Query1, con);
                    da1.Fill(dt1);
                    //Check if data is not null
                    if (dt1.Rows.Count > 0)
                    {
                        //Fill flight number in the textbox in frmFlight
                        this.Invoke(new MethodInvoker(delegate { this.ff.textBox1.Text = dt1.Rows[0][0].ToString(); }));
                        this.Invoke(new MethodInvoker(delegate { this.ff.textBox2.Text = dt1.Rows[1][0].ToString(); }));
                        this.Invoke(new MethodInvoker(delegate { this.ff.textBox3.Text = dt1.Rows[2][0].ToString(); }));
                    }
                    con.Close();
                    mutQueue.ReleaseMutex();
                    //Check if Thread_Display is not running
                    if (this.Thread_Display.ThreadState.ToString().ToLower() == "unstarted")
                    {
                        //Run Thread_Display
                        this.Thread_Display.Start();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Queue");
                MessageBox.Show(ex.Message, ex.Source);
            }
        }

        private void Thread_Display_Delegate()
        {
            try
            {
                string Queue_value = "";
                while (true)
                {
                    //Check if queue is not null
                    if (queue.Count != 0)
                    {
                        //Check if Current_Passenger reach to the limit
                        if (Current_Passenger <= Passenger_Count)
                        {
                            Thread.Sleep(int.Parse(txtSleepCounter.Text));
                            mutDisplay.WaitOne();
                            //Increment number of passenger
                            NumberOfPassenger++;
                            //Displat first item from queue and remove it from the list
                            this.Invoke(new MethodInvoker(delegate { this.fq.listBox1.Items.RemoveAt(0); }));
                            //Display number of current passenger
                            this.Invoke(new MethodInvoker(delegate { this.Text = "Number of Passenger Booking: " + NumberOfPassenger.ToString(); }));
                            //Display Items from queue and save a copy in queuesave then dequeue from the queue orginal
                            //benifit when show in queue will dequeue so we need copy the dets in other queue for wight and flight
                            Queue_value = queue.Peek().ToString();
                            Save.Enqueue(queue.Peek());
                            queue.Dequeue();
                            this.Invoke(new MethodInvoker(delegate { txtFname.Text = Queue_value; }));
                            Queue_value = queue.Peek().ToString();
                            Save.Enqueue(Queue_value);
                            queue.Dequeue();
                            this.Invoke(new MethodInvoker(delegate { txtLname.Text = Queue_value; }));
                            Queue_value = queue.Peek().ToString();
                            Save.Enqueue(Queue_value);
                            queue.Dequeue();
                            this.Invoke(new MethodInvoker(delegate { txtPass.Text = Queue_value; }));
                            Queue_value = queue.Peek().ToString();
                            Save.Enqueue(Queue_value);
                            queue.Dequeue();
                            this.Invoke(new MethodInvoker(delegate { txtGrade.Text = Queue_value; }));
                            Queue_value = queue.Peek().ToString();
                            Save.Enqueue(Queue_value);
                            queue.Dequeue();
                            this.Invoke(new MethodInvoker(delegate { txtSeatLine.Text = Queue_value; }));
                            Queue_value = queue.Peek().ToString();
                            Save.Enqueue(Queue_value);
                            queue.Dequeue();
                            this.Invoke(new MethodInvoker(delegate { txtLocationSeat.Text = Queue_value; }));
                            Queue_value = queue.Peek().ToString();
                            Save.Enqueue(Queue_value);
                            queue.Dequeue();
                            this.Invoke(new MethodInvoker(delegate { txtSeatPosition.Text = Queue_value; }));
                            Queue_value = queue.Peek().ToString();
                            Save.Enqueue(Queue_value);
                            queue.Dequeue();
                            this.Invoke(new MethodInvoker(delegate { txt1.Text = Queue_value; }));
                            Queue_value = queue.Peek().ToString();
                            Save.Enqueue(Queue_value);
                            queue.Dequeue();
                            this.Invoke(new MethodInvoker(delegate { txt2.Text = Queue_value; }));
                            Queue_value = queue.Peek().ToString();
                            Save.Enqueue(Queue_value);
                            queue.Dequeue();
                            this.Invoke(new MethodInvoker(delegate { txt3.Text = Queue_value; }));
                            Queue_value = queue.Peek().ToString();
                            Save.Enqueue(Queue_value);
                            queue.Dequeue();
                            this.Invoke(new MethodInvoker(delegate { txt4.Text = Queue_value; }));
                            Queue_value = queue.Peek().ToString();
                            Save.Enqueue(Queue_value);
                            queue.Dequeue();
                            this.Invoke(new MethodInvoker(delegate { txt5.Text = Queue_value; }));
                            Queue_value = queue.Peek().ToString();
                            Save.Enqueue(Queue_value);
                            queue.Dequeue();
                            this.Invoke(new MethodInvoker(delegate { txt6.Text = Queue_value; }));
                            Queue_value = queue.Peek().ToString();
                            Save.Enqueue(Queue_value);
                            queue.Dequeue();
                            this.Invoke(new MethodInvoker(delegate { txt7.Text = Queue_value; }));
                            mutDisplay.ReleaseMutex();
                            //Increment current passenger
                            Current_Passenger++;
                            //Add it to log 
                            gen.EventLog(txtFname.Text + " " + txtLname.Text + " " + DateTime.Now.ToString(), "Register");
                            LogID++;
                        }
                        else
                        {
                            //there queue is empty from 50 name
                            //SaveCopy = Save;
                            if (Save.Count != 0)
                            {
                                foreach (string x in Save)
                                {
                                    SaveCopy.Enqueue(x);
                                }
                            }
                            //Store in Excel File
                            Connection Conn = new Connection();
                            Conn.OpenConnection();
                            //Data for flight
                            string Data2 = "";
                            for (int i = 0; i < Passenger_Count; i++)
                            {
                                OleDbConnection con;
                                OleDbCommand InsertBooking;
                                con = new OleDbConnection(conlink3);
                                con.Open();
                                //Add Data to sheet2 in excel file
                                Query = "INSERT INTO [sheet2$] (id, fname, lname, passport, GradeType, SeatLine, locationSeat, SeatPosion, airport_name, Source_c, Dest_c, date_c, time_st, time_ar, flight_no)" + "" + "VALUES(@id, @value1, @value2, @value3, @value4, @value5, @value6, @value7, @value8, @value9, @value10, @value11, @value12, @value13, @value14)";
                                InsertBooking = new OleDbCommand(Query, con);
                                InsertBooking.Parameters.AddWithValue("@id", Passenger_ID);
                                InsertBooking.Parameters.AddWithValue("@value1", Save.Peek().ToString());
                                Data2 += Save.Peek().ToString();
                                Save.Dequeue();
                                InsertBooking.Parameters.AddWithValue("@value2", Save.Peek().ToString());
                                Data2 += " " + Save.Peek().ToString();
                                Save.Dequeue();
                                InsertBooking.Parameters.AddWithValue("@value3", Save.Peek().ToString());
                                Save.Dequeue();
                                InsertBooking.Parameters.AddWithValue("@value4", Save.Peek().ToString());
                                Data2 += " " + Save.Peek().ToString();
                                Save.Dequeue();
                                InsertBooking.Parameters.AddWithValue("@value5", Save.Peek().ToString());
                                Data2 += " " + Save.Peek().ToString();
                                Save.Dequeue();
                                InsertBooking.Parameters.AddWithValue("@value6", Save.Peek().ToString());
                                Data2 += " " + Save.Peek().ToString();
                                Save.Dequeue();
                                InsertBooking.Parameters.AddWithValue("@value7", Save.Peek().ToString());
                                Save.Dequeue();
                                InsertBooking.Parameters.AddWithValue("@value8", Save.Peek().ToString());
                                Save.Dequeue();
                                InsertBooking.Parameters.AddWithValue("@value9", Save.Peek().ToString());
                                Save.Dequeue();
                                InsertBooking.Parameters.AddWithValue("@value10", Save.Peek().ToString());
                                Save.Dequeue();
                                InsertBooking.Parameters.AddWithValue("@value11", Save.Peek().ToString());
                                Save.Dequeue();
                                InsertBooking.Parameters.AddWithValue("@value12", Save.Peek().ToString());
                                Save.Dequeue();
                                InsertBooking.Parameters.AddWithValue("@value13", Save.Peek().ToString());
                                Save.Dequeue();
                                InsertBooking.Parameters.AddWithValue("@value14", Save.Peek().ToString());
                                string FlightNumber = Save.Peek().ToString();
                                Save.Dequeue();
                                InsertBooking.ExecuteNonQuery();
                                con.Close();
                                Passenger_ID++;
                                //this.Invoke(new MethodInvoker(delegate { this.ff.Fill(FlightNumber, Data2); }));
                                SaveFlight.Enqueue(FlightNumber);
                                SaveFlight.Enqueue(Data2);
                                Data2 = "";
                            }
                            if (this.Thread_Weight.ThreadState.ToString().ToLower() == "unstarted")
                            {
                                this.Thread_Weight.Start();
                            }
                            else
                            {
                                this.Thread_Weight.Resume();
                            }
                            this.Thread_Display.Suspend();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Display");
                MessageBox.Show(ex.Message);
            }
        }

        private void Thread_Weight_Delegate()
        {
            //Save Copy
            try
            {
                while (true)
                {
                    for (int i = 0; i < Passenger_Count; i++)
                    {
                        double Tax = 0.0;
                        Thread.Sleep(int.Parse(txtSleepCounter.Text));
                        mutWeight.WaitOne();
                        Save.Enqueue(SaveCopy.Peek().ToString());
                        SaveCopy.Dequeue();
                        //Save.Enqueue(SaveCopy.Peek().ToString());
                        SaveCopy.Dequeue();
                        Save.Enqueue(SaveCopy.Peek().ToString());
                        SaveCopy.Dequeue();
                        //Save.Enqueue(SaveCopy.Peek().ToString());
                        SaveCopy.Dequeue();
                        //Save.Enqueue(SaveCopy.Peek().ToString());
                        SaveCopy.Dequeue();
                        //Save.Enqueue(SaveCopy.Peek().ToString());
                        SaveCopy.Dequeue();
                        //Save.Enqueue(SaveCopy.Peek().ToString());
                        SaveCopy.Dequeue();
                        //Save.Enqueue(SaveCopy.Peek().ToString());
                        //SaveCopy.Dequeue();
                        SaveCopy.Dequeue();
                        SaveCopy.Dequeue();
                        SaveCopy.Dequeue();
                        SaveCopy.Dequeue();
                        SaveCopy.Dequeue();
                        SaveCopy.Dequeue();
                        SaveCopy.Dequeue();
                        //SaveCopy.Dequeue();
                        double Weight = rnd.Next(0, 50);
                        Save.Enqueue(Weight);
                        if (Weight > int.Parse(txtkg.Text))
                        {
                            double extra = Weight - 30;
                            double newPrice = (extra * Price);
                            Tax = newPrice + (newPrice * 0.05);
                            //Save.Enqueue(Tax);
                        }
                        Save.Enqueue(Tax);
                    }
                    //Update Excel File
                    //Data for weight
                    string Data1 = "";
                    for (int i = 0; i < Passenger_Count; i++)
                    {
                        OleDbConnection con1;
                        OleDbCommand UpdateBooking;
                        con1 = new OleDbConnection(conlink3);
                        con1.Open();
                        Thread.Sleep(int.Parse(txtSleepCounter.Text));
                        c1++;
                        //this.Invoke(new MethodInvoker(delegate { label9.Text = c1.ToString(); }));
                        Data1 += c1.ToString() + ". ";
                        string name = Save.Peek().ToString();
                        Save.Dequeue();
                        //this.Invoke(new MethodInvoker(delegate { label10.Text = name + " is in weight"; }));
                        Data1 += name + " is in weight ";
                        int a = int.Parse(Save.Peek().ToString());
                        Save.Dequeue();
                        string b = Save.Peek().ToString();
                        Save.Dequeue();
                        Data1 += " - Weight is: " + b;
                        string c = Save.Peek().ToString();
                        Save.Dequeue();
                        Data1 += " - Tax is: " + c;
                        Query = "Update [sheet2$] SET weight_p = '" + b + "', tax = '" + c + "' Where passport = " + a;
                        UpdateBooking = new OleDbCommand(Query, con1);
                        UpdateBooking.ExecuteNonQuery();
                        con1.Close();

                        this.Invoke(new MethodInvoker(delegate { this.fw.listBox1.Items.Add(Data1); }));
                        gen.EventLog(Data1, "Weight");
                        LogID++;
                        Data1 = "";
                    }
                    mutWeight.ReleaseMutex();
                    Current_Passenger = 1;
                    if (this.Thread_Flight.ThreadState.ToString().ToLower() == "unstarted")
                    {
                        //Run Thread_Display
                        this.Thread_Flight.Start();
                    }
                    else
                    {
                        this.Thread_Flight.Resume();
                    }
                    this.Thread_Display.Resume();
                    this.Thread_Weight.Suspend();


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Weight");
                MessageBox.Show(ex.Message);
            }
        }

        private void Thread_Flight_Delegate()
        {
            //dissply data in flight
            while (true)
            {
                for (int i = 0; i < Passenger_Count; i++)
                {
                    Thread.Sleep(int.Parse(txtSleepCounter.Text));
                    mutFlight.WaitOne();
                    string FlightNumber = SaveFlight.Peek().ToString();
                    SaveFlight.Dequeue();
                    string Data2 = SaveFlight.Peek().ToString();
                    SaveFlight.Dequeue();
                    this.Invoke(new MethodInvoker(delegate { this.ff.Fill(FlightNumber, Data2); }));
                    mutFlight.ReleaseMutex();
                }
                Thread_Flight.Suspend();
            }
        }
    }
}
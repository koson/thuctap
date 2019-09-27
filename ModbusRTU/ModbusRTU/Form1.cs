using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Timers;


namespace ModbusRTU
{
    public partial class Form1 : Form
    {
        int button = 0;
        modbus mb = new modbus();
        SerialPort sp = new SerialPort();
        System.Timers.Timer timer = new System.Timers.Timer();
        string dataType;
        bool isPolling = false;
        int pollCount;

        #region GUI Delegate Declarations
        public delegate void GUIDelegate(string paramString);
        public delegate void GUIClear();
        public delegate void GUIStatus(string paramString);
        #endregion

        public Form1()
        {
            InitializeComponent();
            LoadListboxes();
            timer.Elapsed += new ElapsedEventHandler(timer_Elapsed);

        }

        #region Delegate Functions
        public void DoGUIClear()
        {
            if (this.InvokeRequired)
            {
                GUIClear delegateMethod = new GUIClear(this.DoGUIClear);
                this.Invoke(delegateMethod);
            }
            else
                this.listBox1.Items.Clear();
        }
        public void DoGUIStatus(string paramString)
        {
            if (this.InvokeRequired)
            {
                GUIStatus delegateMethod = new GUIStatus(this.DoGUIStatus);
                this.Invoke(delegateMethod, new object[] { paramString });
            }
            else
                this.toolStripStatusLabel1.Text = paramString;
        }
        public void DoGUIUpdate(string paramString)
        {
            if (this.InvokeRequired)
            {
                GUIDelegate delegateMethod = new GUIDelegate(this.DoGUIUpdate);
                this.Invoke(delegateMethod, new object[] { paramString });
            }
            else
                this.listBox1.Items.Add(paramString);
        }
        #endregion

        #region Timer Elapsed Event Handler
        void timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (button == 1)
            {
                PollFunction();
            }
            if (button == 2)
            {
                PollFunction2();
            }


        }
        #endregion

        #region Load Listboxes
        private void LoadListboxes()
        {
            //Three to load - ports, baudrates, datetype.  Also set default textbox values:
            //1) Available Ports:
            string[] ports = SerialPort.GetPortNames();

            foreach (string port in ports)
            {
                comboBox1.Items.Add(port);
            }

            comboBox1.SelectedIndex = 0;

            //2) Baudrates:
            string[] baudrates = { "230400", "115200", "57600", "38400", "19200", "9600" };

            foreach (string baudrate in baudrates)
            {
                comboBox3.Items.Add(baudrate);
            }
            comboBox3.SelectedIndex = 5;
            //3) Datatype:
            string[] dataTypes = { "Decimal", "Hexadecimal" };

            foreach (string dataType in dataTypes)
            {
                comboBox2.Items.Add(dataType);
            }

            comboBox2.SelectedIndex = 1;

            //Textbox defaults:
            textBox3.Text = "20";
            textBox4.Text = "1000";
            textBox1.Text = "1";
            textBox2.Text = "0";

        }
        #endregion

        #region Start and Stop Procedures
        private void StartPoll()
        {
            pollCount = 0;

            //Open COM port using provided settings:
            if (mb.Open(comboBox1.SelectedItem.ToString(), Convert.ToInt32(comboBox3.SelectedItem.ToString()),
                8, Parity.None, StopBits.One))
            {
                //Disable double starts:
                button1.Enabled = false;
                button3.Enabled = false;

                dataType = comboBox2.SelectedItem.ToString();

                //Set polling flag:
                isPolling = true;

                //Start timer using provided values:
                timer.AutoReset = true;
                if (textBox4.Text != "")
                    timer.Interval = Convert.ToDouble(textBox4.Text);
                else
                    timer.Interval = 1000;
                timer.Start();
            }

            toolStripStatusLabel1.Text = mb.modbusStatus;
        }
        private void StopPoll()
        {
            //Stop timer and close COM port:
            try
            {
                isPolling = false;
                timer.Stop();
                button1.Enabled = true;
                button3.Enabled = true;
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }

            toolStripStatusLabel1.Text = mb.modbusStatus;
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            StartPoll();
            button = 1;

        }
        private void Button2_Click_1(object sender, EventArgs e)
        {

            StopPoll();
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            StartPoll();
            button = 2;
        }
        #endregion




        #region Poll Function
        private void PollFunction()
        {
            //Update GUI:
            //DoGUIClear();
            pollCount++;
            DoGUIStatus("Poll count: " + pollCount.ToString());

            //Create array to accept read values:
            short[] values = new short[Convert.ToInt32(textBox3.Text)];
            ushort pollStart;
            ushort pollLength;

            if (textBox2.Text != "")
                pollStart = Convert.ToUInt16(textBox2.Text);
            else
                pollStart = 0;
            if (textBox3.Text != "")
                pollLength = Convert.ToUInt16(textBox3.Text);
            else
                pollLength = 20;

            //Read registers and display data in desired format:
            try
            {
                while (!mb.SendFc3(Convert.ToByte(textBox1.Text), pollStart, pollLength, ref values)) ;
            }
            catch (Exception err)
            {
                DoGUIStatus("Error in modbus read: " + err.Message);
            }

            string itemString;

            switch (dataType)
            {
                case "Decimal":
                    for (int i = 0; i < pollLength; i++)
                    {
                        itemString = "[" + Convert.ToString(pollStart + i + 40001) + "] , MB[" +
                            Convert.ToString(pollStart + i) + "] = " + values[i].ToString();
                        DoGUIUpdate(itemString);
                    }
                    break;
                case "Hexadecimal":
                    for (int i = 0; i < pollLength; i++)
                    {
                        itemString = "[" + Convert.ToString(pollStart + i + 40001) + "] , MB[" +
                            Convert.ToString(pollStart + i) + "] = " + values[i].ToString("X");
                        DoGUIUpdate(itemString);
                    }
                    break;

            }
        }
        #endregion

        #region PollFunction 2
        private void PollFunction2()
        {
            //Update GUI:
            //DoGUIClear();
            pollCount++;
            DoGUIStatus("Poll count: " + pollCount.ToString());

            //Create array to accept read values:
            short[] values = new short[Convert.ToInt32(textBox3.Text)];
            ushort pollStart;
            ushort pollLength;

            if (textBox2.Text != "")
                pollStart = Convert.ToUInt16(textBox2.Text);
            else
                pollStart = 0;
            if (textBox3.Text != "")
                pollLength = Convert.ToUInt16(textBox3.Text);
            else
                pollLength = 20;

            //Read registers and display data in desired format:
            try
            {
                while (!mb.SendFc1(Convert.ToByte(textBox1.Text), pollStart, pollLength, ref values)) ;
            }
            catch (Exception err)
            {
                DoGUIStatus("Error in modbus read: " + err.Message);
            }

            string itemString;
            Int32 poll;
            if (pollLength % 8 != 0)
            {
                poll = pollLength / 8 + 1;
            }
            else
            {
              poll = pollLength/8;
            }
            switch (dataType)
            {
                case "Decimal":
                    for (int i = 0; i < pollStart; i++)
                    {
                        itemString = "[" + Convert.ToString(pollStart + 1 + i*8) + "-" + Convert.ToString(pollStart + (i+1) * 8) + "] , MB[" +
                            Convert.ToString(pollStart + i) + "] = " + values[i].ToString();
                        DoGUIUpdate(itemString);
                    }
                    break;
                case "Hexadecimal":
                    for (int i = 0; i < poll; i++)
                    {
                        itemString = "[" + Convert.ToString(pollStart + 1 + i * 8) + "-" + Convert.ToString(pollStart + (i + 1) * 8) + "] , MB[" +
                            Convert.ToString(i) + "] = " + values[i].ToString("X");
                        DoGUIUpdate(itemString);
                    }
                    break;
            }
        }
        #endregion
        #region Write Function
        //private void WriteFunction()
        //{
        //    //StopPoll();

        //    if (txtWriteRegister.Text != "" && txtWriteValue.Text != "" && textBox1.Text != "")
        //    {
        //        byte address = Convert.ToByte(textBox1.Text);
        //        ushort start = Convert.ToUInt16(txtWriteRegister.Text);
        //        short[] value = new short[1];
        //        value[0] = Convert.ToInt16(txtWriteValue.Text);

        //        try
        //        {
        //            while (!mb.SendFc16(address, start, (ushort)1, value)) ;
        //        }
        //        catch (Exception err)
        //        {
        //            DoGUIStatus("Error in write function: " + err.Message);
        //        }
        //        DoGUIStatus(mb.modbusStatus);
        //    }
        //    else
        //        DoGUIStatus("Enter all fields before attempting a write");

        //    //StartPoll();
        //}
        //private void btnWrite_Click(object sender, EventArgs e)
        //{
        //    WriteFunction();
        //}
        //#endregion

        //#region Data Type Event Handler
        //private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    //restart the data poll if datatype is changed during the process:
        //    if (isPolling)
        //    {
        //        StopPoll();
        //        dataType = comboBox2.SelectedItem.ToString();
        //        StartPoll();
        //    }

        //}
        #endregion



    }
}
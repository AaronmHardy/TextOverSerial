using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Runtime.InteropServices;

namespace TextOverSerial
{
    public partial class Form1 : Form
    {
        // Get idle time from .NET
        // PInvoke GetLastInputInfo
        [DllImport("User32.dll")]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
        // Create a struct to solve a problem writing to data type LASTINPUTINFO
        internal struct LASTINPUTINFO
        {
            public uint cbSize;
            public uint dwTime;
        }
        // Method to call LASTINPUTINFO
        public static uint GetIdleTime()
        {
            LASTINPUTINFO lastInPut = new LASTINPUTINFO();
            lastInPut.cbSize = (uint)System.Runtime.InteropServices.Marshal.SizeOf(lastInPut);
            GetLastInputInfo(ref lastInPut);
            return ((uint)Environment.TickCount - lastInPut.dwTime);
        }

        //Create timer to run UpdateText method
        private System.Windows.Forms.Timer timer1;
        public void InitTimer()
        {
            timer1 = new System.Windows.Forms.Timer();
            timer1.Tick += new EventHandler(Timer1_Tick); // Event called when timer reaches 0
            timer1.Interval = 15000; // Sets interval in ms, 1000ms = 1 second
            timer1.Start(); // Start the timer
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            UpdateText(); // Runs UpdateText when timer reaches 0
        }

		// Initialize variables on startup
        public Form1()
        {
            InitializeComponent();
            label2.Hide();
            checkBox1.Checked = false;
            textBox3.Text = " Away message here";
            textBox3.Enabled = false;
            numericUpDown2.Enabled = false;
        }

        // Used when updating away message or date & time
        private void UpdateText()
        {
            if (DateTimeBox1 == true)
            {
                textBox1.Text = DateTime.Now.ToString("ddd'/'dd'/'MMM hh':'mmtt"); // Formatted "Tue/04/Sept 09:35AM"
            }
            if (DateTimeBox2 == true)
            {
                textBox2.Text = DateTime.Now.ToString("ddd'/'dd'/'MMM hh':'mmtt");
            }
            // Away message
            if (checkBox1.Checked == true && GetIdleTime() > numericUpDown2.Value * 60000)
            {
                SendData(textBox1.Text, textBox3.Text); // Writes away message
            }
            else
            {
                SendData(textBox1.Text, textBox2.Text); // Clears away message
            }
        }

        private void SendData(string firstLine, string secondLine)
        {
            // Instantiate serial port and open for communication
            SerialPort port = new SerialPort("COM" + numericUpDown1.Value.ToString(), int.Parse(comboBox1.SelectedItem.ToString()), Parity.None, 8, StopBits.One);

            try // Open the port
            {
                port.Open();
            }
            catch // Throw message to user if port is unable to open, then return to exit the method
            {
                MessageBox.Show("Cannot open serial port "+ numericUpDown1.Value.ToString());
                return;
            }

            // Write a set of bytes to clear existing data and return to line 0 char 0
            port.Write(new byte[] { 31, 13 }, 0, 2);

            // Write first line, return, new line, write second line
            port.Write(firstLine);
            port.Write("\r\n");
            port.Write(secondLine);

            // Housekeeping
            port.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SendData(textBox1.Text, textBox2.Text);

            // If datetime or away message is set
            if (DateTimeBox1 == true || DateTimeBox2 == true || checkBox1.Checked == true)
            {
                InitTimer();
            }

            // Let the user know text was sent
            label2.Show();
        }
        // Button to hide the window
        private void button2_Click(object sender, EventArgs e)
        {
            Hide();
            notifyIcon1.Visible = true;
            notifyIcon1.ShowBalloonTip(800);
        }
        // Button to unhide the window
        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            this.WindowState = FormWindowState.Normal;
            notifyIcon1.Visible = false;
        }
        // Minimize to system tray
        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                Hide();
                notifyIcon1.Visible = true;
            }
        }
        // Toggle for date & time buttons
        bool DateTimeBox1 = false; // Boolean value used to update text
        private void button3_Click(object sender, EventArgs e)
        {
            if (DateTimeBox1 == true) // If on toggle off
            {
                DateTimeBox1 = false; // do not update text
                textBox1.Text = "";
                textBox1.Enabled = true;
            }
            else // If off toggle on
            {
                DateTimeBox1 = true; // update text
                textBox1.Text = DateTime.Now.ToString("ddd'/'dd'/'MMM hh':'mmtt");
                textBox1.Enabled = false;
            }
        }

        bool DateTimeBox2 = false; // Boolean value used to update text
        private void button4_Click(object sender, EventArgs e)
        {
            if (DateTimeBox2 == true) // If on toggle off
            {
                DateTimeBox2 = false; // do not update text
                textBox2.Text = "";
                textBox2.Enabled = true;
            }
            else // If off toggle on
            {
                DateTimeBox2 = true; // update text
                textBox2.Text = DateTime.Now.ToString("ddd'/'dd'/'MMM hh':'mmtt");
                textBox2.Enabled = false;
            }
        }
        // Check box to enable / disable away message
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == false)
            {
                textBox3.Text = " Away message here";
                textBox3.Enabled = false;
                label3.Enabled = false;
                numericUpDown2.Enabled = false;
            }
            // Disable textbox / disable away message
            else
            {
                textBox3.Text = "";
                textBox3.Enabled = true;
                label3.Enabled = true;
                numericUpDown2.Enabled = true;
            }
        }
    }
}

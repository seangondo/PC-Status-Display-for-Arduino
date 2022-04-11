using System;
using System.Management;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using OpenHardwareMonitor.Hardware;
using System.IO.Ports;
using System.Globalization;
using Timer = System.Windows.Forms.Timer;
using System.Threading;

namespace ImportDLLfiles
{
    public partial class Form1 : Form
    {
        bool serialSend = false;
        String portUse = "";
        String portNameUse = "";
        String CPU1 = "";
        String CPU2 = "";
        String CPU3 = "";
        String CPU4 = "";

        String GPU1 = "";
        String GPU2 = "";
        String GPU3 = "";
        String GPU4 = "";
        String GPU5 = "";

        UpdateVisitor updateVisitor = new UpdateVisitor();
        Computer thisComputer;
        public Form1()
        {
            InitializeTimer();
            InitializeComponent();
            thisComputer = new Computer();

            using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'"))
            {
                var portnames = SerialPort.GetPortNames();
                var ports = searcher.Get().Cast<ManagementBaseObject>().ToList().Select(p => p["Caption"].ToString());
                var portList = portnames.Select(n => n + " - " + ports.FirstOrDefault(s => s.Contains(n))).ToList();

                foreach (string s in portList)
                {
                    comboBox1.Items.Add(s.Substring(7));
                }
            }

            //string[] ports = SerialPort.GetPortNames();
            /*foreach (string port in ports)
            {
                comboBox1.Items.Add(port);
                portUse = port;
            }*/
        }
        public class UpdateVisitor : IVisitor
        {
            public void VisitComputer(IComputer computer)
            {
                computer.Traverse(this);
            }
            public void VisitHardware(IHardware hardware)
            {
                hardware.Update();
                foreach (IHardware subHardware in hardware.SubHardware) subHardware.Accept(this);
            }
            public void VisitSensor(ISensor sensor) { }
            public void VisitParameter(IParameter parameter) { }
        }

        private void InitializeTimer()
        {
            timer1 = new System.Windows.Forms.Timer();
            timer1.Interval = 1000;
            timer1.Enabled = true;
            timer1.Start();
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            mainProgram();
            if(serialPort1.IsOpen)
            {
                 serialPort1.WriteLine(CPU4 + CPU2 + CPU1 + CPU3);
                 //Thread.Sleep(100);
                 serialPort1.WriteLine(GPU5 + GPU2 + GPU1 + GPU3 + GPU4);
                /*serialPort1.WriteLine(CPU4);
                Thread.Sleep(100);
                serialPort1.WriteLine(CPU2);
                Thread.Sleep(100);
                serialPort1.WriteLine(CPU1);
                Thread.Sleep(100);
                serialPort1.WriteLine(CPU3);
                Thread.Sleep(100);
                serialPort1.WriteLine(GPU5);
                Thread.Sleep(100);
                serialPort1.WriteLine(GPU2);
                Thread.Sleep(100);
                serialPort1.WriteLine(GPU1);
                Thread.Sleep(100);
                serialPort1.WriteLine(GPU3);
                Thread.Sleep(100);
                serialPort1.WriteLine(GPU4);*/
                label11.Text = CPU4 + CPU2 + CPU1 + CPU3 + GPU5 + GPU2 + GPU1 + GPU3 + GPU4;
                //serialPort1.WriteLine(GPU5 + GPU2 + GPU1 + GPU3 + GPU4);
            }
        }

        private void mainProgram()
        {
            thisComputer.Open();
            thisComputer.CPUEnabled = true;
            thisComputer.GPUEnabled = true;
            thisComputer.Accept(updateVisitor);
            for (int i = 0; i < thisComputer.Hardware.Length; i++)
            {
                if (thisComputer.Hardware[i].HardwareType == HardwareType.CPU)
                {
                    for (int j = 0; j < thisComputer.Hardware[i].Sensors.Length; j++)
                    {
                        if (thisComputer.Hardware[i].Sensors[j].SensorType == SensorType.Temperature)
                        {
                            if (thisComputer.Hardware[i].Sensors[j].Name == "CPU Package")
                            {
                                float cpuTemp = (float)thisComputer.Hardware[i].Sensors[j].Value;
                                if (serialSend == true)
                                {
                                    CPU1 = Math.Round(cpuTemp, 2, MidpointRounding.ToEven).ToString() + "CPU;";
                                    CPU4 = thisComputer.Hardware[i].Name + "CAB;";
                                    //serialPort1.WriteLine(thisComputer.Hardware[i].Name + "CNM");
                                }
                                label6.Text = thisComputer.Hardware[i].Name;
                                label1.Text = thisComputer.Hardware[i].Sensors[j].Name + "\t : \t" + Math.Round(cpuTemp, 2, MidpointRounding.ToEven).ToString() + "\r";
                            }
                        }
                        if (thisComputer.Hardware[i].Sensors[j].SensorType == SensorType.Load)
                        {
                            CPU2 = Math.Round((float)thisComputer.Hardware[i].Sensors[j].Value, 2, MidpointRounding.ToEven).ToString() + "UCP;";
                            label3.Text = thisComputer.Hardware[i].Sensors[j].Name + "\t : \t" + Math.Round((float)thisComputer.Hardware[i].Sensors[j].Value, 2, MidpointRounding.ToEven).ToString() + " %";
                        }
                        if (thisComputer.Hardware[i].Sensors[j].SensorType == SensorType.Clock)
                        {
                            CPU3 = Math.Round((float)thisComputer.Hardware[i].Sensors[j].Value, 2, MidpointRounding.ToEven).ToString() + "CCP;";
                            label8.Text = thisComputer.Hardware[i].Sensors[j].Name + "\t : \t" + Math.Round((float)thisComputer.Hardware[i].Sensors[j].Value, 0, MidpointRounding.ToEven).ToString() + " MHz";
                        }
                    }
                }
            }
            for (int i = 0; i < thisComputer.Hardware.Length; i++)
            {
                if (thisComputer.Hardware[i].HardwareType == HardwareType.GpuNvidia)
                {
                    for (int j = 0; j < thisComputer.Hardware[i].Sensors.Length; j++)
                    {
                        if (thisComputer.Hardware[i].Sensors[j].SensorType == SensorType.Temperature)
                        {
                            if (serialSend == true)
                            {
                                GPU1 = thisComputer.Hardware[i].Sensors[j].Value.ToString() + "GPU;";
                                GPU5 = thisComputer.Hardware[i].Name.Substring(22) + "GNM;";
                                //serialPort1.WriteLine(thisComputer.Hardware[i].Name.Substring(22) + "GNM");
                            }
                            label7.Text = thisComputer.Hardware[i].Name.Substring(7);
                            label2.Text = thisComputer.Hardware[i].Sensors[j].Name + "\t : \t" + thisComputer.Hardware[i].Sensors[j].Value.ToString() + "\r";
                        }
                        if (thisComputer.Hardware[i].Sensors[j].SensorType == SensorType.Load)
                        {
                            if (thisComputer.Hardware[i].Sensors[j].Name == "GPU Core")
                            {
                                GPU2 = Math.Round((float)thisComputer.Hardware[i].Sensors[j].Value, 2, MidpointRounding.ToEven).ToString() + "UGP;";
                                label4.Text = thisComputer.Hardware[i].Sensors[j].Name + "\t : \t" + Math.Round((float)thisComputer.Hardware[i].Sensors[j].Value, 2, MidpointRounding.ToEven).ToString() + " %";
                            }
                        }
                        if (thisComputer.Hardware[i].Sensors[j].SensorType == SensorType.SmallData)
                        {
                            if (thisComputer.Hardware[i].Sensors[j].Name == "GPU Memory Used")
                            {
                                GPU3 = Math.Round((float)thisComputer.Hardware[i].Sensors[j].Value, 2, MidpointRounding.ToEven).ToString() + "GUS;";
                                label9.Text = thisComputer.Hardware[i].Sensors[j].Name + "\t : \t" + Math.Round((float)thisComputer.Hardware[i].Sensors[j].Value, 2, MidpointRounding.ToEven).ToString() + " MB";
                            }
                            if (thisComputer.Hardware[i].Sensors[j].Name == "GPU Memory Total")
                            {
                                GPU4 = Math.Round((float)thisComputer.Hardware[i].Sensors[j].Value, 2, MidpointRounding.ToEven).ToString() + "GMX;";
                                label10.Text = thisComputer.Hardware[i].Sensors[j].Name + "\t : \t" + Math.Round((float)thisComputer.Hardware[i].Sensors[j].Value, 2, MidpointRounding.ToEven).ToString() + " MB";
                            }
                        }
                    }
                }
            }
            thisComputer.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!serialPort1.IsOpen)
            {
                if(portNameUse != "")
                {
                    serialSend = true;
                    label5.Text = "Serial Working on Port : " + portNameUse;
                    serialPort1.PortName = portNameUse;
                    serialPort1.Open();
                }
                else
                {
                    label5.Text = "Please select port first!";
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(serialPort1.IsOpen)
            {
                serialSend = false;
                label5.Text = "Serial Not Sending Any Message";
                //serialPort1.WriteLine("CPUGPU\n");
                serialPort1.Close();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            String selectedPort = comboBox1.SelectedItem.ToString();
            portNameUse = selectedPort.Substring(selectedPort.Length - 5, 4);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            comboBox1.Text = "Select port...";
            using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'"))
            {
                var portnames = SerialPort.GetPortNames();
                var ports = searcher.Get().Cast<ManagementBaseObject>().ToList().Select(p => p["Caption"].ToString());
                var portList = portnames.Select(n => n + " - " + ports.FirstOrDefault(s => s.Contains(n))).ToList();

                foreach (string s in portList)
                {
                    comboBox1.Items.Add(s.Substring(7));
                }
            }
        }
    }
}

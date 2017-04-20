using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

// Needed for remoting
using System.Runtime.Remoting;
using System.Runtime.Remoting.Channels;
using System.Runtime.Remoting.Channels.Tcp;

using Com.QuantAsylum;
using System.Diagnostics;

// To remotely control the QA400 application
// 1. Add reference to the QA400 application EXE. 
// 2. add "using Com.QuantAsylum;" line above
// 3. Add reference to System.Runtime.Remoting
// 4. Add "using" statements above for remoting

namespace RemotingTest
{
    public partial class Form1 : Form
    {
        QA401Interface AudioAnalyzer;

        public Form1()
        {
            InitializeComponent();

            // Try to connect to QA400 Application. Note the code below is boilerplate, likely needed by any app that wants to connect
            // to the QA400 application. This is routine dotnet remoting code. 
            try  
            {
                TcpChannel tcpChannel = new TcpChannel();
                ChannelServices.RegisterChannel(tcpChannel, false);

                Type requiredType = typeof(QA401Interface);

                AudioAnalyzer = (QA401Interface)Activator.GetObject(requiredType, "tcp://localhost:9401/QuantAsylumQA401Server");
                LogData("QA401 Setup OK");
            }
            catch
            {
                // If the above fails for any reason, make sure the rest of the app can tell. We do that here by setting AudioAnalyzer to null.
                AudioAnalyzer = null;
                LogData("QA401 Setup failed");
            }

        }

        // Routine to log data to textbox
        private void LogData(string s)
        {
            textBox1.AppendText(s + Environment.NewLine);
        }

        // Get Name button. Retrieves the name of the equipment. Also an easy check if the remoting is working and able to talk to the host app. As we expect
        // this will be the first function called after a remoting setup, it's protected by try/catch. The subsequent calls will only check if AudioAnalyzer is
        // null. In other user code, after the setup code succeeds, use a call to GetName() to ensure the connection is working.
        private void button1_Click(object sender, EventArgs e)
        {
            if (AudioAnalyzer == null) return;

            try
            {
                Stopwatch sw = Stopwatch.StartNew();

                string name = AudioAnalyzer.GetName();

                sw.Stop();

                LogData(string.Format("{0}   [0:0.0 mS]", name, sw.Elapsed.TotalMilliseconds));
            }
            catch (Exception ex)
            {
                // If we end up here, then it means the connection to the analyzer failed. 
                LogData(ex.Message);
                AudioAnalyzer = null;
            }
        }

        // Start Running button. Just like pushing the Run/Stop button the app
        private void StartRunning_Click(object sender, EventArgs e)
        {
            if (AudioAnalyzer == null) return;

            AudioAnalyzer.Run();

        }

        // Stop Running button. Just like pushing the Run/Stop button the app
        private void StopRunning(object sender, EventArgs e)
        {
            if (AudioAnalyzer == null) return;

            AudioAnalyzer.Stop();
        }

        // Run Single
        private void button6_Click(object sender, EventArgs e)
        {
            if (AudioAnalyzer == null) return;

            AudioAnalyzer.RunSingle();
        }

        // Get Data button. Grabs the last captured buffers of data.
        private void GetFreqData_Click(object sender, EventArgs e)
        {
            if (AudioAnalyzer == null) return;

            QA401.PointD[] data = AudioAnalyzer.GetData(QA401.ChannelType.LeftIn);

            // Note, data can be null if nothing collected yet
            if (data != null)
                LogData("Retrieved " + data.Length + " points");
        }

        // Grabs the time data captured by the QA400
        private void GetTimeData_Click(object sender, EventArgs e)
        {
            if (AudioAnalyzer == null) return;

            QA401.PointD[] data = AudioAnalyzer.GetTimeData(QA401.ChannelType.LeftIn);

            // Note, data can be null if nothing collected yet
            if (data != null)
                LogData("Retrieved " + data.Length + " points");
        }

        // Run and Compute Power button. Runs single, waits for acq to finish, and then requests host app to computes total power on returned data
        private void ComputePower_Click(object sender, EventArgs e)
        {
            if (AudioAnalyzer == null) return;

            AudioAnalyzer.RunSingle();
            while (AudioAnalyzer.GetAcquisitionState() == QA401.AcquisitionState.Busy)
            {
                Console.WriteLine("Busy");
            }

            QA401.PointD[] data = AudioAnalyzer.GetData(QA401.ChannelType.LeftIn);
            LogData("Power: " + AudioAnalyzer.ComputePowerDB(data));
        }

        // Set Gen1 Button
        private void button7_Click(object sender, EventArgs e)
        {
            AudioAnalyzer.SetGenerator(QA401.GenType.Gen1, true, -10, 1000);
            AudioAnalyzer.RunSingle();
        }

        // Set Gen1 Off button
        private void button10_Click(object sender, EventArgs e)
        {
            AudioAnalyzer.SetGenerator(QA401.GenType.Gen1, false, -10, 1000);
            AudioAnalyzer.RunSingle();
        }

        // Compute THD
        private void button8_Click(object sender, EventArgs e)
        {
            if (AudioAnalyzer == null) return;

            AudioAnalyzer.RunSingle();
            while (AudioAnalyzer.GetAcquisitionState() == QA401.AcquisitionState.Busy)
            {
                Console.WriteLine("Busy");
            }

            QA401.PointD[] data = AudioAnalyzer.GetData(QA401.ChannelType.LeftIn);
            LogData("THD %: " + AudioAnalyzer.ComputeTHDPct(data, 1000, 20000));
        }

        // Compute THDN
        private void button9_Click(object sender, EventArgs e)
        {
            if (AudioAnalyzer == null) return;

            AudioAnalyzer.RunSingle();
            while (AudioAnalyzer.GetAcquisitionState() == QA401.AcquisitionState.Busy)
            {
                Console.WriteLine("Busy");
            }

            QA401.PointD[] data = AudioAnalyzer.GetData(QA401.ChannelType.LeftIn);
            LogData("THD+N %: " + AudioAnalyzer.ComputeTHDNPct(data, 1000, 20, 20000));
        }

        // Find peak dBV
        private void button11_Click(object sender, EventArgs e)
        {
            if (AudioAnalyzer == null) return;

            AudioAnalyzer.SetUnits(QA401.UnitsType.dBV);

            AudioAnalyzer.RunSingle();
            while (AudioAnalyzer.GetAcquisitionState() == QA401.AcquisitionState.Busy)
            {
                Console.WriteLine("Busy");
            }

            QA401.PointD[] data = AudioAnalyzer.GetData(QA401.ChannelType.LeftIn);
            double peak = AudioAnalyzer.ComputePeakPowerDB(data);
            LogData("Peak: " + peak.ToString("0.00"));
        }

        // Find peak dBFS
        private void button12_Click(object sender, EventArgs e)
        {
            if (AudioAnalyzer == null) return;

            AudioAnalyzer.SetUnits(QA401.UnitsType.dBFS);

            AudioAnalyzer.RunSingle();
            while (AudioAnalyzer.GetAcquisitionState() == QA401.AcquisitionState.Busy)
            {
                Console.WriteLine("Busy");
            }

            QA401.PointD[] data = AudioAnalyzer.GetData(QA401.ChannelType.LeftIn);
            double peak = AudioAnalyzer.ComputePeakPowerDB(data);
            LogData("Peak: " + peak.ToString("0.00"));
        }

        private void button14_Click(object sender, EventArgs e)
        {
            AudioAnalyzer.GenerateTone(-20, 440, 3000);
        }

        private void button15_Click(object sender, EventArgs e)
        {
            // Active 20 dB Atten
            AudioAnalyzer.SetInputAtten(QA401.InputAttenState.dB20);
        }

        private void button16_Click(object sender, EventArgs e)
        {
            // De-activatee 20 dB Atten
            AudioAnalyzer.SetInputAtten(QA401.InputAttenState.NoAtten);
        }

       

       


    }
}

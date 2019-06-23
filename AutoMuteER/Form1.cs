using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using NAudio.CoreAudioApi;
using System.Security.Principal;
namespace AutoMuteER
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            bool elv;
            using (WindowsIdentity id = WindowsIdentity.GetCurrent())
            {
                WindowsPrincipal p = new WindowsPrincipal(id);
                elv = p.IsInRole(WindowsBuiltInRole.Administrator);
		if (!elv)
		{
			checkBox1.Enabled = false;
		}
            }
            
            timer1.Start();
            timer2.Interval = Properties.Settings.Default.muteTime * 1000;
            checkBox1.Checked = Properties.Settings.Default.loadWithWindows;
            numericUpDown1.Value = Properties.Settings.Default.maxVolume;
            numericUpDown3.Value = Properties.Settings.Default.muteTime;

            var devices = new MMDeviceEnumerator().EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active);
            if (devices != null)
            {
                comboBox1.Items.AddRange(devices.ToArray());
            }
            comboBox1.SelectedIndex = Properties.Settings.Default.lastDevice;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                var device = (MMDevice)comboBox1.SelectedItem;
                int val = (int)Math.Round(device.AudioMeterInformation.MasterPeakValue * 100);
                progressBar1.Value = val;
                if (val >= Properties.Settings.Default.maxVolume)
                {
                    label1.Text = "Mic Status: Muted";
                    device.AudioEndpointVolume.Mute = true;
                    timer2.Start();
                }
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            var device = (MMDevice)comboBox1.SelectedItem;
            device.AudioEndpointVolume.Mute = false;
            label1.Text = "Mic Status: Unmuted";
            timer2.Stop();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.lastDevice = comboBox1.SelectedIndex;
            Properties.Settings.Default.Save();
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            timer2.Interval = (int)numericUpDown3.Value * 1000;
            Properties.Settings.Default.muteTime = (int)numericUpDown3.Value;
            Properties.Settings.Default.Save();
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.maxVolume = (int)numericUpDown1.Value;
            Properties.Settings.Default.Save();
        }
		
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
	    using (RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            {
		if (checkBox1.Checked)
		{
			key.SetValue("AutoMuteER", "\"" + Application.ExecutablePath + "\"");
			Properties.Settings.Default.loadWithWindows = true;
			Properties.Settings.Default.Save();
		}
		else
		{
			key.DeleteValue("AutoMuteER", false);
			Properties.Settings.Default.loadWithWindows = false;
			Properties.Settings.Default.Save();
		}
	    }
        }
		
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = false;
                this.Hide();
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
            Form1 frm = new Form1();
            frm.Show();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == this.WindowState)
            {
                notifyIcon1.Visible = true;
		notifyIcon1.Icon = SystemIcons.Information;

		notifyIcon1.BalloonTipText = "Minimized";
		notifyIcon1.BalloonTipTitle = "The application is running in the background.";
		notifyIcon1.ShowBalloonTip(500);
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}

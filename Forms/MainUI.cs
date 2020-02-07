using Be.Windows.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Media;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace STC_ISP_NG
{
    public partial class MainUI : Form
    {
        private FileByteProvider byteProvider;
        public MainUI()
        {
            InitializeComponent();
            if (!Directory.Exists("./temp"))
            {
                Directory.CreateDirectory("./temp");
            }
            hexBoxPRG.Refresh();
            comboBoxProtocol.SelectedIndex = 8;
            //comboBoxSerial.SelectedIndex = 0;
            comboBoxSpeed.SelectedIndex = 0;
            comboBoxTrim.SelectedIndex = 0;
            MessageBox.Show("EEPROM在下次下载用户程序时默认是不擦除的\r\n为提高用户程序安全性,建议用户先择下次下载时擦除EEPROM,\r\n即勾选上\"下次下载用户程序时擦除用户EEPROM区\"选项","温馨提示",MessageBoxButtons.OK,MessageBoxIcon.Information);
            updateSerial();
            if (comboBoxSerial.Items.Count>0)
            {
                comboBoxSerial.SelectedIndex = 0;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
        }

        private void button3_Click(object sender, EventArgs e)
        {
            consoleControl1.StopProcess();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }
        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBoxSerial.SelectedItem==null)
            {
                MessageBox.Show("请在串口选项中选择单片机所在串口");
                return;
            }
            button5.Enabled = false;
            button4.Enabled = false;
            consoleControl1.StartProcess("./runtime/python.exe", $"./runtime/stcgal/__main__.py "+ generateFlashCommand(textBoxHex.Text,
                comboBoxProtocol.SelectedItem.ToString(),
                comboBoxSerial.SelectedItem.ToString(),
                comboBoxSpeed.SelectedItem.ToString(),
                comboBoxTrim.SelectedItem.ToString(),
                checkBoxEraseEEPROM.Checked));
            while (consoleControl1.IsProcessRunning)
            {
                Application.DoEvents();
                Thread.Sleep(15);
            }
            if (checkBox1.Checked) {
                using (var soundPlayer = new SoundPlayer(@".\assets\1.wav"))
                {
                    soundPlayer.Play(); // can also use soundPlayer.PlaySync()
                }
            }
            button5.Enabled = true;
            button4.Enabled = true;
        }
        private void updateSerial() {
            string[] ports = SerialPort.GetPortNames();
            comboBoxSerial.Items.Clear();
            foreach (string name in ports)
            {
                comboBoxSerial.Items.Add(name);
            }
        }

        private void comboBoxSerial_MouseClick(object sender, MouseEventArgs e)
        {
            updateSerial();
            if (comboBoxSerial.Items.Count > 0)
            {
                comboBoxSerial.SelectedIndex = 0;
            }
        }

        private void buttonHex_Click(object sender, EventArgs e)
        {
            if (byteProvider!=null)
            {
                byteProvider.Dispose();
            }
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Hex文件|*.hex|所有文件|*.*";
            if (fileDialog.ShowDialog()==DialogResult.OK)
            {
                File.Copy(fileDialog.FileName,"./temp/" + Path.GetFileName(fileDialog.FileName),true);
                byteProvider = new FileByteProvider("./temp/" + Path.GetFileName(fileDialog.FileName));
                hexBoxPRG.ByteProvider = byteProvider;
                textBoxHex.Text = fileDialog.FileName;
            }
            
        }
        public string GetHDDSerial()
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMedia");

            foreach (ManagementObject wmi_HD in searcher.Get())
            {
                // get the hardware serial no.
                if (wmi_HD["SerialNumber"] != null)
                    return wmi_HD["SerialNumber"].ToString();
            }

            return string.Empty;
        }
        private void buttonHardDriveID_Click(object sender, EventArgs e)
        {
            MessageBox.Show("本机磁盘号为:\r\n"+GetHDDSerial(),"STC-ISP NG");
        }
        public string generateFlashCommand(string romPath,string protocol, string port, string baud, string trim, bool eraseEEPROM) {
            string commandGenerated = $"-p {port} -b {baud}";
            if (trim!="0")
            {
                commandGenerated += $" -t {trim}";
            }
            if (protocol == "自动检测")
            {
                commandGenerated += $" -P auto";
            }
            commandGenerated += $" -o";

            if (eraseEEPROM)
            {
                commandGenerated += $" eeprom_erase_enabled=true";
            }
            else {
                commandGenerated += $" eeprom_erase_enabled=false";
            }

            commandGenerated += $" {romPath}";
            return commandGenerated;
        }
        private void button5_Click(object sender, EventArgs e)
        {
            if (comboBoxSerial.SelectedItem == null)
            {
                MessageBox.Show("请在串口选项中选择单片机所在串口");
                return;
            }
            button5.Enabled = false;
            button4.Enabled = false;
            string proto = "";
            if (comboBoxProtocol.SelectedItem.ToString()=="自动检测")
            {
                proto = "auto";
            }
            else
            {
                proto = comboBoxProtocol.SelectedItem.ToString();
            }
            consoleControl1.StartProcess("./runtime/python.exe", $"./runtime/stcgal/__main__.py -P {proto} -p {comboBoxSerial.SelectedItem.ToString()}");
            while (consoleControl1.IsProcessRunning)
            {
                Application.DoEvents();
                Thread.Sleep(15);
            }
            button5.Enabled = true;
            button4.Enabled = true;
        }

        private void consoleControl1_OnConsoleOutput(object sender, ConsoleControl.ConsoleEventArgs args)
        {
            consoleControl1.InternalRichTextBox.ScrollToCaret();
        }

        private void MainUI_FormClosing(object sender, FormClosingEventArgs e)
        {
            consoleControl1.StopProcess();
        }

        private void checkBoxEraseEEPROM_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void buttonClearCustom_Click(object sender, EventArgs e)
        {
            textBoxCustomCommands.Text = "";
        }
    }
}

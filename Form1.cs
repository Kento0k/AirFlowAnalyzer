using EasyModbus;
using System;
using System.IO.Ports;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Text.RegularExpressions;
using System.Media;

namespace AirFlowAnalyzer
{
    public partial class Form1 : Form
    {
        Queue<Measurement> measurements = new Queue<Measurement>();
        Measurement firstMeasurement = new Measurement();
        TimeSpan experimentTime = new TimeSpan(0, 0, 0);
        String comPort;
        bool stop, excelWrite, docExist = false, firstTempLeftFlag = true, firstHumLeftFlag = true, firstTempRightFlag = true, firstHumRightFlag = true, firstFRFlag = true, firstMeasurementFlag = true, firstMeasurmentDone = false;
        float sumTemperatureLeft = 0, sumHumidityLeft = 0, sumTemperatureRight = 0, sumHumidityRight = 0;
        float exTemperatureLeft, exHumidityLeft, exTemperatureRight, exHumidityRight, exFlowRate;
        float avrgTemperatureLeft, avrgHumidityLeft, avrgTemperatureRight, avrgHumidityRight;
        int tempLeftMeasureCnt = 0, humLeftMeasureCnt = 0, tempRightMeasureCnt = 0, humRightMeasureCnt = 0;
        int autoRowCnt = 2, manualRowCnt = 2, lastRow = 0;
        int baudrate = 9600;
        ExcelPackage excel;
        ExcelWorksheet sheet;
        FileInfo document; 
        
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        async void button1_Click(object sender, EventArgs e)
        {
            ModbusClient ivitLeft, ivitRight, flowMeter;
            float temperatureLeft = -1000, humidityLeft = -1000, temperatureRight = -1000, humidityRight = -1000, flowRate = 0;
            int holdingReg = 3, inputReg = 4;
            stop = false;
            button1.Enabled = false;
            button2.Enabled = true;

            if (firstMeasurementFlag)
            {
                firstMeasurement.timestamp = DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss");
                firstMeasurementFlag = false;
            }

            timer1.Start();
            timer2.Start();
            timer3.Start();
            while (!stop)
            {
                ivitLeft = await Modbus.Connect(comPort, 3, baudrate, Parity.None, StopBits.One, 100);
                if (ivitLeft.Connected)
                {
                    try
                    {
                        temperatureLeft = await Modbus.ReadRegisters(ivitLeft, 0x0022, 2, inputReg);
                        humidityLeft = await Modbus.ReadRegisters(ivitLeft, 0x0016, 2, inputReg);
                        pictureBox1.Image = Image.FromFile(Directory.GetCurrentDirectory() + @"\resources\images\green_circle.png");
                        pictureBox1.Refresh();
                    } catch (System.TimeoutException)
                    {
                        pictureBox1.Image = Image.FromFile(Directory.GetCurrentDirectory() + @"\resources\images\red_circle.png");
                        pictureBox1.Refresh();
                    }

                    while (ivitLeft.Connected)
                    {
                        ivitLeft.Disconnect();
                    }

                    if (temperatureLeft != -1000)
                    {
                        textBox10.Text = (Math.Round(temperatureLeft, 2)).ToString();
                        textBox10.Refresh();
                        exTemperatureLeft = temperatureLeft;
                        if (firstTempLeftFlag)
                        {
                            firstMeasurement.temperatureLeft = temperatureLeft;
                            firstTempLeftFlag = false;
                        }
                        sumTemperatureLeft += temperatureLeft;
                        tempLeftMeasureCnt++;
                    }
                    if (humidityLeft != -1000)
                    {
                        if (humidityLeft > 100)
                            humidityLeft = 100;
                        textBox2.Text = (Math.Round(humidityLeft, 2)).ToString();
                        textBox2.Refresh();
                        exHumidityLeft = humidityLeft;
                        if (firstHumLeftFlag)
                        {
                            firstMeasurement.humidityLeft = humidityLeft;
                            firstHumLeftFlag = false;
                        }
                        sumHumidityLeft += humidityLeft;
                        humLeftMeasureCnt++;
                    }
                    await Task.Delay(200);
                } else
                {
                    pictureBox1.Image = Image.FromFile(Directory.GetCurrentDirectory() + @"\resources\images\red_circle.png");
                    pictureBox1.Refresh();
                }

                ivitRight = await Modbus .Connect(comPort, 2, baudrate, Parity.None, StopBits.One, 100);

                if (ivitRight.Connected)
                {
                    try
                    {
                        temperatureRight = await Modbus.ReadRegisters(ivitRight, 0x0022, 2, inputReg);
                        humidityRight = await Modbus.ReadRegisters(ivitRight, 0x0016, 2, inputReg);
                        pictureBox2.Image = Image.FromFile(Directory.GetCurrentDirectory() + @"\resources\images\green_circle.png");
                        pictureBox2.Refresh();
                    } catch (TimeoutException)
                    {
                        pictureBox2.Image = Image.FromFile(Directory.GetCurrentDirectory() + @"\resources\images\red_circle.png");
                        pictureBox2.Refresh();
                    }

                    while (ivitRight.Connected)
                    {
                        ivitRight.Disconnect();
                    }

                    if (temperatureRight != -1000)
                    {
                        textBox3.Text = (Math.Round(temperatureRight, 2)).ToString();
                        textBox3.Refresh();
                        exTemperatureRight = temperatureRight;
                        if (firstTempRightFlag)
                        {
                            firstMeasurement.temperatureRight = temperatureRight;
                            firstTempRightFlag = false;
                        }
                        sumTemperatureRight += temperatureRight;
                        tempRightMeasureCnt++;
                    }
                    if (humidityRight != -1000)
                    {
                        if (humidityRight > 100)
                            humidityRight = 100;
                        textBox4.Text = (Math.Round(humidityRight, 2)).ToString();
                        textBox4.Refresh();
                        exHumidityRight = humidityRight;
                        if (firstHumRightFlag)
                        {
                            firstMeasurement.humidityRight = humidityRight;
                            firstHumRightFlag = false;
                        }
                        sumHumidityRight += humidityRight;
                        humRightMeasureCnt++;
                    }
                    await Task.Delay(200);
                } else
                {
                    pictureBox2.Image = Image.FromFile(Directory.GetCurrentDirectory() + @"\resources\images\red_circle.png");
                    pictureBox2.Refresh();
                }

                flowMeter = await Modbus.Connect(comPort, 7, baudrate, Parity.None, StopBits.Two, 100);

                if (flowMeter.Connected) {
                    
                    try
                    {
                        flowRate = FlowCalculation.CalculateFlowVelocity(await Modbus.ReadRegisters(flowMeter, 0x1009, 2, holdingReg));
                        pictureBox3.Image = Image.FromFile(Directory.GetCurrentDirectory() + @"\resources\images\green_circle.png");
                        pictureBox3.Refresh();
                    } catch (System.TimeoutException)
                    {
                        pictureBox3.Image = Image.FromFile(Directory.GetCurrentDirectory() + @"\resources\images\red_circle.png");
                        pictureBox3.Refresh();
                    }

                    while (flowMeter.Connected)
                    {
                        flowMeter.Disconnect();
                    }

                    if (flowRate > 0)
                    {
                        textBox9.Text = (Math.Round(flowRate, 2)).ToString();
                        textBox9.Refresh();
                        exFlowRate = flowRate;
                        if (firstFRFlag)
                        {
                            firstMeasurement.flowRate = flowRate;
                            firstFRFlag = false;
                        }
                    } 
                } else
                {
                    pictureBox3.Image = Image.FromFile(Directory.GetCurrentDirectory() + @"\resources\images\red_circle.png");
                    pictureBox3.Refresh();
                }

                if (!firstTempLeftFlag && !firstHumLeftFlag && !firstTempRightFlag && !firstHumRightFlag && !firstFRFlag && !firstMeasurmentDone)
                {
                    measurements.Enqueue(firstMeasurement);
                    firstMeasurmentDone = true;
                }

                await Task.Delay(1000);
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            timer1.Interval = 1000;
            timer2.Interval = 60000;
            timer3.Interval = 30000;
            timer1.Tick += new EventHandler(timer1_Tick);
            timer2.Tick += new EventHandler(timer2_Tick);
            timer3.Tick += new EventHandler(timer3_Tick);
            pictureBox1.Image = Image.FromFile(Directory.GetCurrentDirectory() + @"\resources\images\gray_circle.png");
            pictureBox2.Image = Image.FromFile(Directory.GetCurrentDirectory() + @"\resources\images\gray_circle.png");
            pictureBox3.Image = Image.FromFile(Directory.GetCurrentDirectory() + @"\resources\images\gray_circle.png");
            pictureBox4.Image = Image.FromFile(Directory.GetCurrentDirectory() + @"\resources\images\folder.png");
            pictureBox6.Image = Image.FromFile(Directory.GetCurrentDirectory() + @"\resources\images\pribor.png");
            pictureBox7.Image = Image.FromFile(Directory.GetCurrentDirectory() + @"\resources\images\line.png");
            comboBox1.Items.Add("сек.");
            comboBox1.Items.Add("мин.");
            comboBox1.SelectedIndex = 0;
            Update();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            stop = true;
            button1.Enabled = true;
            button2.Enabled = false;
            timer1.Stop();
            timer2.Stop();
            timer3.Stop();
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            excelWrite = true;
            button3.Enabled = false;
            button5.Enabled = true;
        }

        private void openFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_SerialPort");
            string uporName = @"Silicon Labs CP210x USB to UART Bridge *";
            // string bluetoothName = @"Стандартный последовательный порт по соединению Bluetooth *";
            foreach (ManagementObject service in searcher.Get())
            {
                if (Regex.IsMatch(service["Name"].ToString(), uporName))
                {
                    label1.Text = "Устройство обнаружено" + " (" + service["DeviceId"].ToString() + ")";
                    label1.ForeColor = Color.Green;
                    label1.Refresh();
                    comPort = service["DeviceId"].ToString();
                    button1.Enabled = true;

                }

                else
                {
                    label1.Text = "Устройство не обнаружено";
                    label1.ForeColor = Color.Red;
                    label1.Refresh();
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            excelWrite = false;
            button5.Enabled = false;
            button3.Enabled = true;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel Files(.xlsx) | *.xlsx";
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
            {
                return;
            } else
            {
                document = new FileInfo(openFileDialog1.FileName);
                excel = new ExcelPackage(document);
                textBox1.Text = openFileDialog1.FileName;

                if (!docExist)
                {
                    if (excel.Workbook.Worksheets.First().Dimension == null) {
                        excel.Workbook.Properties.Title = "Испытание";
                        sheet = excel.Workbook.Worksheets.First();
                        sheet.Name = "УПОР";
                        sheet.Cells[1, 1].Value = "Время";
                        sheet.Cells[1, 2].Value = "Температура на входе, °С";
                        sheet.Cells[1, 3].Value = "Относительная влажность на входе, %";
                        sheet.Cells[1, 4].Value = "Температура на выходе, °С";
                        sheet.Cells[1, 5].Value = "Относительная влажность на выходе, %";
                        sheet.Cells[1, 6].Value = "CO2 на входе, %";
                        sheet.Cells[1, 7].Value = "СO2 на выходе, &";
                        sheet.Cells[1, 8].Value = "Подано газа, л";
                        sheet.Cells[1, 9].Value = "Разность давлений, Па";
                        sheet.Cells[1, 10].Value = "Расход газа, м³/ч";

                        sheet.Cells[1, 1, 1, 10].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        sheet.Cells[1, 1, 1, 10].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        sheet.Cells[1, 1, 1, 10].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        sheet.Cells[1, 1, 1, 10].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                        excel.Save();
                        docExist = true;
                    } else
                    {
                        sheet = excel.Workbook.Worksheets.First();
                        autoRowCnt = excel.Workbook.Worksheets.First().Dimension.End.Row + 1;
                        manualRowCnt = autoRowCnt;
                        sheet.Cells[1, 1, autoRowCnt, 10].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        sheet.Cells[1, 1, autoRowCnt, 10].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        sheet.Cells[1, 1, autoRowCnt, 10].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        sheet.Cells[1, 1, autoRowCnt, 10].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        excel.Save();
                        docExist = true;
                    }
                }
                else
                {
                    sheet = excel.Workbook.Worksheets[0];
                }

                button3.Enabled = true;
                button8.Enabled = true;
            }
        }

        private void saveFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }

        private void openFileDialog1_FileOk_1(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            while (measurements.Count != 0)
                measurements.Dequeue();
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
        }

        private void button7_Click(object sender, EventArgs e)
        {
            bool timerWasEnabled = timer3.Enabled;
            timer3.Stop();
            if (comboBox1.SelectedIndex == 0) {
                timer3.Interval = 1000 * (Int32.Parse(textBox11.Text));
            } else
            {
                timer3.Interval = 60000 * (Int32.Parse(textBox11.Text));
            }
            if (timerWasEnabled)
                timer3.Start();
            button7.Enabled = false;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            button7.Enabled = true;
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label29_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }


        private void textBox13_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && e.KeyChar != 46 && e.KeyChar !=8)
                e.Handled = true;
        }

        private void textBox14_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && e.KeyChar != 46 && e.KeyChar !=8)
                e.Handled = true;
        }

        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && e.KeyChar != 46 && e.KeyChar != 8)
                e.Handled = true;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            pictureBox5.Image = Image.FromFile(Directory.GetCurrentDirectory() + @"\resources\images\green_circle.png");
            pictureBox5.Refresh();
            try
            {
                if(textBox12.Text != "")
                    sheet.Cells[manualRowCnt, 6].Value = Convert.ToDouble(textBox12.Text.Replace(".", ","));
                if (textBox13.Text != "")
                    sheet.Cells[manualRowCnt, 7].Value = Convert.ToDouble(textBox13.Text.Replace(".", ","));
                if (textBox14.Text != "")
                    sheet.Cells[manualRowCnt, 8].Value = Convert.ToDouble(textBox14.Text.Replace(".", ","));
                if (textBox15.Text != "")
                    sheet.Cells[manualRowCnt, 9].Value = Convert.ToDouble(textBox15.Text.Replace(".", ","));
                manualRowCnt++;
                excel.Save();
            }
            catch (System.NullReferenceException) { 
            }
            pictureBox5.Image = null;
            pictureBox5.Refresh();
        }

        private void label27_Click(object sender, EventArgs e)
        {

        }

        private void textBox15_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && e.KeyChar != 46 && e.KeyChar != 8)
                e.Handled = true;
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object Sender, EventArgs e)
        {
            experimentTime = experimentTime.Add(new TimeSpan(0, 0, 1));
            label13.Text = experimentTime.ToString(@"hh\:mm\:ss");
            label13.Refresh();
        }

        private void timer2_Tick(object Sender, EventArgs e)
        {
            avrgTemperatureLeft = sumTemperatureLeft / tempLeftMeasureCnt;
            sumTemperatureLeft = 0;
            tempLeftMeasureCnt = 0;
            textBox5.Text = (Math.Round(avrgTemperatureLeft, 2)).ToString();
            textBox5.Refresh();

            avrgHumidityLeft = sumHumidityLeft / humLeftMeasureCnt;
            sumHumidityLeft = 0;
            humLeftMeasureCnt = 0;
            textBox6.Text = (Math.Round(avrgHumidityLeft, 2)).ToString();
            textBox6.Refresh();

            avrgTemperatureRight = sumTemperatureRight / tempRightMeasureCnt;
            sumTemperatureRight = 0;
            tempRightMeasureCnt = 0;
            textBox7.Text = (Math.Round(avrgTemperatureRight, 2)).ToString();
            textBox7.Refresh();

            avrgHumidityRight = sumHumidityRight / humRightMeasureCnt;
            sumHumidityRight = 0;
            humRightMeasureCnt = 0;
            textBox8.Text = (Math.Round(avrgHumidityRight, 2)).ToString();
            textBox8.Refresh();
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            new SoundPlayer(Directory.GetCurrentDirectory() + @"\resources\sounds\beep.wav").Play();
            measurements.Enqueue(new Measurement(DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss"),
                   (float)Math.Round(exTemperatureLeft, 2),
                   (float)Math.Round(exTemperatureRight, 2),
                   (float)Math.Round(exHumidityLeft, 2),
                   (float)Math.Round(exHumidityRight, 2),
                   (float)Math.Round(exFlowRate, 2)));
            if (excelWrite)
            {
                pictureBox5.Image = Image.FromFile(Directory.GetCurrentDirectory() + @"\resources\images\green_circle.png");
                pictureBox5.Refresh();
                while (measurements.Count != 0)
                {
                    sheet.Cells[autoRowCnt, 1].Value = measurements.Peek().timestamp;
                    sheet.Cells[autoRowCnt, 2].Value = Math.Round(measurements.Peek().temperatureLeft, 2);
                    sheet.Cells[autoRowCnt, 3].Value = Math.Round(measurements.Peek().humidityLeft, 2);
                    sheet.Cells[autoRowCnt, 4].Value = Math.Round(measurements.Peek().temperatureRight, 2);
                    sheet.Cells[autoRowCnt, 5].Value = Math.Round(measurements.Peek().humidityRight, 2);
                    sheet.Cells[autoRowCnt, 10].Value = Math.Round(measurements.Peek().flowRate, 2);
                    measurements.Dequeue();

                    sheet.Cells[autoRowCnt, 1, autoRowCnt, 10].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                    sheet.Cells[autoRowCnt, 1, autoRowCnt, 10].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                    sheet.Cells[autoRowCnt, 1, autoRowCnt, 10].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                    sheet.Cells[autoRowCnt, 1, autoRowCnt, 10].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;

                    sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                    excel.Save();
                    autoRowCnt++;
                }
            }
            pictureBox5.Image = null;
            pictureBox5.Refresh();
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }
    }
}

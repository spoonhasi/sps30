using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using System.IO.Ports;
using System.Threading;
using System.Windows.Threading;

namespace SPS30_TEST
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        SerialPort selectedPort;
        Thread threadPortSignal;
        bool threadPortSignalStop = true;

        public MainWindow()
        {
            InitializeComponent();

            makePortComboBox();
        }

        private void makePortComboBox()
        {
            comboBoxPort.Items.Clear();

            string[] portlist = SerialPort.GetPortNames();
            if (portlist.Length > 0)
            {
                for (int i = 0; i < portlist.Length; i++)
                {
                    comboBoxPort.Items.Add(portlist[i]);
                }
                comboBoxPort.SelectedIndex = 0;
            }
        }

        public void ThreadPortSignal()
        {
            byte[] bufferTmp = new byte[256];
            byte[] buffer = new byte[256];
            int bufferOffset = 0;
            int readBufferOffset = 0;
            while (!threadPortSignalStop)
            {
                Thread.Sleep(20); //20밀리초마다 작동합니다.
                int bufferCnt = selectedPort.BytesToRead;
                if (bufferCnt > 0)
                {
                    selectedPort.Read(bufferTmp, 0, bufferCnt);
                    for (int i = 0; i < bufferCnt; i++)
                    {
                        buffer[bufferOffset] = bufferTmp[i];
                        bufferOffset++;
                        if (bufferOffset == 256)
                        {
                            bufferOffset = 0;
                        }
                    }
                }

                int cntCal = 0;
                if (bufferOffset >= readBufferOffset)
                {
                    cntCal = bufferOffset - readBufferOffset;
                }
                else
                {
                    cntCal = 256 + bufferOffset - readBufferOffset;
                }

                int startOffset = 0;
                bool startIs = false;
                int endOffset = 0;
                bool endIs = false;
                for (int i = 0; i < cntCal; i++)
                {
                    int thisOffset = readBufferOffset + i;
                    if (thisOffset > 255)
                    {
                        thisOffset = thisOffset - 256;
                    }
                    int nextOffset = readBufferOffset + i + 1;
                    if (nextOffset > 255)
                    {
                        nextOffset = nextOffset - 256;
                    }
                    if (startIs)
                    {
                        if (buffer[thisOffset] == 0x7E)
                        {
                            endOffset = thisOffset + 1;
                            if (endOffset > 255)
                            {
                                endOffset = endOffset - 256;
                            }
                            endIs = true;
                        }
                    }

                    if (!startIs)
                    {
                        if (buffer[thisOffset] == 0x7E && buffer[nextOffset] == 0x00)
                        {
                            startOffset = thisOffset;
                            startIs = true;
                        }
                    }
                }

                if (endIs)
                {
                    int cntMisoData = 0;
                    if (endOffset >= startOffset)
                    {
                        cntMisoData = endOffset - startOffset;
                    }
                    else
                    {
                        cntMisoData = 256 + endOffset - startOffset;
                    }

                    byte[] misoData = new byte[cntMisoData];
                    for (int i = 0; i < cntMisoData; i++)
                    {
                        int thisOffset = startOffset + i;
                        if (thisOffset > 255)
                        {
                            thisOffset = thisOffset - 256;
                        }
                        misoData[i] = buffer[thisOffset];
                    }
                    readBufferOffset = endOffset;
                    MISO(misoData); // 포트에서 받은 (완전한)Data를 GetMISO함수로 보냅니다. GetMISO함수에서 값을 해석하는 작업을 합니다.
                }
            }
        }

        public void MISO(byte[] buffer)
        {
            // 메뉴얼에 따른 값 변환입니다. 2바이트를 1바이트로 줄이는 작업이기 때문에 List를 이용합니다.
            List<byte> miso = new List<byte>();
            miso.Add(buffer[0]);
            for (int i = 1; i < buffer.Length - 1; i++)
            {
                if (buffer[i] == 0x7D)
                {
                    if (buffer[i + 1] == 0x5E)
                    {
                        miso.Add(0x7E);
                    }
                    else if (buffer[i + 1] == 0x5D)
                    {
                        miso.Add(0x7D);
                    }
                    else if (buffer[i + 1] == 0x31)
                    {
                        miso.Add(0x11);
                    }
                    else if (buffer[i + 1] == 0x33)
                    {
                        miso.Add(0x13);
                    }
                    i++;
                }
                else
                {
                    miso.Add(buffer[i]);
                }
            }
            miso.Add(buffer[buffer.Length - 1]);

            // 변환까지 마친 응답값을 16진수로 표현합니다. 값을 확인해서 메뉴얼과 비교하기 위한 용도입니다. 개발중에만 활용하면 됩니다.
            string resultText = "\r\n" + "응답메세지 시작";
            for (int i = 0; i < miso.Count; i++)
            {
                string hex = "0x" + String.Format("{0:X2}", miso[i]);
                resultText = resultText + "\r\n" + hex;
            }
            resultText = resultText + "\r\n" + "응답메세지 끝";

            // 변환까지 마친 응답값으로 해석합니다.
            if (miso[3] != 0x00) // 0x00만 NoError입니다. 그외에 값이 [3]자리에 있을 경우 에러입니다. 해당값의 에러메세지는 메뉴얼에 있습니다.
            {
                switch (miso[3]) // 각 에러값에 해당하는 메세지입니다.
                {
                    case 0x01:
                        resultText = resultText + "\r\n" + "Wrong data length for this command (too much or little data)";
                        break;
                    case 0x02:
                        resultText = resultText + "\r\n" + "Unknown command";
                        break;
                    case 0x03:
                        resultText = resultText + "\r\n" + "No access right for command";
                        break;
                    case 0x04:
                        resultText = resultText + "\r\n" + "Illegal command parameter or parameter out of allowed range";
                        break;
                    case 0x28:
                        resultText = resultText + "\r\n" + "Internal function argument out of range";
                        break;
                    case 0x43:
                        resultText = resultText + "\r\n" + "Command not allowed in current state";
                        break;
                    default:
                        resultText = resultText + "\r\n" + "Unknown error";
                        break;
                }
            }
            else // NoError인 경우입니다. 어떤 명령의 응답인지 [2]자리의 값으로 알수 있습니다. 이 부분에서 응답값을 텍스트나 그래프에 보내야 합니다.
            {
                if (miso[2] == 0x03) // Read Measured Values (0x03) : 오염수치를 가져옵니다. 그래프에 반영해야할 실제 값들입니다. 장비가 1초마다 새값을 만듭니다.
                {
                    if (miso.Count == 7) // 값을 가져오고 1초내에 값을 또 요청하면 새값이 없으므로 응답값에 Data가 없습니다. 응답값이 7바이트면 Data는 없습니다.
                    {
                        resultText = resultText + "\r\n" + "Empty response (too fast or sensor off)";
                    }
                    else if (miso.Count == 27) // 응답형식이 Big-endian unsigned 16-bit integer values 일 경우 사용합니다. 정수값만 있습니다.
                    {
                        List<float> listPM = new List<float>();
                        for (int i = 5; i < 24; i += 2)
                        {
                            string resultPM = String.Format("{0:X2}", miso[i]) + String.Format("{0:X2}", miso[i + 1]);
                            listPM.Add((UInt16.Parse(resultPM, System.Globalization.NumberStyles.HexNumber)));
                        }

                        resultText = resultText + "\r\n" + "Mass Concentration PM1.0 [µg/m³]: " + listPM[0];
                        resultText = resultText + "\r\n" + "Mass Concentration PM2.5 [µg/m³]: " + listPM[1];
                        resultText = resultText + "\r\n" + "Mass Concentration PM4.0 [µg/m³]: " + listPM[2];
                        resultText = resultText + "\r\n" + "Mass Concentration PM10 [µg/m³]: " + listPM[3];
                        resultText = resultText + "\r\n" + "Number Concentration PM0.5 [#/cm³]: " + listPM[4];
                        resultText = resultText + "\r\n" + "Number Concentration PM1.0 [#/cm³]: " + listPM[5];
                        resultText = resultText + "\r\n" + "Number Concentration PM2.5 [#/cm³]: " + listPM[6];
                        resultText = resultText + "\r\n" + "Number Concentration PM4.0 [#/cm³]: " + listPM[7];
                        resultText = resultText + "\r\n" + "Number Concentration PM10 [#/cm³]: " + listPM[8];
                        resultText = resultText + "\r\n" + "Typical Particle Size [µm]: " + listPM[9];
                    }
                    else // 응답형식이 Big-endian IEEE754 float values 일 경우 사용합니다. 소수점도 있습니다.
                    {
                        List<float> listPM = new List<float>();
                        for (int i = 5; i < 45; i += 4)
                        {
                            string resultPM = String.Format("{0:X2}", miso[i]) + String.Format("{0:X2}", miso[i + 1]) + String.Format("{0:X2}", miso[i + 2]) + String.Format("{0:X2}", miso[i + 3]);
                            listPM.Add(IEEEtoFloat(resultPM));
                        }

                        resultText = resultText + "\r\n" + "Mass Concentration PM1.0 [µg/m³]: " + listPM[0];
                        resultText = resultText + "\r\n" + "Mass Concentration PM2.5 [µg/m³]: " + listPM[1];
                        resultText = resultText + "\r\n" + "Mass Concentration PM4.0 [µg/m³]: " + listPM[2];
                        resultText = resultText + "\r\n" + "Mass Concentration PM10 [µg/m³]: " + listPM[3];
                        resultText = resultText + "\r\n" + "Number Concentration PM0.5 [#/cm³]: " + listPM[4];
                        resultText = resultText + "\r\n" + "Number Concentration PM1.0 [#/cm³]: " + listPM[5];
                        resultText = resultText + "\r\n" + "Number Concentration PM2.5 [#/cm³]: " + listPM[6];
                        resultText = resultText + "\r\n" + "Number Concentration PM4.0 [#/cm³]: " + listPM[7];
                        resultText = resultText + "\r\n" + "Number Concentration PM10 [#/cm³]: " + listPM[8];
                        resultText = resultText + "\r\n" + "Typical Particle Size [µm]: " + listPM[9];
                    }
                }
                else if (miso[2] == 0x00) // Start Measurement (0x00) : 작동명령의 응답입니다.
                {
                    resultText = resultText + "\r\n" + "Start Measurement";
                }
                else if (miso[2] == 0x01) // Stop Measurement (0x01) : 정지명령의 응답입니다.
                {
                    resultText = resultText + "\r\n" + "Stop Measurement";
                }
                else if (miso[2] == 0x56) // Start Fan Cleaning (0x56) : 팬을 잠시동안만 최대치로 돌려 청소합니다.
                {
                    resultText = resultText + "\r\n" + "Fan Cleaning";
                }
                else if (miso[2] == 0xD3) // Device Reset (0xD3) : 장비를 리셋합니다. 정지명령과 효과가 같은것 같습니다. (확실치 않음)
                {
                    resultText = resultText + "\r\n" + "Device Reset";
                }
                else if (miso[2] == 0xD1) // Read Version (0xD1): Firmware, Hardware, SHDLC의 버전을 읽어옵니다.
                {
                    string firmwareV = Convert.ToInt32(miso[5]).ToString() + "." + Convert.ToInt32(miso[6]).ToString();
                    string hardwareV = Convert.ToInt32(miso[8]).ToString();
                    string shdlcV = Convert.ToInt32(miso[10]).ToString() + "." + Convert.ToInt32(miso[11]).ToString();

                    // 여기에서 값들을 활용하면 됩니다.
                    resultText = resultText + "\r\n" + "Firmware V" + firmwareV + ", Hardware V" + hardwareV + ", SHDLC V" + shdlcV;
                }
                else if (miso[2] == 0xD0) // Device Information (0xD0) : Product Type과 Serial Number 둘중 하나의 응답값입니다. 이 두개는 명령값을 공유하므로 응답Data로 구분합니다.
                {
                    string deviceInformation = "";
                    for (int i = 5; i < miso.Count; i++)
                    {
                        if (miso[i] == 0x00)
                        {
                            break;
                        }
                        deviceInformation = deviceInformation + Convert.ToChar(miso[i]);
                    }
                    if (deviceInformation == "00080000") // Product Type의 응답값이 00080000이 나오면 SPS30이라는 의미입니다. SPS30은 항상 이 값이 나옵니다.
                    {
                        resultText = resultText + "\r\n" + "Product Type: " + deviceInformation;
                    }
                    else // 응답값이 00080000이 아니라면 Serial Number라는 의미입니다.
                    {
                        resultText = resultText + "\r\n" + "Serial Number: " + deviceInformation;
                    }
                }
                else if (miso[2] == 0x10) // Sleep (0x10) : 장비를 Sleep상태로 만듭니다. Wake-up 명령이 있기전까지 모든 명령이 안먹히게 됩니다.
                {
                    resultText = resultText + "\r\n" + "Device Sleep";
                }
                else if (miso[2] == 0x11) // Wake-up (0x11) : 장비를 Sleep상태에서 깨웁니다.
                {
                    resultText = resultText + "\r\n" + "Device Wake-up";
                }
                else if (miso[2] == 0x80) // Read/Write Auto Cleaning Interval (0x80) : 현재 설정되어 있는 자동 팬 청소 주기를 읽어오거나 새로 설정합니다. 초단위입니다.
                {
                    if (miso.Count == 7) // 응답값이 7byte면 Data는 없다는 뜻입니다. 팬 청소 주기를 설정하는 명령을 보낸 경우, 이렇게 응답이 옵니다.
                    {
                        resultText = resultText + "\r\n" + "Write Auto Cleaning Interval";
                        // 자동청소 주기 읽어오기
                        byte[] txData = { 0x00 };
                        MOSI(0x80, txData);
                    }
                    else // 현재 설정되어 있는 자동 팬 청소 주기를 읽어오라고 했을 경우의 응답값입니다.
                    {
                        string autoCleaningIntervalStr = "";
                        for (int i = 5; i < 9; i++)
                        {
                            autoCleaningIntervalStr = autoCleaningIntervalStr + String.Format("{0:X2}", miso[i]);
                        }
                        Int32 autoCleaningInterval = (Int32.Parse(autoCleaningIntervalStr, System.Globalization.NumberStyles.HexNumber));

                        // 여기에서 값들을 활용하면 됩니다.
                        resultText = resultText + "\r\n" + "Auto Cleaning Interval : " + autoCleaningInterval;
                    }
                }
            }
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate
            {
                textBoxResultMessage.Text = textBoxResultMessage.Text + resultText;
                textBoxResultMessage.ScrollToEnd();
            }));
        }

        public float IEEEtoFloat(string hex32Input)
        {
            int len = hex32Input.Length / 2;

            byte[] bytes = new byte[len];
            for (int i = 0; i < len; i++)
            {
                bytes[i] = Convert.ToByte(hex32Input.Substring(i * 2, 2), 16);
            }

            Array.Reverse(bytes);

            float data = BitConverter.ToSingle(bytes, 0);
            return data;
        }

        public int MOSI(byte cmd, byte[] txData)
        {
            List<byte> mosi = new List<byte>();

            mosi.Add(0x7E); // Start
            mosi.Add(0x00); // ADR
            mosi.Add(cmd);  // CMD

            byte dataLength = (byte)txData.Length;

            if (dataLength == 0x7E) // L
            {
                mosi.Add(0x7D);
                mosi.Add(0x5E);
            }
            else if (dataLength == 0x7D)
            {
                mosi.Add(0x7D);
                mosi.Add(0x5D);
            }
            else if (dataLength == 0x11)
            {
                mosi.Add(0x7D);
                mosi.Add(0x31);
            }
            else if (dataLength == 0x13)
            {
                mosi.Add(0x7D);
                mosi.Add(0x33);
            }
            else
            {
                mosi.Add(dataLength);
            }

            for (int i = 0; i < txData.Length; i++) // TX Data
            {
                if (txData[i] == 0x7E)
                {
                    mosi.Add(0x7D);
                    mosi.Add(0x5E);
                }
                else if (txData[i] == 0x7D)
                {
                    mosi.Add(0x7D);
                    mosi.Add(0x5D);
                }
                else if (txData[i] == 0x11)
                {
                    mosi.Add(0x7D);
                    mosi.Add(0x31);
                }
                else if (txData[i] == 0x13)
                {
                    mosi.Add(0x7D);
                    mosi.Add(0x33);
                }
                else
                {
                    mosi.Add(txData[i]);
                }
            }

            byte checkSum = 0;
            checkSum += 0x00;
            checkSum += cmd;
            checkSum += dataLength;
            for (int i = 0; i < txData.Length; i++)
            {
                checkSum += txData[i];
            }

            checkSum = (byte)~checkSum;

            if (checkSum == 0x7E) // CHK
            {
                mosi.Add(0x7D);
                mosi.Add(0x5E);
            }
            else if (checkSum == 0x7D)
            {
                mosi.Add(0x7D);
                mosi.Add(0x5D);
            }
            else if (checkSum == 0x11)
            {
                mosi.Add(0x7D);
                mosi.Add(0x31);
            }
            else if (checkSum == 0x13)
            {
                mosi.Add(0x7D);
                mosi.Add(0x33);
            }
            else
            {
                mosi.Add(checkSum);
            }

            mosi.Add(0x7E); // Stop

            // 16진수로 완성된 명령을 눈으로 보기 위함입니다. 개발중에만 사용합니다.
            string resultText = "\r\n" + "요청메세지 시작";
            for (int i = 0; i < mosi.Count; i++)
            {
                string hex = "0x" + String.Format("{0:X2}", mosi[i]);
                resultText = resultText + "\r\n" + hex;
            }
            resultText = resultText + "\r\n" + "요청메세지 종료";

            Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate
            {
                textBoxResultMessage.Text = textBoxResultMessage.Text + resultText;
                textBoxResultMessage.ScrollToEnd();
            }));

            byte[] buffer = mosi.ToArray();
            try
            {
                selectedPort.Write(buffer, 0, buffer.Length);
            }
            catch (InvalidOperationException)
            {

            }

            return 1;
        }

        private void ButtonConnect_Click(object sender, RoutedEventArgs e)
        {
            if (comboBoxPort.SelectedItem == null)
            {
                return;
            }

            SerialPort newSelectedPort = new SerialPort();
            newSelectedPort.PortName = comboBoxPort.SelectedItem.ToString();

            // 아래 설정값(BaudRate, DataBits, StopBits, Parity)은 메뉴얼에 나와있는 값입니다. 변경해서는 안됩니다.
            newSelectedPort.BaudRate = 115200;
            newSelectedPort.DataBits = 8;
            newSelectedPort.StopBits = StopBits.One;
            newSelectedPort.Parity = Parity.None;

            try
            {
                newSelectedPort.Open();
            }
            catch (UnauthorizedAccessException)
            {
                MessageBox.Show(selectedPort.PortName + " PORT는 이미 사용중입니다.");
                return;
            }
            selectedPort = newSelectedPort;

            threadPortSignalStop = false;
            threadPortSignal = new Thread(() => ThreadPortSignal());
            threadPortSignal.IsBackground = true;
            threadPortSignal.Start();

            textBoxResultMessage.Text = textBoxResultMessage.Text + "\r\n" + selectedPort.PortName + " PORT Connect";
            textBoxResultMessage.ScrollToEnd();
        }

        private void ButtonDisconnect_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                threadPortSignalStop = true;
                threadPortSignal.Join();
            }
            catch (NullReferenceException)
            {

            }

            if (selectedPort.IsOpen)
            {
                selectedPort.Close();
                textBoxResultMessage.Text = textBoxResultMessage.Text + "\r\n" + selectedPort.PortName + " PORT Disconnect";
                textBoxResultMessage.ScrollToEnd();
            }
        }

        private void ButtonStart_Click(object sender, RoutedEventArgs e)
        {
            byte[] txData = { 0x01, 0x03 };
            MOSI(0x00, txData);
        }

        private void ButtonStartUnsigned16_Click(object sender, RoutedEventArgs e)
        {
            byte[] txData = { 0x01, 0x05 };
            MOSI(0x00, txData);
        }

        private void ButtonStop_Click(object sender, RoutedEventArgs e)
        {
            byte[] txData = { };
            MOSI(0x01, txData);
        }

        private void ButtonSleep_Click(object sender, RoutedEventArgs e)
        {
            byte[] txData = { };
            MOSI(0x10, txData);
        }

        private void ButtonWake_up_Click(object sender, RoutedEventArgs e)
        {
            byte[] txData = { };
            MOSI(0x11, txData);
            MOSI(0x11, txData);
        }

        private void ButtonFanCleaning_Click(object sender, RoutedEventArgs e)
        {
            byte[] txData = { };
            MOSI(0x56, txData);
        }

        private void ButtonReadAutoCleaningInterval_Click(object sender, RoutedEventArgs e)
        {
            byte[] txData = { 0x00 };
            MOSI(0x80, txData);
        }

        private void ButtonWriteAutoCleaningInterval_Click(object sender, RoutedEventArgs e)
        {
            // Write Auto Cleaning Interval (0x80) : 자동 팬 청소 주기를 새로 설정하는 명령입니다. 초단위입니다. "0"으로 설정하면 자동 팬 청소를 하지 않습니다.
            UInt32 inputInt32 = Convert.ToUInt32(textBoxAutoCleaning.Text); // *** 텍스트박스에서 숫자를 읽어와서 빅인디안16진수 32바이트로 고쳐서 전송합니다. ***
            string inputStr = inputInt32.ToString("X8");
            string inputStr01 = inputStr.Substring(0, 2);
            string inputStr02 = inputStr.Substring(2, 2);
            string inputStr03 = inputStr.Substring(4, 2);
            string inputStr04 = inputStr.Substring(6, 2);
            byte[] txData = { 0x00, Convert.ToByte(inputStr01, 16), Convert.ToByte(inputStr02, 16), Convert.ToByte(inputStr03, 16), Convert.ToByte(inputStr04, 16) };
            MOSI(0x80, txData);
        }

        private void ButtonProductType_Click(object sender, RoutedEventArgs e)
        {
            byte[] txData = { 0x00 };
            MOSI(0xD0, txData);
        }

        private void ButtonSerialNumber_Click(object sender, RoutedEventArgs e)
        {
            byte[] txData = { 0x03 };
            MOSI(0xD0, txData);
        }

        private void ButtonReadVersion_Click(object sender, RoutedEventArgs e)
        {
            byte[] txData = { };
            MOSI(0xD1, txData);
        }

        private void ButtonReadDeviceStatusRegister_Click(object sender, RoutedEventArgs e)
        {
            byte[] txData = { 0x00 };
            MOSI(0xD2, txData);
        }

        private void ButtonDeviceReset_Click(object sender, RoutedEventArgs e)
        {
            byte[] txData = { };
            MOSI(0xD3, txData);
        }

        private void ButtonReadMeasuredValues_Click(object sender, RoutedEventArgs e)
        {
            byte[] txData = { };
            MOSI(0x03, txData);
        }
    }
}

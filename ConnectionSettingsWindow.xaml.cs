using System.Windows;
using System.Windows.Controls;

namespace MitsubishiPlcToPC
{
    public partial class ConnectionSettingsWindow : Window
    {
        public ConnectionSettings Settings { get; private set; }

        public ConnectionSettingsWindow(ConnectionSettings currentSettings)
        {
            InitializeComponent();

            Settings = currentSettings;
            LoadSettings();
        }

        private void LoadSettings()
        {
            // Load connection type
            if (Settings.ConnectionType == "Ethernet")
            {
                EthernetRadio.IsChecked = true;
            }
            else
            {
                SerialRadio.IsChecked = true;
            }

            // Load Ethernet settings
            IpAddressTextBox.Text = Settings.IpAddress;
            PortTextBox.Text = Settings.Port;
            ProtocolComboBox.SelectedIndex = Settings.Protocol == "TCP" ? 0 : 1;

            // Load Serial settings
            SerialPortComboBox.SelectedIndex = int.Parse(Settings.SerialPort.Replace("COM", "")) - 1;

            switch (Settings.BaudRate)
            {
                case "9600": BaudRateComboBox.SelectedIndex = 0; break;
                case "19200": BaudRateComboBox.SelectedIndex = 1; break;
                case "38400": BaudRateComboBox.SelectedIndex = 2; break;
                case "57600": BaudRateComboBox.SelectedIndex = 3; break;
                case "115200": BaudRateComboBox.SelectedIndex = 4; break;
            }

            DataBitsComboBox.SelectedIndex = Settings.DataBits == "7" ? 0 : 1;
            StopBitsComboBox.SelectedIndex = Settings.StopBits == "1" ? 0 : 1;

            switch (Settings.Parity)
            {
                case "None": ParityComboBox.SelectedIndex = 0; break;
                case "Odd": ParityComboBox.SelectedIndex = 1; break;
                case "Even": ParityComboBox.SelectedIndex = 2; break;
            }

            // Load PLC settings
            switch (Settings.PlcType)
            {
                case "Q Series": PlcTypeComboBox.SelectedIndex = 0; break;
                case "iQ-R Series": PlcTypeComboBox.SelectedIndex = 1; break;
                case "iQ-F Series": PlcTypeComboBox.SelectedIndex = 2; break;
                case "L Series": PlcTypeComboBox.SelectedIndex = 3; break;
                case "FX Series": PlcTypeComboBox.SelectedIndex = 4; break;
            }

            StationNumberTextBox.Text = Settings.StationNumber;
        }

        private void ConnectionType_Changed(object sender, RoutedEventArgs e)
        {
            if (EthernetRadio.IsChecked == true)
            {
                EthernetPanel.Visibility = Visibility.Visible;
                SerialPanel.Visibility = Visibility.Collapsed;
            }
            else
            {
                EthernetPanel.Visibility = Visibility.Collapsed;
                SerialPanel.Visibility = Visibility.Visible;
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            Settings = new ConnectionSettings
            {
                ConnectionType = EthernetRadio.IsChecked == true ? "Ethernet" : "Serial",
                IpAddress = IpAddressTextBox.Text,
                Port = PortTextBox.Text,
                Protocol = ((ComboBoxItem)ProtocolComboBox.SelectedItem).Content.ToString(),
                SerialPort = ((ComboBoxItem)SerialPortComboBox.SelectedItem).Content.ToString(),
                BaudRate = ((ComboBoxItem)BaudRateComboBox.SelectedItem).Content.ToString(),
                DataBits = ((ComboBoxItem)DataBitsComboBox.SelectedItem).Content.ToString(),
                StopBits = ((ComboBoxItem)StopBitsComboBox.SelectedItem).Content.ToString(),
                Parity = ((ComboBoxItem)ParityComboBox.SelectedItem).Content.ToString(),
                PlcType = ((ComboBoxItem)PlcTypeComboBox.SelectedItem).Content.ToString(),
                StationNumber = StationNumberTextBox.Text
            };

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }

    public class ConnectionSettings
    {
        public string ConnectionType { get; set; } = "Ethernet";
        public string IpAddress { get; set; } = "192.168.1.100";
        public string Port { get; set; } = "5000";
        public string Protocol { get; set; } = "TCP";
        public string SerialPort { get; set; } = "COM1";
        public string BaudRate { get; set; } = "115200";
        public string DataBits { get; set; } = "8";
        public string StopBits { get; set; } = "1";
        public string Parity { get; set; } = "None";
        public string PlcType { get; set; } = "Q Series";
        public string StationNumber { get; set; } = "0";
    }
}
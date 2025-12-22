using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;

using ActUtlTypeLib;

namespace MitsubishiPlcToPC
{
    public partial class MainWindow : Window
    {
        // MX Component ActUtlType
        private ActUtlType plc;

        // Timer for auto refresh
        private DispatcherTimer autoRefreshTimer;

        // Connection settings
        private ConnectionSettings settings;

        // Data collections
        private ObservableCollection<PlcAddress> w100Data;
        private ObservableCollection<PlcAddress> y1000Data;
        private ObservableCollection<PlcAddress> w0Data;
        private ObservableCollection<PlcAddress> x1000Data;
        private ObservableCollection<PlcAddress> d1000Data;
        private ObservableCollection<PlcAddress> d1011Data;

        public MainWindow()
        {
            InitializeComponent();

            // Initialize PLC object
            plc = new ActUtlType();

            // Initialize timer
            autoRefreshTimer = new DispatcherTimer();
            autoRefreshTimer.Interval = TimeSpan.FromSeconds(1);
            autoRefreshTimer.Tick += AutoRefreshTimer_Tick;

            // Initialize default settings
            settings = new ConnectionSettings
            {
                ConnectionType = "Ethernet",
                IpAddress = "192.168.1.100",
                Port = "5000",
                PlcType = "Q Series"
            };

            // Initialize data
            InitializeDataGrids();
        }

        private void InitializeDataGrids()
        {
            // W100-W110 (PLC → 로봇)
            w100Data = new ObservableCollection<PlcAddress>
            {
                new PlcAddress { Address = "W100", Description = "인덱스 홀더 체결(1~8)", Value = 0 },
                new PlcAddress { Address = "W101", Description = "인덱스 홀더 체결(9~16)", Value = 0 },
                new PlcAddress { Address = "W102", Description = "인덱스 현재 번호", Value = 0 },
                new PlcAddress { Address = "W103", Description = "예비", Value = 0 },
                new PlcAddress { Address = "W104", Description = "예비", Value = 0 },
                new PlcAddress { Address = "W105", Description = "예비", Value = 0 },
                new PlcAddress { Address = "W106", Description = "예비", Value = 0 },
                new PlcAddress { Address = "W107", Description = "예비", Value = 0 },
                new PlcAddress { Address = "W108", Description = "예비", Value = 0 },
                new PlcAddress { Address = "W109", Description = "예비", Value = 0 },
                new PlcAddress { Address = "W110", Description = "예비", Value = 0 }
            };
            W100Grid.ItemsSource = w100Data;

            // Y1000-Y1010 (PLC → 로봇 ATC)
            y1000Data = new ObservableCollection<PlcAddress>
            {
                new PlcAddress { Address = "Y1000", Description = "ATC IN 위치확인", Value = 0 },
                new PlcAddress { Address = "Y1001", Description = "ATC OUT 위치확인", Value = 0 },
                new PlcAddress { Address = "Y1002", Description = "ATC 준비", Value = 0 },
                new PlcAddress { Address = "Y1003", Description = "ATC 에러", Value = 0 },
                new PlcAddress { Address = "Y1004", Description = "ATC 회전중", Value = 0 },
                new PlcAddress { Address = "Y1005", Description = "ATC 회전완료", Value = 0 },
                new PlcAddress { Address = "Y1006", Description = "예비", Value = 0 },
                new PlcAddress { Address = "Y1007", Description = "예비", Value = 0 },
                new PlcAddress { Address = "Y1008", Description = "예비", Value = 0 },
                new PlcAddress { Address = "Y1009", Description = "예비", Value = 0 },
                new PlcAddress { Address = "Y1010", Description = "예비", Value = 0 }
            };
            Y1000Grid.ItemsSource = y1000Data;

            // W0-W10 (로봇 → PLC)
            w0Data = new ObservableCollection<PlcAddress>
            {
                new PlcAddress { Address = "W0", Description = "인덱스 회전 번호", Value = 0 },
                new PlcAddress { Address = "W1", Description = "예비", Value = 0 },
                new PlcAddress { Address = "W2", Description = "예비", Value = 0 },
                new PlcAddress { Address = "W3", Description = "예비", Value = 0 },
                new PlcAddress { Address = "W4", Description = "예비", Value = 0 },
                new PlcAddress { Address = "W5", Description = "예비", Value = 0 },
                new PlcAddress { Address = "W6", Description = "예비", Value = 0 },
                new PlcAddress { Address = "W7", Description = "예비", Value = 0 },
                new PlcAddress { Address = "W8", Description = "예비", Value = 0 },
                new PlcAddress { Address = "W9", Description = "예비", Value = 0 },
                new PlcAddress { Address = "W10", Description = "예비", Value = 0 }
            };
            W0Grid.ItemsSource = w0Data;

            // X1000-X1010 (로봇 → PLC ATC)
            x1000Data = new ObservableCollection<PlcAddress>
            {
                new PlcAddress { Address = "X1000", Description = "ATC IN 이송요구", Value = 0 },
                new PlcAddress { Address = "X1001", Description = "ATC OUT 이송요구", Value = 0 },
                new PlcAddress { Address = "X1002", Description = "ATC 회전 요구", Value = 0 },
                new PlcAddress { Address = "X1003", Description = "ATC 에러 리셋 요구", Value = 0 },
                new PlcAddress { Address = "X1004", Description = "예비", Value = 0 },
                new PlcAddress { Address = "X1005", Description = "예비", Value = 0 },
                new PlcAddress { Address = "X1006", Description = "예비", Value = 0 },
                new PlcAddress { Address = "X1007", Description = "예비", Value = 0 },
                new PlcAddress { Address = "X1008", Description = "예비", Value = 0 },
                new PlcAddress { Address = "X1009", Description = "예비", Value = 0 },
                new PlcAddress { Address = "X1010", Description = "예비", Value = 0 }
            };
            X1000Grid.ItemsSource = x1000Data;

            // D1000-D1010 (PLC ↔ PC)
            d1000Data = new ObservableCollection<PlcAddress>
            {
                new PlcAddress { Address = "D1000", Description = "버전번호", Value = 0 },
                new PlcAddress { Address = "D1001", Description = "인덱스 홀더 체결확인", Value = 0 },
                new PlcAddress { Address = "D1002", Description = "인덱스 현재 번호", Value = 0 },
                new PlcAddress { Address = "D1003", Description = "ATC IN, OUT 위치확인(0:IN, 1:OUT)", Value = 0 },
                new PlcAddress { Address = "D1004", Description = "ATC 준비완료(0:준비,1:회전가능)", Value = 0 },
                new PlcAddress { Address = "D1005", Description = "ATC 에러코드(0:정상, 1:에러)", Value = 0 },
                new PlcAddress { Address = "D1006", Description = "ATC 상태코드(0:회전중, 1:회전완료)", Value = 0 },
                new PlcAddress { Address = "D1007", Description = "예비", Value = 0 },
                new PlcAddress { Address = "D1008", Description = "예비", Value = 0 },
                new PlcAddress { Address = "D1009", Description = "예비", Value = 0 },
                new PlcAddress { Address = "D1010", Description = "예비", Value = 0 }
            };
            D1000Grid.ItemsSource = d1000Data;

            // D1011-D1021 (PC → PLC)
            d1011Data = new ObservableCollection<PlcAddress>
            {
                new PlcAddress { Address = "D1011", Description = "인덱스 회전 번호", Value = 0 },
                new PlcAddress { Address = "D1012", Description = "ATC IN 명령", Value = 0 },
                new PlcAddress { Address = "D1013", Description = "ATC OUT 명령", Value = 0 },
                new PlcAddress { Address = "D1014", Description = "ATC 회전요구", Value = 0 },
                new PlcAddress { Address = "D1015", Description = "ATC 에러리셋요구", Value = 0 },
                new PlcAddress { Address = "D1016", Description = "예비", Value = 0 },
                new PlcAddress { Address = "D1017", Description = "예비", Value = 0 },
                new PlcAddress { Address = "D1018", Description = "예비", Value = 0 },
                new PlcAddress { Address = "D1019", Description = "예비", Value = 0 },
                new PlcAddress { Address = "D1020", Description = "예비", Value = 0 },
                new PlcAddress { Address = "D1021", Description = "예비", Value = 0 }
            };
            D1011Grid.ItemsSource = d1011Data;
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            ConnectionSettingsWindow settingsWindow = new ConnectionSettingsWindow(settings);
            if (settingsWindow.ShowDialog() == true)
            {
                settings = settingsWindow.Settings;
            }
        }

        private void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Set ActLogicalStationNumber
                plc.ActLogicalStationNumber = int.Parse(settings.StationNumber);

                // Open connection
                int result = plc.Open();

                if (result == 0)
                {
                    // Success
                    StatusIndicator.Fill = new SolidColorBrush(Colors.Green);
                    StatusText.Text = "PLC 연결됨";
                    StatusText.Foreground = new SolidColorBrush(Colors.Green);

                    ConnectButton.IsEnabled = false;
                    DisconnectButton.IsEnabled = true;
                    AutoRefreshCheckBox.IsEnabled = true;

                    WelcomeScreen.Visibility = Visibility.Collapsed;
                    MonitoringScreen.Visibility = Visibility.Visible;

                    ConnectionInfo.Visibility = Visibility.Visible;
                    PlcTypeText.Text = $"PLC Type: {settings.PlcType}";

                    if (settings.ConnectionType == "Ethernet")
                    {
                        ConnectionDetailsText.Text = $"{settings.IpAddress}:{settings.Port}";
                    }
                    else
                    {
                        ConnectionDetailsText.Text = $"{settings.SerialPort} - {settings.BaudRate}bps";
                    }

                    // Initial read
                    ReadAllData();

                    MessageBox.Show("PLC 연결 성공!", "연결 성공", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show($"PLC 연결 실패!\n에러 코드: {result}", "연결 실패", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"연결 중 오류 발생:\n{ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DisconnectButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                autoRefreshTimer.Stop();
                plc.Close();

                StatusIndicator.Fill = new SolidColorBrush(Colors.Gray);
                StatusText.Text = "연결 안됨";
                StatusText.Foreground = new SolidColorBrush(Colors.Black);

                ConnectButton.IsEnabled = true;
                DisconnectButton.IsEnabled = false;
                AutoRefreshCheckBox.IsEnabled = false;
                AutoRefreshCheckBox.IsChecked = false;

                ConnectionInfo.Visibility = Visibility.Collapsed;
                WelcomeScreen.Visibility = Visibility.Visible;
                MonitoringScreen.Visibility = Visibility.Collapsed;

                MessageBox.Show("PLC 연결 해제됨", "연결 해제", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"연결 해제 중 오류:\n{ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AutoRefreshCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            autoRefreshTimer.Start();
        }

        private void AutoRefreshCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            autoRefreshTimer.Stop();
        }

        private void AutoRefreshTimer_Tick(object sender, EventArgs e)
        {
            ReadAllData();
        }

        private void ReadAllData()
        {
            try
            {
                // Read W addresses
                ReadWordDevices(w100Data, "W", 100);
                ReadWordDevices(w0Data, "W", 0);

                // Read D addresses
                ReadWordDevices(d1000Data, "D", 1000);
                ReadWordDevices(d1011Data, "D", 1011);

                // Read X/Y bit devices
                ReadBitDevices(x1000Data, "X", 1000);
                ReadBitDevices(y1000Data, "Y", 1000);

                // Update status displays
                UpdateStatusDisplays();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"데이터 읽기 오류:\n{ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ReadWordDevices(ObservableCollection<PlcAddress> collection, string deviceType, int startAddress)
        {
            int size = collection.Count;
            short[] values = new short[size];

            int result = plc.ReadDeviceBlock($"{deviceType}{startAddress}", size, out values[0]);

            if (result == 0)
            {
                for (int i = 0; i < size; i++)
                {
                    collection[i].Value = values[i];
                }
            }
        }

        private void ReadBitDevices(ObservableCollection<PlcAddress> collection, string deviceType, int startAddress)
        {
            for (int i = 0; i < collection.Count; i++)
            {
                short value;
                int result = plc.GetDevice($"{deviceType}{startAddress + i}", out value);

                if (result == 0)
                {
                    collection[i].Value = value;
                }
            }
        }

        private void UpdateStatusDisplays()
        {
            // Update ATC status
            var d1004 = d1000Data[4].Value; // ATC 준비완료
            var d1006 = d1000Data[6].Value; // ATC 상태코드
            var d1003 = d1000Data[3].Value; // ATC 위치
            var d1005 = d1000Data[5].Value; // ATC 에러

            AtcReadyStatus.Text = d1004 == 1 ? "준비: 완료" : "준비: 대기";
            AtcReadyStatus.Background = d1004 == 1 ? new SolidColorBrush(Color.FromRgb(200, 230, 201)) : new SolidColorBrush(Color.FromRgb(245, 245, 245));

            AtcRotationStatus.Text = d1006 == 1 ? "회전완료" : "회전중";
            AtcRotationStatus.Background = d1006 == 1 ? new SolidColorBrush(Color.FromRgb(187, 222, 251)) : new SolidColorBrush(Color.FromRgb(255, 249, 196));

            AtcPositionStatus.Text = d1003 == 0 ? "위치: IN" : "위치: OUT";
            AtcPositionStatus.Background = new SolidColorBrush(Color.FromRgb(227, 242, 253));

            AtcErrorStatus.Text = d1005 == 0 ? "에러: 정상" : "에러: 발생";
            AtcErrorStatus.Background = d1005 == 0 ? new SolidColorBrush(Color.FromRgb(232, 245, 233)) : new SolidColorBrush(Color.FromRgb(255, 205, 210));

            // Update index status
            RequestedIndexText.Text = d1011Data[0].Value.ToString();
            CurrentIndexText.Text = d1000Data[2].Value.ToString();
            OutputStatusText.Text = Convert.ToString(d1000Data[1].Value, 2).PadLeft(16, '0');
        }

        private void WriteDevice(string address, short value)
        {
            try
            {
                int result = plc.SetDevice(address, value);

                if (result == 0)
                {
                    // Success - refresh data
                    ReadAllData();
                }
                else
                {
                    MessageBox.Show($"쓰기 실패!\n주소: {address}\n에러 코드: {result}", "쓰기 오류", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"쓰기 중 오류:\n{ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AtcInButton_Click(object sender, RoutedEventArgs e)
        {
            WriteDevice("D1012", 1);
            System.Threading.Tasks.Task.Delay(500).ContinueWith(_ =>
            {
                Dispatcher.Invoke(() => WriteDevice("D1012", 0));
            });
        }

        private void AtcOutButton_Click(object sender, RoutedEventArgs e)
        {
            WriteDevice("D1013", 1);
            System.Threading.Tasks.Task.Delay(500).ContinueWith(_ =>
            {
                Dispatcher.Invoke(() => WriteDevice("D1013", 0));
            });
        }

        private void AtcRotateButton_Click(object sender, RoutedEventArgs e)
        {
            WriteDevice("D1014", 1);
            System.Threading.Tasks.Task.Delay(500).ContinueWith(_ =>
            {
                Dispatcher.Invoke(() => WriteDevice("D1014", 0));
            });
        }

        private void AtcResetButton_Click(object sender, RoutedEventArgs e)
        {
            WriteDevice("D1015", 1);
            System.Threading.Tasks.Task.Delay(500).ContinueWith(_ =>
            {
                Dispatcher.Invoke(() => WriteDevice("D1015", 0));
            });
        }

        private void IndexRotateButton_Click(object sender, RoutedEventArgs e)
        {
            if (int.TryParse(IndexNumberTextBox.Text, out int value))
            {
                if (value >= 1 && value <= 16)
                {
                    WriteDevice("D1011", (short)value);
                }
                else
                {
                    MessageBox.Show("인덱스 번호는 1~16 사이여야 합니다.", "입력 오류", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            else
            {
                MessageBox.Show("올바른 숫자를 입력하세요.", "입력 오류", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}
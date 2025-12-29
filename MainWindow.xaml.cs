using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;

using ActUtlTypeLib;

namespace MitsubishiPlcToPC
{
    public class ConnectionSettings
    {

        public int LogicalStationNumber { get; set; } = 1;

        public string PlcType { get; set; } = "Q Series";
        public string StationNumber { get; set; } = "0";
    }
    public partial class MainWindow : Window
    {
        private ActUtlType plc;

        private DispatcherTimer autoRefreshTimer;

        private ConnectionSettings settings;

        private ObservableCollection<PlcAddress> w100Data;
        private ObservableCollection<PlcAddress> y1000Data;
        private ObservableCollection<PlcAddress> w0Data;
        private ObservableCollection<PlcAddress> x1000Data;
        private ObservableCollection<PlcAddress> d1000Data;
        private ObservableCollection<PlcAddress> d1011Data;

        public MainWindow()
        {
            InitializeComponent();

            plc = new ActUtlType();

            autoRefreshTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(2) // 실시간 모니터링에 적합한 간격
            };
            autoRefreshTimer.Tick += AutoRefreshTimer_Tick;

            settings = new ConnectionSettings
            {
                PlcType = "Q Series",
                StationNumber = "0"
            };

            InitializeDataGrids();
        }


        private void InitializeDataGrids()
        {
            w100Data = CreatePlcAddressCollection(
                new[] { "W100", "W101", "W102", "W103", "W104", "W105", "W106", "W107", "W108", "W109", "W110" },
                new[] { "인덱스 홀더 체결(1~8)", "인덱스 홀더 체결(9~16)", "인덱스 현재 번호", "예비", "예비", "예비", "예비", "예비", "예비", "예비", "예비" }
            );
            W100Grid.ItemsSource = w100Data;

            y1000Data = CreatePlcAddressCollection(
                new[] { "Y1000", "Y1001", "Y1002", "Y1003", "Y1004", "Y1005", "Y1006", "Y1007", "Y1008", "Y1009", "Y1010" },
                new[] { "ATC IN 위치확인", "ATC OUT 위치확인", "ATC 준비", "ATC 에러", "ATC 회전중", "ATC 회전완료", "예비", "예비", "예비", "예비", "예비" }
            );
            Y1000Grid.ItemsSource = y1000Data;

            w0Data = CreatePlcAddressCollection(
                new[] { "W0", "W1", "W2", "W3", "W4", "W5", "W6", "W7", "W8", "W9", "W10" },
                new[] { "인덱스 회전 번호", "예비", "예비", "예비", "예비", "예비", "예비", "예비", "예비", "예비", "예비" }
            );
            W0Grid.ItemsSource = w0Data;

            x1000Data = CreatePlcAddressCollection(
                new[] { "X1000", "X1001", "X1002", "X1003", "X1004", "X1005", "X1006", "X1007", "X1008", "X1009", "X1010" },
                new[] { "ATC IN 이송요구", "ATC OUT 이송요구", "ATC 회전 요구", "ATC 에러 리셋 요구", "예비", "예비", "예비", "예비", "예비", "예비", "예비" }
            );
            X1000Grid.ItemsSource = x1000Data;

            d1000Data = CreatePlcAddressCollection(
                new[] { "D1000", "D1001", "D1002", "D1003", "D1004", "D1005", "D1006", "D1007", "D1008", "D1009", "D1010" },
                new[] { "버전번호", "인덱스 홀더 체결확인", "인덱스 현재 번호", "ATC IN, OUT 위치확인(0:IN, 1:OUT)", "ATC 준비완료(0:준비,1:회전가능)", "ATC 에러코드(0:정상, 1:에러)", "ATC 상태코드(0:회전중, 1:회전완료)", "예비", "예비", "예비", "예비" }
            );
            D1000Grid.ItemsSource = d1000Data;

            d1011Data = CreatePlcAddressCollection(
                new[] { "D1011", "D1012", "D1013", "D1014", "D1015", "D1016", "D1017", "D1018", "D1019", "D1020", "D1021" },
                new[] { "인덱스 회전 번호", "ATC IN 명령", "ATC OUT 명령", "ATC 회전요구", "ATC 에러리셋요구", "예비", "예비", "예비", "예비", "예비", "예비" }
            );
            D1011Grid.ItemsSource = d1011Data;
        }

        // 공통 생성 함수 추가
        private ObservableCollection<PlcAddress> CreatePlcAddressCollection(string[] addresses, string[] descriptions)
        {
            var collection = new ObservableCollection<PlcAddress>();
            for (int i = 0; i < addresses.Length; i++)
            {
                collection.Add(new PlcAddress
                {
                    Address = addresses[i],
                    Description = descriptions[i],
                    Value = 0
                });
            }
            return collection;
        }

        private void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                plc.ActLogicalStationNumber = int.Parse(settings.StationNumber);
                int result = plc.Open();
                MessageBox.Show(result.ToString());
                if (result == 0)
                {
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
                ReadWordDevices(w100Data, "W", 100);
                ReadWordDevices(w0Data, "W", 0);

                ReadWordDevices(d1000Data, "D", 1000);
                ReadWordDevices(d1011Data, "D", 1011);

                ReadBitDevices(x1000Data, "X", 1000);
                ReadBitDevices(y1000Data, "Y", 1000);

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
            int[] values = new int[size];
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
                int value;
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
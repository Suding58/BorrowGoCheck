using NAudio.Wave;
using System.Configuration;
using System.Management;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace BorrowGoCheck
{
    class Program
    {
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();

        const int SW_HIDE = 0;

        static string ip = ConfigurationManager.AppSettings["IpAddress"];
        static string port = ConfigurationManager.AppSettings["Port"];

        static string apiUrl = $"http://{ip}:{port}";
        static string apiCheckUrl = $"{apiUrl}/api/check-stat";
        const int checkIntervalSeconds = 40;
        static string? hwid = null;
        static int? itemId = null;

        static string status = "UNKNOWN";
        static bool isOnline = false;

        static string messageCommand = "";
        static bool showingBlockMessage = false;

        [DllImport("Kernel32")]
        private static extern bool SetConsoleCtrlHandler(SetConsoleCtrlEventHandler handler, bool add);
        private delegate bool SetConsoleCtrlEventHandler(CtrlType sig);

        // Win32 API สำหรับจัดตำแหน่ง MessageBox
        [DllImport("user32.dll")]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        const uint SWP_NOSIZE = 0x0001;
        const uint SWP_NOZORDER = 0x0004;

        private enum CtrlType
        {
            CTRL_C_EVENT = 0,
            CTRL_BREAK_EVENT = 1,
            CTRL_CLOSE_EVENT = 2,
            CTRL_LOGOFF_EVENT = 5,
            CTRL_SHUTDOWN_EVENT = 6
        }


        [STAThread]
        static void Main()
        {
            SetConsoleCtrlHandler(Handler, true);
            HideConsoleWindow();

            hwid = GetHardwareId();

            CheckAvailable();
            CheckAdminCommand();

            Console.ReadKey();
        }

        private static async void CheckAvailable()
        {
            while (true)
            {
                try
                {
                    if (await IsApiAvailable())
                    {
                        var itemData = await CheckStatus();

                        if (!isOnline)
                        {
                            UpdateOnlineStatus(true);
                            isOnline = true;
                        }

                        if (itemData == null || itemData?.hwid == null || itemData?.hwid == "")
                        {
                             RegisterHardwareId();
                        }
                        else if (itemData.status != "BORROWED")
                        {
                            status = itemData?.status ?? "UNKNOWN";
                            itemId = itemData?.id ?? 0;

                            string warningMessage = status switch
                            {
                                "AVAILABLE" => "เครื่องนี้ยังไม่ถูกยืมผ่านระบบ\nกรุณาดำเนินการยืนคืนก่อนใช้งาน",
                                "WAITAPPROVAL" => "เครื่องนี้อยู่ระหว่างรอการอนุมัติ\nกรุณารอการอนุมัติก่อนใช้งาน",
                                "MAINTENANCE" => "เครื่องนี้อยู่ระหว่างซ่อมบำรุง\nไม่สามารถใช้งานได้ในขณะนี้",
                                _ => "เครื่องไม่ถูกลงทะเบียนยืมคืน\nกรุณาแจ้งผู้ดูแลระบบ"
                            };
                            PlaySoundWithNAudio($"{status}_SOUND.mp3");
                            ShowBlockingMessageBox(warningMessage, 30);
                        }
                    }
                    else
                    {
                        Console.WriteLine("🔌 ไม่สามารถเชื่อมต่อ API ได้ (อาจล่มอยู่) รอตรวจสอบใหม่ในอีก 60 วินาที...");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ พบข้อผิดพลาดระหว่างตรวจสอบ: {ex.Message}");
                }

                await Task.Delay(TimeSpan.FromSeconds(checkIntervalSeconds));
            }
        }

        private static async void CheckAdminCommand()
        {
            while (true)
            {
                try
                {
                    if (await IsApiAvailable())
                    {
                        var itemData = await CheckStatus();
                        if (itemData != null)
                            CheckAdminCommands();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ พบข้อผิดพลาดระหว่างตรวจสอบ: {ex.Message}");
                }

                await Task.Delay(TimeSpan.FromSeconds(10));
            }
        }


        private static bool Handler(CtrlType signal)
        {
            switch (signal)
            {
                case CtrlType.CTRL_BREAK_EVENT:
                case CtrlType.CTRL_C_EVENT:
                case CtrlType.CTRL_LOGOFF_EVENT:
                case CtrlType.CTRL_SHUTDOWN_EVENT:
                case CtrlType.CTRL_CLOSE_EVENT:
                    UpdateOnlineStatus(false);
                    Environment.Exit(0);
                    return false;

                default:
                    return false;
            }
        }

        static async void UpdateOnlineStatus(bool isOnline)
        {

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    await client.PutAsync($"{apiUrl}/api/check-stat/{hwid}/{isOnline}", null);
                }
            }
            catch
            {

            }
        }

        static async void PlaySoundWithNAudio(string soundFileName)
        {
            try
            {
                string soundPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, soundFileName);

                if (File.Exists(soundPath))
                {
                    await Task.Run(() =>
                    {
                        using (var audioFile = new AudioFileReader(soundPath))
                        using (var outputDevice = new WaveOutEvent())
                        {
                            outputDevice.Init(audioFile);
                            outputDevice.Play();

                            // รอให้เล่นเสร็จ
                            while (outputDevice.PlaybackState == PlaybackState.Playing)
                            {
                                Thread.Sleep(100);
                            }
                        }
                    });
                    Console.WriteLine($"🔊 เล่นเสียงด้วย NAudio: {soundFileName}");
                }
                else
                {
                    int round = 1;
                    while (round <= 5)
                    {
                        Console.Beep(800, 700);
                        round++;
                        await Task.Delay(1000); // หน่วงเวลา 1 วินาทีระหว่างรอบ
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ เล่นเสียงไม่สำเร็จ: {ex.Message}");
            }
        }

        static async void RegisterHardwareId()
        {
            using (HttpClient client = new HttpClient())
            {
                string computerName = Environment.MachineName;
                await client.PutAsync($"{apiUrl}/api/com-name/{computerName.ToUpper()}/{hwid}", null);
            }
        }

        private static async Task<bool> IsApiAvailable()
        {
            try
            {
                using var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(5);
                var response = await client.GetAsync($"{apiCheckUrl}/{hwid}");
                return response.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }


        static void HideConsoleWindow()
        {
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);
        }

        static string GetHardwareId()
        {
            string cpuId = GetWMIValue("Win32_Processor", "ProcessorId");
            string macAddress = GetMacAddress();

            string combined = $"{cpuId}-{macAddress}";

            using (var sha256 = SHA256.Create())
            {
                byte[] hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(combined));
                return BitConverter.ToString(hash).Replace("-", "").ToLower();
            }
        }

        static string GetWMIValue(string wmiClass, string wmiProperty)
        {
            try
            {
                using var searcher = new ManagementObjectSearcher($"SELECT {wmiProperty} FROM {wmiClass}");
                foreach (var obj in searcher.Get())
                {
                    return obj[wmiProperty]?.ToString() ?? "";
                }
            }
            catch { }
            return "";
        }

        static string GetMacAddress()
        {
            try
            {
                return NetworkInterface
                    .GetAllNetworkInterfaces()
                    .Where(nic => nic.OperationalStatus == OperationalStatus.Up &&
                                  nic.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                    .Select(nic => nic.GetPhysicalAddress().ToString())
                    .FirstOrDefault() ?? "";
            }
            catch { }
            return "";
        }


        static async Task<dynamic> CheckStatus()
        {
            using (HttpClient client = new HttpClient())
            {
                var response = await client.GetAsync($"{apiCheckUrl}/{hwid}");
                if (response.IsSuccessStatusCode)
                {
                    string json = await response.Content.ReadAsStringAsync();
                    if (JsonHelper.IsValidJson(json))
                    {
                        using var doc = JsonDocument.Parse(json);
                        var root = doc.RootElement;
                        if (root.GetProperty("success").GetBoolean())
                        {
                            var data = root.GetProperty("data");
                            return new
                            {
                                id = data.GetProperty("id").GetInt32(),
                                status = data.GetProperty("status").GetString(),
                                hwid = data.GetProperty("hwid").GetString(),
                            };
                        }
                    }
                }
            }
            return null;
        }

        static async void CheckAdminCommands()
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    var response = await client.GetAsync($"{apiUrl}/api/manage-items/command/{itemId}");
                    if (response.IsSuccessStatusCode)
                    {
                        string json = await response.Content.ReadAsStringAsync();
                        if (JsonHelper.IsValidJson(json))
                        {
                            using var doc = JsonDocument.Parse(json);
                            var root = doc.RootElement;
                            if (root.GetProperty("success").GetBoolean())
                            {
                                var data = root.GetProperty("data").EnumerateArray();
                                foreach (var cmd in data)
                                {
                                    int commandId = cmd.GetProperty("id").GetInt32();
                                    string command = cmd.GetProperty("command").GetString();

                                    DeleteCommand(commandId);

                                    switch (command)
                                    {
                                        case "MESSAGE_BOX":
                                            string message = cmd.GetProperty("message").GetString();
                                            if (message?.Length > 0)
                                            {
                                                messageCommand = message;
                                                if (!showingBlockMessage)
                                                    ShowBlockingMessageBox(message, 10);
                                            }
                                            break;

                                        case "SHUTDOWN":
                                            ShutdownPC();
                                            break;

                                        case "RESTART":
                                            RestartPC();
                                            break;
                                    }

                                }

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ พบข้อผิดพลาดระหว่างตรวจสอบคำสั่ง: {ex.Message}");
            }
        }

        static async void DeleteCommand(int commandId)
        {
            using (HttpClient client = new HttpClient())
            {
                var response = await client.DeleteAsync($"{apiUrl}/api/command/{commandId}");
                if (!response.IsSuccessStatusCode)
                {
                    Console.WriteLine($"ไม่สามารถลบคำสั่ง id {commandId}");
                }
            }
        }

        static void ShutdownPC()
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = "shutdown",
                Arguments = "/s /t 0", // /s = shutdown, /t 0 = ทันที
                CreateNoWindow = true,
                UseShellExecute = false
            });
        }

        static void RestartPC()
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = "shutdown",
                Arguments = "/r /t 0", // /r = restart
                CreateNoWindow = true,
                UseShellExecute = false
            });
        }

        static void ShowBlockingMessageBox(string message, int seconds)
        {
            if (Application.MessageLoop)
            {
                // รันใน UI thread โดยตรง
                ShowFormBlocking(message, seconds);
            }
            else
            {
                // เรียกผ่าน Application.Run ถ้าไม่มี UI
                Thread uiThread = new Thread(() =>
                {
                    Application.EnableVisualStyles(); // เฉพาะกรณีต้องการรัน form ใหม่จาก scratch
                    Application.SetCompatibleTextRenderingDefault(false);
                    ShowFormBlocking(message, seconds);
                    Application.Run(); // รัน loop แยกถ้าไม่มีอยู่แล้ว
                });

                uiThread.SetApartmentState(ApartmentState.STA);
                uiThread.Start();
            }
        }

        static void ShowFormBlocking(string message, int seconds)
        {
            Form form = new Form()
            {
                WindowState = FormWindowState.Maximized,
                FormBorderStyle = FormBorderStyle.None,
                StartPosition = FormStartPosition.Manual,
                ControlBox = false,
                ShowInTaskbar = false,
                TopMost = true
            };

            Label label = new Label()
            {
                Text = message,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Tahoma", 32, FontStyle.Bold)
            };

            form.Controls.Add(label);

            var timer = new System.Windows.Forms.Timer();
            timer.Interval = 1000;
            int count = seconds;
            showingBlockMessage = true;
            timer.Tick += (s, e) =>
            {
                count--;
                if (messageCommand.Length > 0)
                {
                    label.Text = messageCommand;
                    messageCommand = "";
                }

                if (count <= 0)
                {
                    showingBlockMessage = false;
                    timer.Stop();
                    form.Close();
                }
            };

            form.Shown += (s, e) =>
            {
                timer.Start();
                form.TopMost = false;
                form.TopMost = true;
                form.BringToFront();
                form.Activate();
            };

            form.TopMost = true;
            form.ShowDialog();
        }
    }
}

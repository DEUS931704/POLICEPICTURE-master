using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Globalization;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;

namespace POLICEPICTURE
{
    internal static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                // 獲取當前IP地址列表
                List<string> ipAddresses = GetAllLocalIPv4Addresses();
                string ipMessage = "檢測到的IP地址:\n" + string.Join("\n", ipAddresses);

                // 檢查是否在允許的網域內
                bool isInAllowedNetwork = IsInAllowedNetwork(ipAddresses);

                // 無論結果如何，都顯示當前IP地址
                if (isInAllowedNetwork)
                {
                    MessageBox.Show($"{ipMessage}\n\n網域驗證成功！您可以使用此應用程式。",
                        "IP地址信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"{ipMessage}\n\n網域驗證失敗！此應用程式只提供新竹市警察局網域內執行。",
                        "IP地址信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    // 網域驗證失敗時退出程式
                    return;
                }

                // 設置默認區域和文化
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("zh-TW");
                Thread.CurrentThread.CurrentCulture = new CultureInfo("zh-TW");

                // 初始化日誌系統
                Logger.Initialize();
                Logger.Log("應用程序啟動");
                Logger.Log($"網域驗證結果: {(isInAllowedNetwork ? "成功" : "失敗")}");
                Logger.Log($"檢測到的IP地址: {string.Join(", ", ipAddresses)}");

                // 啟用視覺樣式
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                // 設置全局異常處理
                Application.ThreadException += Application_ThreadException;
                AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

                // 運行主窗體
                Application.Run(new Form1());
            }
            catch (Exception ex)
            {
                // 記錄致命錯誤
                try
                {
                    Logger.Log($"應用程序發生嚴重錯誤: {ex.Message}\n{ex.StackTrace}", Logger.LogLevel.Error);
                }
                catch
                {
                    // 如果日誌記錄失敗，顯示消息框
                }

                MessageBox.Show($"應用程序發生嚴重錯誤，即將關閉:\n{ex.Message}", "嚴重錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 獲取所有本機IPv4地址
        /// </summary>
        /// <returns>IPv4地址列表</returns>
        private static List<string> GetAllLocalIPv4Addresses()
        {
            List<string> ipAddresses = new List<string>();

            try
            {
                // 方法1: 使用NetworkInterface獲取所有網絡介面的IP
                foreach (NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces())
                {
                    // 只檢查已啟用的介面且非虛擬介面
                    if (ni.OperationalStatus != OperationalStatus.Up ||
                        ni.NetworkInterfaceType == NetworkInterfaceType.Loopback)
                        continue;

                    // 獲取該介面的IP屬性
                    IPInterfaceProperties ipProps = ni.GetIPProperties();

                    // 檢查IPv4地址
                    foreach (UnicastIPAddressInformation ip in ipProps.UnicastAddresses)
                    {
                        // 只收集IPv4地址
                        if (ip.Address.AddressFamily == AddressFamily.InterNetwork)
                        {
                            string ipAddress = ip.Address.ToString();
                            // 避免重複添加
                            if (!ipAddresses.Contains(ipAddress))
                            {
                                ipAddresses.Add(ipAddress);
                            }
                        }
                    }
                }

                // 方法2: 使用DNS獲取本機IP (備用方法)
                if (ipAddresses.Count == 0)
                {
                    string hostName = Dns.GetHostName();
                    IPHostEntry hostEntry = Dns.GetHostEntry(hostName);

                    foreach (IPAddress ip in hostEntry.AddressList)
                    {
                        if (ip.AddressFamily == AddressFamily.InterNetwork)
                        {
                            string ipAddress = ip.ToString();
                            if (!ipAddresses.Contains(ipAddress))
                            {
                                ipAddresses.Add(ipAddress);
                            }
                        }
                    }
                }

                // 如果沒有找到任何IP地址
                if (ipAddresses.Count == 0)
                {
                    ipAddresses.Add("未檢測到IPv4地址");
                }
            }
            catch (Exception ex)
            {
                // 發生異常時添加錯誤信息
                ipAddresses.Add($"檢測IP時發生錯誤: {ex.Message}");
            }

            return ipAddresses;
        }

        /// <summary>
        /// 檢查是否在允許的網域 (10.108.X.X) 內
        /// </summary>
        /// <param name="ipAddresses">要檢查的IP地址列表</param>
        /// <returns>如果在允許的網域內返回true，否則返回false</returns>
        private static bool IsInAllowedNetwork(List<string> ipAddresses)
        {
            try
            {
                // 檢查每個IP地址是否符合10.108.X.X格式
                foreach (string ip in ipAddresses)
                {
                    if (ip.StartsWith("10.108."))
                    {
                        return true;
                    }
                }

                // 如果沒有找到符合條件的IP地址
                return false;
            }
            catch (Exception)
            {
                // 出現錯誤時，為安全起見返回false
                return false;
            }
        }

        /// <summary>
        /// 處理UI線程異常
        /// </summary>
        private static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            Logger.Log($"UI線程異常: {e.Exception.Message}\n{e.Exception.StackTrace}", Logger.LogLevel.Error);
            MessageBox.Show($"應用程序發生錯誤:\n{e.Exception.Message}", "錯誤",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// 處理未捕獲的異常
        /// </summary>
        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var ex = e.ExceptionObject as Exception;
            Logger.Log($"未處理的異常: {ex?.Message ?? "未知錯誤"}\n{ex?.StackTrace ?? "無堆疊信息"}", Logger.LogLevel.Error);

            MessageBox.Show($"應用程序發生嚴重錯誤，即將關閉:\n{ex?.Message ?? "未知錯誤"}", "嚴重錯誤",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
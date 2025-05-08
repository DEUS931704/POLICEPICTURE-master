using System;
using System.Globalization;

namespace POLICEPICTURE
{
    /// <summary>
    /// 日期時間工具類 - 提供各種日期時間相關的輔助功能
    /// </summary>
    public static class DateUtility
    {
        /// <summary>
        /// 西元紀年與民國紀年的年份差
        /// </summary>
        private const int ROC_YEAR_OFFSET = 1911;

        /// <summary>
        /// 將西元日期時間轉換為民國年日期時間字串（不顯示"民國"二字）
        /// </summary>
        /// <param name="dateTime">西元日期時間</param>
        /// <param name="includeTime">是否包含時間部分</param>
        /// <returns>民國年格式的日期時間字串</returns>
        public static string ToRocDateString(DateTime dateTime, bool includeTime = false)
        {
            try
            {
                // 計算民國年 (西元年 - 1911)，但不加"民國"二字
                int rocYear = dateTime.Year - ROC_YEAR_OFFSET;

                // 格式化基本日期，不包含"民國"二字
                string dateString = $"{rocYear} 年 {dateTime.Month} 月 {dateTime.Day} 日";

                // 如果需要包含時間
                if (includeTime)
                {
                    dateString += $" {dateTime.Hour:D2}:{dateTime.Minute:D2}";
                }

                return dateString;
            }
            catch (Exception ex)
            {
                Logger.Log($"轉換民國年日期時出錯: {ex.Message}", Logger.LogLevel.Error);

                // 發生錯誤時使用西元年顯示
                int rocYear = dateTime.Year - ROC_YEAR_OFFSET;
                return $"{rocYear}年{dateTime.Month}月{dateTime.Day}日";
            }
        }

        /// <summary>
        /// 使用DateTimePicker的CustomFormat格式將西元年轉為民國年格式（不顯示"民國"二字）
        /// </summary>
        /// <returns>適用於DateTimePicker的CustomFormat字符串</returns>
        public static string GetRocDateTimePickerFormat()
        {
            // 返回不含"民國"二字的格式
            return "yyy '年' MM '月' dd '日' HH:mm";
        }

        /// <summary>
        /// 將DateTimePicker控件設置為顯示民國年（不顯示"民國"二字）
        /// </summary>
        /// <param name="picker">DateTimePicker控件</param>
        public static void SetupRocDateTimePicker(System.Windows.Forms.DateTimePicker picker)
        {
            if (picker == null) return;

            try
            {
                // 設置台灣文化和日曆
                CultureInfo twCulture = new CultureInfo("zh-TW");
                twCulture.DateTimeFormat.Calendar = new TaiwanCalendar();

                // 顯示民國年但不含"民國"二字
                picker.CustomFormat = "yyy '年' MM '月' dd '日' HH:mm";
                picker.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            }
            catch (Exception ex)
            {
                Logger.Log($"設置民國年DateTimePicker時出錯: {ex.Message}", Logger.LogLevel.Error);
                picker.CustomFormat = "yyyy年MM月dd日 HH:mm";
                picker.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            }
        }

        /// <summary>
        /// 檢查字符串是否為民國年格式
        /// </summary>
        /// <param name="dateString">要檢查的日期字符串</param>
        /// <returns>是否為民國年格式</returns>
        public static bool IsRocDateFormat(string dateString)
        {
            // 檢查是否有年月日格式，無需檢查"民國"二字
            return !string.IsNullOrEmpty(dateString) && dateString.Contains("年") && dateString.Contains("月") && dateString.Contains("日");
        }

        /// <summary>
        /// 嘗試將各種格式的日期字符串解析為DateTime
        /// </summary>
        /// <param name="dateString">日期字符串</param>
        /// <param name="result">解析結果</param>
        /// <returns>是否成功解析</returns>
        public static bool TryParseDateTime(string dateString, out DateTime result)
        {
            result = DateTime.MinValue;

            if (string.IsNullOrWhiteSpace(dateString))
                return false;

            // 嘗試解析民國年格式（不需要有"民國"二字的前綴）
            if (IsRocDateFormat(dateString))
            {
                try
                {
                    // 提取年、月、日
                    string cleanStr = dateString.Replace("民國", "").Replace("年", " ").Replace("月", " ").Replace("日", " ");
                    string[] parts = cleanStr.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    if (parts.Length >= 3)
                    {
                        if (int.TryParse(parts[0], out int rocYear) &&
                            int.TryParse(parts[1], out int month) &&
                            int.TryParse(parts[2], out int day))
                        {
                            // 判斷是西元年還是民國年
                            int year;
                            if (rocYear > 1911) // 假設是西元年
                            {
                                year = rocYear;
                            }
                            else // 假設是民國年
                            {
                                year = rocYear + ROC_YEAR_OFFSET;
                            }

                            // 創建DateTime
                            result = new DateTime(year, month, day);

                            // 如果有時間部分
                            if (parts.Length >= 4 && parts[3].Contains(":"))
                            {
                                string[] timeParts = parts[3].Split(':');
                                if (timeParts.Length >= 2)
                                {
                                    if (int.TryParse(timeParts[0], out int hour) &&
                                        int.TryParse(timeParts[1], out int minute))
                                    {
                                        result = new DateTime(year, month, day, hour, minute, 0);
                                    }
                                }
                            }

                            return true;
                        }
                    }
                }
                catch
                {
                    // 解析失敗，繼續嘗試標準格式
                }
            }

            // 嘗試標準DateTime解析
            return DateTime.TryParse(dateString, out result);
        }
    }
}
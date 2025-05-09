using System;
using System.IO;
using System.Reflection;
using System.Collections.Generic;

namespace POLICEPICTURE
{
    /// <summary>
    /// 嵌入資源管理輔助類 - 用於處理嵌入到執行檔中的資源
    /// </summary>
    public static class EmbeddedResourceHelper
    {
        /// <summary>
        /// 默認資源命名空間
        /// </summary>
        private const string DEFAULT_NAMESPACE = "POLICEPICTURE";

        /// <summary>
        /// 獲取所有嵌入資源的名稱列表
        /// </summary>
        /// <returns>資源名稱列表</returns>
        public static List<string> GetEmbeddedResourceNames()
        {
            List<string> resourceNames = new List<string>();
            Assembly assembly = Assembly.GetExecutingAssembly();

            try
            {
                string[] names = assembly.GetManifestResourceNames();
                resourceNames.AddRange(names);

                Logger.Log($"找到 {names.Length} 個嵌入資源", Logger.LogLevel.Info);
                foreach (string name in names)
                {
                    Logger.Log($"嵌入資源: {name}", Logger.LogLevel.Debug);
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"獲取嵌入資源名稱列表時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
            }

            return resourceNames;
        }

        /// <summary>
        /// 從嵌入資源中提取文件到臨時目錄
        /// </summary>
        /// <param name="resourceName">資源名稱 (可選，默認為 template.docx)</param>
        /// <param name="fileName">輸出檔案名稱 (可選，默認使用資源名稱)</param>
        /// <returns>提取後的文件路徑，失敗則返回 null</returns>
        public static string ExtractResourceToTempFile(string resourceName = "template.docx", string fileName = null)
        {
            try
            {
                // 如果沒有指定檔案名稱，使用資源名稱
                if (string.IsNullOrEmpty(fileName))
                {
                    fileName = resourceName;
                }

                // 添加命名空間前綴（如果資源名稱沒有包含命名空間）
                string fullResourceName = resourceName;
                if (!resourceName.Contains("."))
                {
                    fullResourceName = $"{DEFAULT_NAMESPACE}.{resourceName}";
                }

                // 日誌記錄
                Logger.Log($"嘗試提取嵌入資源: {fullResourceName}", Logger.LogLevel.Info);

                // 獲取程序集
                Assembly assembly = Assembly.GetExecutingAssembly();

                // 嘗試獲取資源流
                using (Stream resourceStream = assembly.GetManifestResourceStream(fullResourceName))
                {
                    // 如果找不到指定的資源名稱，嘗試查找完整的資源名稱列表
                    if (resourceStream == null)
                    {
                        Logger.Log($"找不到嵌入資源: {fullResourceName}，嘗試查找其他可能的資源名稱", Logger.LogLevel.Warning);

                        // 獲取所有資源名稱
                        string[] allResourceNames = assembly.GetManifestResourceNames();

                        // 記錄所有找到的資源名稱
                        Logger.Log($"找到 {allResourceNames.Length} 個嵌入資源:", Logger.LogLevel.Info);
                        foreach (string name in allResourceNames)
                        {
                            Logger.Log($"  - {name}", Logger.LogLevel.Info);

                            // 如果發現包含要查找的資源名稱的項目
                            if (name.EndsWith(resourceName, StringComparison.OrdinalIgnoreCase))
                            {
                                Logger.Log($"找到可能匹配的資源: {name}", Logger.LogLevel.Info);
                                fullResourceName = name;

                                // 嘗試使用這個資源名稱
                                using (Stream foundStream = assembly.GetManifestResourceStream(name))
                                {
                                    if (foundStream != null)
                                    {
                                        // 創建臨時文件路徑
                                        string tempFilePath = Path.Combine(Path.GetTempPath(), $"{Path.GetFileNameWithoutExtension(fileName)}_{Guid.NewGuid().ToString().Substring(0, 8)}{Path.GetExtension(fileName)}");

                                        // 提取資源到臨時文件
                                        using (FileStream fileStream = new FileStream(tempFilePath, FileMode.Create))
                                        {
                                            foundStream.CopyTo(fileStream);
                                        }

                                        Logger.Log($"成功提取資源到臨時文件: {tempFilePath}", Logger.LogLevel.Info);
                                        return tempFilePath;
                                    }
                                }
                            }
                        }

                        Logger.Log("無法找到匹配的嵌入資源", Logger.LogLevel.Error);
                        return null;
                    }

                    // 創建臨時文件路徑
                    string tempPath = Path.Combine(Path.GetTempPath(), $"{Path.GetFileNameWithoutExtension(fileName)}_{Guid.NewGuid().ToString().Substring(0, 8)}{Path.GetExtension(fileName)}");

                    // 提取資源到臨時文件
                    using (FileStream fileStream = new FileStream(tempPath, FileMode.Create))
                    {
                        resourceStream.CopyTo(fileStream);
                    }

                    Logger.Log($"成功提取資源到臨時文件: {tempPath}", Logger.LogLevel.Info);
                    return tempPath;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"提取嵌入資源時發生錯誤: {ex.Message}\n{ex.StackTrace}", Logger.LogLevel.Error);
                return null;
            }
        }

        /// <summary>
        /// 從嵌入資源提取範本文件 (專門用於 template.docx)
        /// </summary>
        /// <returns>範本文件路徑，如果失敗則返回 null</returns>
        public static string ExtractTemplateFile()
        {
            // 嘗試多種可能的資源名稱
            string[] possibleNames = new string[]
            {
                "template.docx",                 // 直接嵌入在根目錄
                "Templates.template.docx",       // 在 Templates 子目錄
                "Resources.template.docx",       // 在 Resources 子目錄
                "Resources.Templates.template.docx", // 在 Resources/Templates 子目錄
            };

            // 嘗試所有可能的名稱
            foreach (string name in possibleNames)
            {
                string fullName = $"{DEFAULT_NAMESPACE}.{name}";
                string templatePath = ExtractResourceToTempFile(fullName);

                if (!string.IsNullOrEmpty(templatePath))
                {
                    return templatePath;
                }
            }

            // 嘗試使用資源名稱列表查找
            List<string> allResources = GetEmbeddedResourceNames();
            foreach (string resource in allResources)
            {
                if (resource.EndsWith("template.docx", StringComparison.OrdinalIgnoreCase))
                {
                    // 找到了匹配的資源
                    Logger.Log($"發現匹配的範本資源: {resource}", Logger.LogLevel.Info);
                    return ExtractResourceToTempFile(resource);
                }
            }

            // 如果都無法找到，返回 null
            Logger.Log("無法找到嵌入的範本資源", Logger.LogLevel.Error);
            return null;
        }
    }
}
using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using POLICEPICTURE;
using System.Net;

namespace POLICEPICTURE
{
    /// <summary>
    /// 使用 Office Interop 操作 Word 文檔的高級幫助類
    /// </summary>
    public class WordInteropHelper
    {
        // 進度報告間隔 - 處理大量照片時的性能優化
        private const int PROGRESS_REPORT_INTERVAL = 10;

        /// <summary>
        /// 表格信息類，用於保存表格和其中的 %%PICTURE%% 標記位置
        /// </summary>
        public class TableInfo
        {
            public Table Table { get; set; }
            public List<CellMarkerInfo> PictureMarkers { get; set; } = new List<CellMarkerInfo>();
        }

        /// <summary>
        /// 單元格標記信息類，用於保存標記所在的單元格和位置
        /// </summary>
        public class CellMarkerInfo
        {
            public Cell Cell { get; set; }
            public int Row { get; set; }
            public int Column { get; set; }
            public Range MarkerRange { get; set; }
        }

        /// <summary>
        /// 從模板生成文檔並處理照片
        /// </summary>
        public static async System.Threading.Tasks.Task<bool> GenerateDocumentAsync(
            string templatePath,
            string outputPath,
            string unit,
            string caseDesc,
            string time,
            string location,
            string photographer,
            IReadOnlyList<PhotoItem> photos,
            ProgressReportHandler progressReport = null)
        {
            // 使用 Task.Run 在後台線程執行耗時操作
            return await System.Threading.Tasks.Task.Run(() =>
            {
                // Word 應用實例、文檔和範圍變數
                Microsoft.Office.Interop.Word.Application wordApp = null;
                Microsoft.Office.Interop.Word.Document doc = null;
                object missing = System.Reflection.Missing.Value;

                try
                {
                    // 報告進度 - 10%
                    progressReport?.Invoke(10, "準備生成文檔...");

                    // 檢查照片數量，提前顯示警告
                    if (photos.Count > 100)
                    {
                        Logger.Log($"警告：處理大量照片({photos.Count}張)可能需要較長時間", Logger.LogLevel.Warning);
                        progressReport?.Invoke(10, $"準備處理 {photos.Count} 張照片，這可能需要較長時間...");
                    }

                    // 驗證模板路徑
                    if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
                    {
                        throw new FileNotFoundException("找不到範本檔案", templatePath);
                    }

                    // 確保輸出目錄存在
                    string outputDir = Path.GetDirectoryName(outputPath);
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // 報告進度 - 20%
                    progressReport?.Invoke(20, "初始化 Word 應用...");

                    // 創建 Word 應用實例
                    wordApp = new Microsoft.Office.Interop.Word.Application();
                    wordApp.Visible = false; // 隱藏 Word 應用

                    // 報告進度 - 25%
                    progressReport?.Invoke(25, "加載範本文檔...");

                    // 優化：設置打開文檔選項以提高性能
                    object readOnly = false;
                    object isVisible = false;
                    object openAndRepair = false;

                    // 打開模板文件，使用優化參數
                    doc = wordApp.Documents.Open(
                        templatePath,
                        ReadOnly: readOnly,
                        Visible: isVisible,
                        OpenAndRepair: openAndRepair);

                    // 報告進度 - 30%
                    progressReport?.Invoke(30, "填充文檔內容...");

                    // 格式化時間，只保留年月日，使用西元年格式
                    string formattedTime = string.Empty;
                    if (!string.IsNullOrEmpty(time))
                    {
                        try
                        {
                            // 嘗試解析時間字符串
                            if (DateUtility.TryParseDateTime(time, out DateTime dateTime))
                            {
                                // 使用民國年顯示，但不顯示"民國"二字
                                int rocYear = dateTime.Year - 1911;
                                formattedTime = $"{rocYear} 年 {dateTime.Month} 月 {dateTime.Day} 日";
                            }
                            else
                            {
                                // 如果無法解析，則使用原始字符串
                                formattedTime = time;
                            }
                        }
                        catch
                        {
                            formattedTime = time;
                        }
                    }
                    

                    // 批量替換文檔中的佔位符，提高性能
                   Dictionary<string, string> replacements = new Dictionary<string, string>
                    {
                        // 處理單位顯示，確保刑事警察大隊科偵隊等可以正確顯示
                        { "%%UNIT%%", unit?.Replace(" ", "\n") ?? string.Empty }, // 主單位和子單位間使用換行符
                        { "%%CASE%%", caseDesc ?? string.Empty },
                        { "%%TIME%%", formattedTime },
                        { "%%ADDRESS%%", location ?? string.Empty },
                        { "%%NAME%%", photographer ?? string.Empty },
                        { "%%Description%%", string.Empty },
                        // 添加其他需要替換的佔位符
                        { "%%DATE%%", string.Empty },
                        { "%%SERIAL%%", string.Empty },
                        { "%%NUMBER%%", string.Empty }
                    };

                    // 批量替換所有佔位符
                    BatchReplaceTextInDocument(doc, replacements);

                    // 報告進度 - 40%
                    progressReport?.Invoke(40, "查找照片標記和表格...");

                    // 使用新方法查找包含 %%PICTURE%% 標記的表格
                    TableInfo templateTableInfo = FindTableWithPictureMarkers(doc);

                    if (templateTableInfo != null && templateTableInfo.Table != null && templateTableInfo.PictureMarkers.Count > 0)
                    {
                        Logger.Log($"找到包含 %%PICTURE%% 標記的表格，標記數量: {templateTableInfo.PictureMarkers.Count}", Logger.LogLevel.Info);

                        // 如果有照片並且找到了標記表格，處理照片
                        if (photos.Count > 0)
                        {
                            // 報告進度 - 50%
                            progressReport?.Invoke(50, "處理照片...");

                            // 使用找到的表格和標記處理照片 - 優化版本
                            ProcessPhotosInTemplateTableOptimized(doc, templateTableInfo, photos, progressReport);
                        }
                    }
                    else
                    {
                        Logger.Log("未找到包含 %%PICTURE%% 標記的表格，使用替代方法", Logger.LogLevel.Warning);

                        // 如果沒有找到表格但有照片需要處理，使用替代方法
                        if (photos.Count > 0)
                        {
                            // 報告進度 - 45%
                            progressReport?.Invoke(45, "使用替代方法處理照片...");
                            ProcessPhotosAlternativeOptimized(doc, photos, progressReport);
                        }
                    }

                    // 最後，刪除文檔中任何剩餘的 %%PICTURE%% 標記
                    ReplaceTextInDocument(doc, "%%PICTURE%%", string.Empty);

                    // 報告進度 - 90%
                    progressReport?.Invoke(90, "保存文檔...");

                    // 保存文檔
                    doc.SaveAs2(outputPath);

                    // 報告進度 - 100%
                    progressReport?.Invoke(100, "文件生成完成");

                    return true;
                }
                catch (Exception ex)
                {
                    Logger.Log($"使用 Word Interop 生成文件時發生錯誤: {ex.Message}\n{ex.StackTrace}", Logger.LogLevel.Error);
                    MessageBox.Show($"生成文件時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                finally
                {
                    // 清理資源
                    if (doc != null)
                    {
                        doc.Close(SaveChanges: false);
                        Marshal.ReleaseComObject(doc);
                    }

                    if (wordApp != null)
                    {
                        wordApp.Quit();
                        Marshal.ReleaseComObject(wordApp);
                    }

                    // 強制垃圾回收，確保 COM 對象被釋放
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            });
        }

        /// <summary>
        /// 找到包含 %%PICTURE%% 標記的表格及其標記位置
        /// </summary>
        private static TableInfo FindTableWithPictureMarkers(Document doc)
        {
            try
            {
                // 遍歷所有表格
                foreach (Table table in doc.Tables)
                {
                    TableInfo tableInfo = new TableInfo { Table = table };

                    // 檢查表格中的每個單元格
                    for (int rowIndex = 1; rowIndex <= table.Rows.Count; rowIndex++)
                    {
                        for (int colIndex = 1; colIndex <= table.Columns.Count; colIndex++)
                        {
                            try
                            {
                                Cell cell = table.Cell(rowIndex, colIndex);

                                // 檢查單元格中是否包含標記
                                if (cell.Range.Text.Contains("%%PICTURE%%"))
                                {
                                    // 創建標記範圍
                                    int start = cell.Range.Start + cell.Range.Text.IndexOf("%%PICTURE%%");
                                    int end = start + "%%PICTURE%%".Length;
                                    Range markerRange = doc.Range(start, end);

                                    // 添加到標記列表
                                    tableInfo.PictureMarkers.Add(new CellMarkerInfo
                                    {
                                        Cell = cell,
                                        Row = rowIndex,
                                        Column = colIndex,
                                        MarkerRange = markerRange
                                    });

                                    Logger.Log($"在表格第 {rowIndex} 行第 {colIndex} 列找到 %%PICTURE%% 標記", Logger.LogLevel.Debug);
                                }
                            }
                            catch (Exception ex)
                            {
                                Logger.Log($"檢查表格單元格時出錯: {ex.Message}", Logger.LogLevel.Warning);
                            }
                        }
                    }

                    // 如果找到包含標記的表格，返回該表格信息
                    if (tableInfo.PictureMarkers.Count > 0)
                    {
                        return tableInfo;
                    }
                }

                return null; // 沒有找到包含標記的表格
            }
            catch (Exception ex)
            {
                Logger.Log($"查找包含 %%PICTURE%% 標記的表格時出錯: {ex.Message}", Logger.LogLevel.Error);
                return null;
            }
        }

        /// <summary>
        /// 在模板表格中處理照片，根據照片數量複製表格 - 優化版本
        /// </summary>
        private static void ProcessPhotosInTemplateTableOptimized(Document doc, TableInfo templateTableInfo, IReadOnlyList<PhotoItem> photos, ProgressReportHandler progressReport)
        {
            try
            {
                // 獲取原始表格和標記
                Table originalTable = templateTableInfo.Table;
                int markersPerTable = templateTableInfo.PictureMarkers.Count;

                if (markersPerTable == 0)
                {
                    Logger.Log("標記信息異常，無法處理照片", Logger.LogLevel.Error);
                    return;
                }

                // 計算需要的表格數量
                int tablesNeeded = (int)Math.Ceiling((double)photos.Count / markersPerTable);
                Logger.Log($"照片數量: {photos.Count}, 每個表格標記數: {markersPerTable}, 需要表格數: {tablesNeeded}", Logger.LogLevel.Info);

                // 優化：如果表格數量過多，顯示額外的進度信息
                if (tablesNeeded > 20)
                {
                    progressReport?.Invoke(50, $"需要創建 {tablesNeeded} 個表格，請耐心等待...");
                }

                // 複製表格 (如果需要)
                List<Table> allTables = new List<Table> { originalTable };

                // 優化：分批處理表格複製，減少每次報告進度
                int tableProgressStep = Math.Max(1, tablesNeeded / 20); // 每 5% 的表格創建報告一次進度

                for (int i = 1; i < tablesNeeded; i++)
                {
                    // 僅在固定間隔報告進度，減少 UI 更新頻率
                    if (i % tableProgressStep == 0 || i == tablesNeeded - 1)
                    {
                        int progressPercent = 50 + (i * 5 / tablesNeeded);
                        progressReport?.Invoke(progressPercent, $"創建表格 {i + 1}/{tablesNeeded}...");
                    }

                    // 插入分頁符
                    Range endRange = doc.Content.Duplicate;
                    endRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    endRange.InsertBreak(WdBreakType.wdPageBreak);

                    // 優化：複製表格 - 使用更高效的方法
                    originalTable.Range.Copy();
                    endRange = doc.Content.Duplicate;
                    endRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    endRange.Paste();

                    // 添加到表格列表
                    allTables.Add(doc.Tables[doc.Tables.Count]);

                    // 僅對部分表格記錄日誌，避免過多日誌
                    if (i % 10 == 0 || i == tablesNeeded - 1)
                    {
                        Logger.Log($"已創建第 {i + 1}/{tablesNeeded} 個照片表格", Logger.LogLevel.Info);
                    }
                }

                // 報告照片處理進度
                progressReport?.Invoke(55, $"開始處理 {photos.Count} 張照片...");

                // 處理所有照片
                int photoIndex = 0;
                int lastReportedProgress = 55;

                // 優化：設置照片處理的進度報告間隔
                int photosProgressInterval = Math.Max(1, photos.Count / 40); // 每 1% 的照片處理報告一次進度

                // 處理每個表格
                for (int tableIndex = 0; tableIndex < allTables.Count && photoIndex < photos.Count; tableIndex++)
                {
                    Table currentTable = allTables[tableIndex];

                    // 為每個表格重新查找標記
                    List<CellMarkerInfo> tableMarkers = new List<CellMarkerInfo>();

                    // 如果是第一個表格，使用已找到的標記
                    if (tableIndex == 0)
                    {
                        tableMarkers = templateTableInfo.PictureMarkers;
                    }
                    // 否則重新查找表格中的標記
                    else
                    {
                        for (int rowIndex = 1; rowIndex <= currentTable.Rows.Count; rowIndex++)
                        {
                            for (int colIndex = 1; colIndex <= currentTable.Columns.Count; colIndex++)
                            {
                                try
                                {
                                    Cell cell = currentTable.Cell(rowIndex, colIndex);

                                    if (cell.Range.Text.Contains("%%PICTURE%%"))
                                    {
                                        int start = cell.Range.Start + cell.Range.Text.IndexOf("%%PICTURE%%");
                                        int end = start + "%%PICTURE%%".Length;
                                        Range markerRange = doc.Range(start, end);

                                        tableMarkers.Add(new CellMarkerInfo
                                        {
                                            Cell = cell,
                                            Row = rowIndex,
                                            Column = colIndex,
                                            MarkerRange = markerRange
                                        });
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Logger.Log($"檢查複製表格單元格時出錯: {ex.Message}", Logger.LogLevel.Warning);
                                }
                            }
                        }
                    }

                    // 在當前表格中處理照片
                    for (int markerIndex = 0; markerIndex < tableMarkers.Count && photoIndex < photos.Count; markerIndex++)
                    {
                        var markerInfo = tableMarkers[markerIndex];
                        var photo = photos[photoIndex];

                        // 優化：減少進度報告頻率，僅在固定間隔報告
                        if (photoIndex % photosProgressInterval == 0 || photoIndex == photos.Count - 1)
                        {
                            int progressValue = 55 + (photoIndex * 35 / photos.Count);

                            // 避免重複報告相同進度
                            if (progressValue > lastReportedProgress)
                            {
                                progressReport?.Invoke(progressValue, $"處理照片 {photoIndex + 1}/{photos.Count}...");
                                lastReportedProgress = progressValue;
                            }
                        }

                        try
                        {
                            // 處理照片
                            ProcessPhotoInCell(markerInfo.Cell, markerInfo.MarkerRange, photo);

                            // 僅對部分照片記錄詳細日誌
                            if (photoIndex % PROGRESS_REPORT_INTERVAL == 0 || photoIndex == photos.Count - 1)
                            {
                                Logger.Log($"成功處理照片 {photoIndex + 1}/{photos.Count}: {Path.GetFileName(photo.FilePath)}", Logger.LogLevel.Debug);
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Log($"處理照片 {photoIndex + 1} 時出錯: {ex.Message}", Logger.LogLevel.Error);
                        }

                        photoIndex++;
                    }
                }

                // 最終進度報告
                progressReport?.Invoke(85, $"完成處理 {photoIndex} 張照片");
            }
            catch (Exception ex)
            {
                Logger.Log($"在模板表格中處理照片時出錯: {ex.Message}\n{ex.StackTrace}", Logger.LogLevel.Error);
            }
        }

        /// <summary>
        /// 在單元格中處理照片，針對垂直照片進行特殊處理
        /// </summary>
        private static void ProcessPhotoInCell(Cell cell, Range markerRange, PhotoItem photo)
        {
            try
            {
                // 確保照片文件存在
                if (!File.Exists(photo.FilePath))
                {
                    Logger.Log($"照片文件不存在: {photo.FilePath}", Logger.LogLevel.Error);
                    markerRange.Text = "[照片文件不存在]";
                    return;
                }

                // 添加照片描述
                if (!string.IsNullOrEmpty(photo.Description))
                {
                    // 在標記前插入描述
                    Range descriptionRange = markerRange.Duplicate;
                    descriptionRange.Collapse(WdCollapseDirection.wdCollapseStart);
                    descriptionRange.InsertBefore(photo.Description + "\r\n");
                    descriptionRange.Bold = 1; // 加粗描述
                }

                // 獲取單元格的大小以限制圖片尺寸
                float maxWidth = 400; // 默認最大寬度
                float maxHeight = 300; // 默認最大高度

                try
                {
                    // 嘗試獲取單元格寬度和高度
                    float cellWidth = (float)cell.Width;
                    if (cellWidth > 0)
                    {
                        maxWidth = cellWidth * 0.8f; // 使用單元格寬度的 80%
                    }

                    // 單元格高度在 Word 中可能難以獲取
                    // 使用保守估計的高度
                    maxHeight = maxWidth * 0.75f;
                }
                catch (Exception ex)
                {
                    Logger.Log($"獲取單元格尺寸時出錯: {ex.Message}", Logger.LogLevel.Warning);
                }

                // 確保尺寸不會太小
                maxWidth = Math.Max(maxWidth, 150);
                maxHeight = Math.Max(maxHeight, 120);

                // 計算等比例縮放後的尺寸，並判斷是否為垂直照片
                bool isVertical = false;
                float originalWidth = 0;
                float originalHeight = 0;

                // 先獲取照片原始尺寸
                using (System.Drawing.Image img = System.Drawing.Image.FromFile(photo.FilePath))
                {
                    originalWidth = img.Width;
                    originalHeight = img.Height;
                    isVertical = img.Height > img.Width;

                    Logger.Log($"照片 {Path.GetFileName(photo.FilePath)} 尺寸: {img.Width}x{img.Height}, " +
                              (isVertical ? "垂直照片" : "水平照片"), Logger.LogLevel.Debug);
                }

                // 對於垂直照片，使用不同的處理策略 - 預先縮放
                float ratio;
                float finalWidth, finalHeight;

                if (isVertical)
                {
                    Logger.Log($"處理垂直照片: {Path.GetFileName(photo.FilePath)}", Logger.LogLevel.Info);

                    // 對於非常細長的垂直照片，使用特別的縮放比例
                    if (originalHeight > originalWidth * 2) // 如果高度超過寬度的2倍
                    {
                        // 增加允許的最大高度，但設定上限
                        maxHeight = Math.Min(maxHeight * 1.5f, 450);
                    }

                    // 使用修改後的縮放比例
                    ratio = Math.Min(maxWidth / originalWidth, maxHeight / originalHeight);
                    // 不放大圖片
                    ratio = Math.Min(ratio, 1.0f);

                    // 計算最終尺寸
                    finalWidth = originalWidth * ratio;
                    finalHeight = originalHeight * ratio;
                }
                else
                {
                    // 水平照片正常處理
                    ratio = Math.Min(maxWidth / originalWidth, maxHeight / originalHeight);

                    // 不放大圖片
                    ratio = Math.Min(ratio, 1.0f);

                    finalWidth = originalWidth * ratio;
                    finalHeight = originalHeight * ratio;
                }

                // 插入圖片
                markerRange.Text = ""; // 清除標記
                InlineShape shape = markerRange.InlineShapes.AddPicture(
                    FileName: photo.FilePath,
                    LinkToFile: false,
                    SaveWithDocument: true);

                // 設置圖片尺寸
                shape.Width = finalWidth;
                shape.Height = finalHeight;

                // 居中圖片
                shape.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            }
            catch (Exception ex)
            {
                Logger.Log($"在單元格中處理照片時出錯: {ex.Message}", Logger.LogLevel.Error);
                markerRange.Text = $"[照片錯誤: {ex.Message}]";
                markerRange.Bold = 1;
                markerRange.Font.Color = WdColor.wdColorRed;
            }
        }

        /// <summary>
        /// 安全檢查表格中的單元格是否存在
        /// </summary>
        private static bool SafeCheckTableCell(Table table, int rowIndex, int colIndex)
        {
            try
            {
                if (table == null || rowIndex <= 0 || colIndex <= 0)
                    return false;

                // 檢查行是否存在
                if (rowIndex > table.Rows.Count)
                    return false;

                // 獲取行
                Row row = table.Rows[rowIndex];
                if (row == null)
                    return false;

                // 檢查列是否存在
                if (colIndex > row.Cells.Count)
                    return false;

                // 檢查單元格是否存在
                Cell cell = row.Cells[colIndex];
                return (cell != null);
            }
            catch
            {
                // 如果出現任何異常，表示單元格訪問有問題
                return false;
            }
        }

        /// <summary>
        /// 安全獲取表格單元格
        /// </summary>
        private static Cell SafeGetTableCell(Table table, int rowIndex, int colIndex)
        {
            try
            {
                if (!SafeCheckTableCell(table, rowIndex, colIndex))
                    return null;

                return table.Cell(rowIndex, colIndex);
            }
            catch (Exception ex)
            {
                Logger.Log($"安全獲取表格單元格時出錯: 行={rowIndex}, 列={colIndex}, 錯誤:{ex.Message}", Logger.LogLevel.Warning);
                return null;
            }
        }

        /// <summary>
        /// 使用替代方法處理照片 - 在文檔末尾添加照片表格（優化版本）
        /// </summary>
        private static bool ProcessPhotosAlternativeOptimized(Document doc, IReadOnlyList<PhotoItem> photos, ProgressReportHandler progressReport)
        {
            try
            {
                if (photos.Count == 0)
                {
                    return true; // 沒有照片要處理
                }

                // 報告狀態
                Logger.Log($"使用替代方法處理 {photos.Count} 張照片", Logger.LogLevel.Info);

                // 將游標移動到文檔末尾
                doc.Content.Collapse(WdCollapseDirection.wdCollapseEnd);

                // 添加一個換頁符
                doc.Content.InsertBreak(WdBreakType.wdPageBreak);

                // 添加照片標題
                Paragraph titlePara = doc.Content.Paragraphs.Add();
                titlePara.Range.Text = "案件照片";
                titlePara.Range.Bold = 1;
                titlePara.Range.Font.Size = 16;
                titlePara.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                titlePara.Format.SpaceAfter = 12;
                titlePara.Range.InsertParagraphAfter();

                // 計算需要的表格數量和行數
                int photosPerRow = 2; // 每行2張照片
                int rowCount = (int)Math.Ceiling((double)photos.Count / photosPerRow);

                progressReport?.Invoke(50, $"創建照片表格（{rowCount} 行 × 2 列）...");

                // 創建表格 - 每行2列
                Table photoTable = doc.Tables.Add(doc.Content.Paragraphs.Last.Range, rowCount, photosPerRow);
                photoTable.Borders.Enable = 1;
                photoTable.AllowAutoFit = true;

                // 設置表格格式
                photoTable.PreferredWidth = 500;
                photoTable.Columns[1].PreferredWidth = 250;
                photoTable.Columns[2].PreferredWidth = 250;

                // 批量處理照片
                progressReport?.Invoke(55, $"開始處理 {photos.Count} 張照片...");

                // 優化：設置進度報告間隔
                int progressInterval = Math.Max(1, photos.Count / 40);
                int lastReportedProgress = 55;

                // 插入照片到表格
                for (int i = 0; i < photos.Count; i++)
                {
                    // 計算行列位置
                    int currentRow = (i / photosPerRow) + 1;
                    int currentCol = (i % photosPerRow) + 1;

                    // 優化：減少進度報告頻率
                    if (i % progressInterval == 0 || i == photos.Count - 1)
                    {
                        int progressValue = 55 + (i * 35 / photos.Count);

                        // 避免重複報告相同進度
                        if (progressValue > lastReportedProgress)
                        {
                            progressReport?.Invoke(progressValue, $"處理照片 {i + 1}/{photos.Count}...");
                            lastReportedProgress = progressValue;
                        }
                    }

                    var photo = photos[i];

                    try
                    {
                        Cell cell = photoTable.Cell(currentRow, currentCol);

                        // 添加照片描述
                        if (!string.IsNullOrEmpty(photo.Description))
                        {
                            Paragraph descPara = cell.Range.Paragraphs.First;
                            descPara.Range.Text = photo.Description;
                            descPara.Range.Bold = 1;
                            descPara.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            descPara.Range.InsertParagraphAfter();
                        }

                        // 添加照片
                        InlineShape shape = cell.Range.InlineShapes.AddPicture(
                            FileName: photo.FilePath,
                            LinkToFile: false,
                            SaveWithDocument: true);

                        // 獲取照片原始尺寸並計算適當的尺寸
                        float finalWidth, finalHeight;
                        bool isVertical = false;

                        using (System.Drawing.Image img = System.Drawing.Image.FromFile(photo.FilePath))
                        {
                            isVertical = img.Height > img.Width;

                            // 調整照片大小
                            float maxWidth = 230; // 單元格寬度減去邊距
                            float maxHeight = 180;

                            float ratio;
                            // 根據照片方向選擇適當的縮放比例
                            ratio = Math.Min(maxWidth / img.Width, maxHeight / img.Height);

                            // 不放大圖片
                            ratio = Math.Min(ratio, 1.0f);

                            // 計算最終尺寸
                            finalWidth = img.Width * ratio;
                            finalHeight = img.Height * ratio;

                            if (isVertical)
                            {
                                Logger.Log($"處理垂直照片: {Path.GetFileName(photo.FilePath)}", Logger.LogLevel.Info);
                            }
                        }

                        // 設置圖片尺寸
                        shape.Width = finalWidth;
                        shape.Height = finalHeight;

                        // 設置居中
                        shape.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                        // 僅對部分照片記錄成功信息
                        if (i % PROGRESS_REPORT_INTERVAL == 0 || i == photos.Count - 1)
                        {
                            Logger.Log($"成功添加照片 {i + 1}/{photos.Count}: {Path.GetFileName(photo.FilePath)}", Logger.LogLevel.Debug);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"處理照片 {i + 1} 時發生錯誤: {ex.Message}", Logger.LogLevel.Error);

                        // 獲取單元格顯示錯誤信息
                        try
                        {
                            Cell cell = photoTable.Cell(currentRow, currentCol);
                            cell.Range.Text = $"[照片錯誤: {Path.GetFileName(photo.FilePath)}]";
                        }
                        catch
                        {
                            // 忽略獲取單元格的錯誤
                        }
                    }
                }

                // 最終進度報告
                progressReport?.Invoke(85, "完成照片處理");

                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"使用替代方法處理照片時發生錯誤: {ex.Message}\n{ex.StackTrace}", Logger.LogLevel.Error);
                return false;
            }
        }

        /// <summary>
        /// 批量替換文檔中的文本（優化版本）
        /// </summary>
        private static void BatchReplaceTextInDocument(Document doc, Dictionary<string, string> replacements)
        {
            if (doc == null || replacements == null || replacements.Count == 0)
            {
                return;
            }

            try
            {
                // 嘗試使用標準方法進行替換
                try
                {
                    Range range = doc.Content;

                    // 對所有替換項逐一進行處理
                    foreach (var replacement in replacements)
                    {
                        range.Find.ClearFormatting();
                        range.Find.Replacement.ClearFormatting();
                        range.Find.Text = replacement.Key;
                        range.Find.Replacement.Text = replacement.Value;
                        range.Find.Execute(Replace: WdReplace.wdReplaceAll);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"使用標準替換方法失敗: {ex.Message}，嘗試替代方法", Logger.LogLevel.Warning);

                    // 使用替代方法逐段替換
                    foreach (var replacement in replacements)
                    {
                        // 替換段落中的文本
                        foreach (Paragraph para in doc.Paragraphs)
                        {
                            if (para.Range.Text.Contains(replacement.Key))
                            {
                                para.Range.Text = para.Range.Text.Replace(replacement.Key, replacement.Value);
                            }
                        }

                        // 替換表格中的文本
                        foreach (Table table in doc.Tables)
                        {
                            foreach (Row row in table.Rows)
                            {
                                foreach (Cell cell in row.Cells)
                                {
                                    if (cell.Range.Text.Contains(replacement.Key))
                                    {
                                        cell.Range.Text = cell.Range.Text.Replace(replacement.Key, replacement.Value);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"批量替換文本時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
            }
        }

        /// <summary>
        /// 在文檔中替換指定文本
        /// </summary>
        private static void ReplaceTextInDocument(Document doc, string findText, string replaceText)
        {
            if (doc == null || string.IsNullOrEmpty(findText))
            {
                return;
            }

            try
            {
                // 嘗試使用常規 Find 方法
                try
                {
                    Range range = doc.Content;
                    range.Find.ClearFormatting();
                    range.Find.Replacement.ClearFormatting();
                    range.Find.Text = findText;
                    range.Find.Replacement.Text = replaceText;
                    range.Find.Execute(Replace: WdReplace.wdReplaceAll);
                }
                catch (Exception ex)
                {
                    Logger.Log($"使用標準替換方法失敗: {ex.Message}，嘗試替代方法", Logger.LogLevel.Warning);

                    // 如果標準方法失敗，使用替代方法逐段替換
                    // 替換段落中的文本
                    foreach (Paragraph para in doc.Paragraphs)
                    {
                        if (para.Range.Text.Contains(findText))
                        {
                            para.Range.Text = para.Range.Text.Replace(findText, replaceText);
                        }
                    }

                    // 替換表格中的文本
                    foreach (Table table in doc.Tables)
                    {
                        foreach (Row row in table.Rows)
                        {
                            foreach (Cell cell in row.Cells)
                            {
                                if (cell.Range.Text.Contains(findText))
                                {
                                    cell.Range.Text = cell.Range.Text.Replace(findText, replaceText);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"替換文本時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
            }
        }
    }
}
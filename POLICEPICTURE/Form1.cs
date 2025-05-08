using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using POLICEPICTURE.Properties;
using System.Globalization;

namespace POLICEPICTURE
{
    // 這是 Form1.cs 中與照片管理相關的部分，包含擴展的照片處理功能
    public partial class Form1 : Form
    {
        // 以下是照片管理相關的方法

        // 新增這些成員變數
        private const string APP_VERSION = "1.0.5"; // 應用程式版本常數
        private readonly UserSettings settings; // 使用者設定
        private readonly ErrorProvider errorProvider; // 錯誤提供者
        private ProgressForm progressForm; // 進度表單

        // 定義單位數據結構，用於存儲單位對應關係
        private readonly Dictionary<string, string[]> unitMapping = new Dictionary<string, string[]>()
        {
            { "刑事警察大隊", new string[] { "偵一隊", "偵二隊", "偵三隊", "科偵隊" } },
            { "第一分局", new string[] { "偵查隊", "西門所", "北門所", "樹林頭所", "南寮所", "湳雅所" } },
            { "第二分局", new string[] { "偵查隊", "東門所", "東勢所", "埔頂所", "關東橋所", "文華所" } },
            { "第三分局", new string[] { "偵查隊", "香山所", "南門所", "朝山所", "青草湖所", "中華所" } }
        };

        public Form1()
        {
            InitializeComponent();

            // 初始化變數
            errorProvider = new ErrorProvider(this);
            settings = UserSettings.Load();

            // 設置列表視圖
            SetupListView();

            // 設定 DateTimePicker 使用民國年但不顯示"民國"字樣
            // 設置台灣文化
            CultureInfo taiwanCulture = new CultureInfo("zh-TW");
            taiwanCulture.DateTimeFormat.Calendar = new TaiwanCalendar();

            // 設置日期選擇器的屬性
            dtpDateTime.CustomFormat = "yyy '年' MM '月' dd '日' HH:mm";
            dtpDateTime.Format = DateTimePickerFormat.Custom;

            // 訂閱照片管理器的事件
            PhotoManager.Instance.PhotosChanged += PhotoManager_PhotosChanged;

            // 選擇第一個大單位
            if (cmbMainUnit.Items.Count > 0)
            {
                cmbMainUnit.SelectedIndex = 0;
            }

            // 從設定中填充表單
            if (!string.IsNullOrEmpty(settings.LastUnit))
            {
                // 嘗試從保存的完整單位名稱中拆分出大單位和小單位
                string[] unitParts = settings.LastUnit.Split(new char[] { ' ' }, 2);

                // 設置大單位
                if (unitParts.Length > 0)
                {
                    int index = cmbMainUnit.Items.IndexOf(unitParts[0]);
                    if (index >= 0)
                    {
                        cmbMainUnit.SelectedIndex = index;

                        // 設置小單位
                        if (unitParts.Length > 1 && cmbSubUnit.Items.Contains(unitParts[1]))
                        {
                            cmbSubUnit.SelectedItem = unitParts[1];
                        }
                    }
                }
            }

            if (!string.IsNullOrEmpty(settings.LastPhotographer))
                txtPhotographer.Text = settings.LastPhotographer;

            // 更新狀態列
            UpdateStatusBar("應用程式就緒");
        }

        // 修改事件處理方法
        private void CmbMainUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 清空小單位下拉選單
            cmbSubUnit.Items.Clear();

            // 獲取選中的大單位
            string selectedMainUnit = cmbMainUnit.SelectedItem as string;

            // 如果選中了有效的大單位，則加載對應的小單位
            if (!string.IsNullOrEmpty(selectedMainUnit) && unitMapping.ContainsKey(selectedMainUnit))
            {
                cmbSubUnit.Items.AddRange(unitMapping[selectedMainUnit]);

                // 預選第一個小單位
                if (cmbSubUnit.Items.Count > 0)
                {
                    cmbSubUnit.SelectedIndex = 0;
                }
            }

            // 更新狀態
            UpdateUnitStatus();
        }
        private void CmbSubUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 更新狀態
            UpdateUnitStatus();
        }

        // 修改UpdateUnitStatus方法
        private void UpdateUnitStatus()
        {
            // 更新狀態欄顯示當前選擇的單位
            string mainUnit = cmbMainUnit.SelectedItem as string ?? "";
            string subUnit = cmbSubUnit.SelectedItem as string ?? "";

            if (!string.IsNullOrEmpty(mainUnit) && !string.IsNullOrEmpty(subUnit))
            {
                UpdateStatusBar($"當前選擇單位: {mainUnit} {subUnit}");
            }
        }

        /// <summary>
        /// 設置列表視圖
        /// </summary>
        private void SetupListView()
        {
            // 啟用整行選擇和網格線
            lvPhotos.FullRowSelect = true;
            lvPhotos.GridLines = true;

            // 啟用虛擬模式以提高性能
            lvPhotos.VirtualMode = true;
            lvPhotos.RetrieveVirtualItem += LvPhotos_RetrieveVirtualItem;

            // 啟用拖放功能，用於照片排序
            lvPhotos.AllowDrop = true;
            lvPhotos.ItemDrag += LvPhotos_ItemDrag;
            lvPhotos.DragEnter += LvPhotos_DragEnter;
            lvPhotos.DragDrop += LvPhotos_DragDrop;

            // 添加右鍵菜單
            lvPhotos.ContextMenuStrip = CreatePhotoContextMenu();
        }

        /// <summary>
        /// 創建照片列表的右鍵菜單
        /// </summary>
        private ContextMenuStrip CreatePhotoContextMenu()
        {
            ContextMenuStrip menu = new ContextMenuStrip();

            // 添加照片
            ToolStripMenuItem addItem = new ToolStripMenuItem("添加照片");
            addItem.Image = SystemIcons.Application.ToBitmap(); // 或使用自定義圖標
            addItem.Click += (s, e) => BtnAddPhoto_Click(s, e);
            menu.Items.Add(addItem);

            // 移除照片
            ToolStripMenuItem removeItem = new ToolStripMenuItem("移除選中照片");
            removeItem.Click += (s, e) => BtnRemovePhoto_Click(s, e);
            menu.Items.Add(removeItem);

            // 分隔線
            menu.Items.Add(new ToolStripSeparator());

            // 排序子菜單
            ToolStripMenuItem sortMenu = new ToolStripMenuItem("排序方式");

            ToolStripMenuItem sortByDateAsc = new ToolStripMenuItem("按日期升序");
            sortByDateAsc.Click += (s, e) =>
            {
                PhotoManager.Instance.SortPhotosByDate(true);
                UpdatePhotoListView();
            };
            sortMenu.DropDownItems.Add(sortByDateAsc);

            ToolStripMenuItem sortByDateDesc = new ToolStripMenuItem("按日期降序");
            sortByDateDesc.Click += (s, e) =>
            {
                PhotoManager.Instance.SortPhotosByDate(false);
                UpdatePhotoListView();
            };
            sortMenu.DropDownItems.Add(sortByDateDesc);

            ToolStripMenuItem sortByNameAsc = new ToolStripMenuItem("按文件名升序");
            sortByNameAsc.Click += (s, e) =>
            {
                PhotoManager.Instance.SortPhotosByFileName(true);
                UpdatePhotoListView();
            };
            sortMenu.DropDownItems.Add(sortByNameAsc);

            ToolStripMenuItem sortByNameDesc = new ToolStripMenuItem("按文件名降序");
            sortByNameDesc.Click += (s, e) =>
            {
                PhotoManager.Instance.SortPhotosByFileName(false);
                UpdatePhotoListView();
            };
            sortMenu.DropDownItems.Add(sortByNameDesc);

            menu.Items.Add(sortMenu);

            // 批量描述
            ToolStripMenuItem batchDescItem = new ToolStripMenuItem("批量設定描述");
            batchDescItem.Click += (s, e) => ShowBatchDescriptionDialog();
            menu.Items.Add(batchDescItem);

            // 顯示EXIF資訊
            ToolStripMenuItem showExifItem = new ToolStripMenuItem("顯示詳細資訊");
            showExifItem.Click += (s, e) => ShowPhotoDetailDialog();
            menu.Items.Add(showExifItem);

            // 匯出照片清單
            ToolStripMenuItem exportItem = new ToolStripMenuItem("匯出照片清單");
            exportItem.Click += (s, e) => ExportPhotoList();
            menu.Items.Add(exportItem);

            // 開啟菜單前檢查項目啟用狀態
            menu.Opening += (s, e) =>
            {
                bool hasPhotos = PhotoManager.Instance.Count > 0;
                bool hasSelection = lvPhotos.SelectedIndices.Count > 0;

                removeItem.Enabled = hasSelection;
                sortMenu.Enabled = hasPhotos && PhotoManager.Instance.Count > 1;
                batchDescItem.Enabled = hasPhotos;
                showExifItem.Enabled = hasSelection;
                exportItem.Enabled = hasPhotos;
            };

            return menu;
        }

        /// <summary>
        /// 虛擬項目檢索事件
        /// </summary>
        private void LvPhotos_RetrieveVirtualItem(object sender, RetrieveVirtualItemEventArgs e)
        {
            var photos = PhotoManager.Instance.GetAllPhotos();
            if (e.ItemIndex < photos.Count)
            {
                var photo = photos[e.ItemIndex];
                var item = new ListViewItem((e.ItemIndex + 1).ToString()); // 序號列

                // 添加子項
                item.SubItems.Add(Path.GetFileName(photo.FilePath)); // 檔案名稱
                if (photo.CaptureTime.HasValue)
                {
                    item.SubItems.Add(photo.CaptureTime.Value.ToString("yyyy/MM/dd HH:mm:ss")); // 實際日期
                }
                else
                {
                    item.SubItems.Add("未知");
                }

                e.Item = item;
            }
        }

        /// <summary>
        /// 照片列表項目拖動開始事件
        /// </summary>
        private void LvPhotos_ItemDrag(object sender, ItemDragEventArgs e)
        {
            // 只允許拖動列表項目
            if (e.Item is ListViewItem)
            {
                DoDragDrop(e.Item, DragDropEffects.Move);
                Logger.Log("開始拖動照片項目", Logger.LogLevel.Debug);
            }
        }

        /// <summary>
        /// 照片列表拖放進入事件
        /// </summary>
        private void LvPhotos_DragEnter(object sender, DragEventArgs e)
        {
            // 檢查拖放的數據類型
            if (e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                e.Effect = DragDropEffects.Move;
            }
            else if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // 也接受文件拖放以添加照片
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                bool hasValidImages = files.Any(file => IsImageFile(file));

                if (hasValidImages)
                {
                    e.Effect = DragDropEffects.Copy;
                }
                else
                {
                    e.Effect = DragDropEffects.None;
                }
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        /// <summary>
        /// 照片列表拖放事件
        /// </summary>
        private void LvPhotos_DragDrop(object sender, DragEventArgs e)
        {
            // 處理項目重新排序
            if (e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                // 獲取拖動項的索引
                int dragIndex = lvPhotos.SelectedIndices[0];

                // 獲取拖放位置的項
                Point targetPoint = lvPhotos.PointToClient(new Point(e.X, e.Y));
                ListViewItem targetItem = lvPhotos.GetItemAt(targetPoint.X, targetPoint.Y);
                int targetIndex = targetItem != null ? targetItem.Index : lvPhotos.Items.Count - 1;

                // 調用 PhotoManager 重新排序照片
                if (PhotoManager.Instance.ReorderPhotos(dragIndex, targetIndex))
                {
                    // 更新列表視圖
                    UpdatePhotoListView();

                    // 選中移動後的項目
                    lvPhotos.Items[targetIndex].Selected = true;
                    lvPhotos.Items[targetIndex].Focused = true;

                    Logger.Log($"已將照片從位置 {dragIndex} 移動到位置 {targetIndex}", Logger.LogLevel.Info);
                }
            }
            // 處理文件拖放以添加照片
            else if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                // 過濾出圖片文件
                var imageFiles = files.Where(file => IsImageFile(file)).ToList();

                if (imageFiles.Count > 0)
                {
                    // 添加照片
                    int addedCount = PhotoManager.Instance.AddPhotos(imageFiles);

                    if (addedCount > 0)
                    {
                        UpdatePhotoListView();
                        MessageBox.Show($"已成功添加 {addedCount} 張照片", "添加成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Logger.Log($"通過拖放添加了 {addedCount} 張照片", Logger.LogLevel.Info);
                    }
                    else
                    {
                        MessageBox.Show("無法添加照片，請檢查文件格式或照片數量限制", "添加失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        /// <summary>
        /// 判斷文件是否為圖片
        /// </summary>
        private bool IsImageFile(string filePath)
        {
            // 檢查文件擴展名
            string ext = Path.GetExtension(filePath).ToLower();
            string[] validExtensions = { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif" };

            return validExtensions.Contains(ext);
        }

        /// <summary>
        /// 照片管理器變更事件處理
        /// </summary>
        private void PhotoManager_PhotosChanged(object sender, PhotoManager.PhotoCollectionChangedEventArgs e)
        {
            // 確保在UI線程中執行
            if (InvokeRequired)
            {
                Invoke(new Action(() => PhotoManager_PhotosChanged(sender, e)));
                return;
            }

            // 根據變更類型更新UI
            switch (e.Type)
            {
                case PhotoManager.PhotoCollectionChangedEventArgs.ChangeType.Add:
                case PhotoManager.PhotoCollectionChangedEventArgs.ChangeType.Remove:
                case PhotoManager.PhotoCollectionChangedEventArgs.ChangeType.Reorder:
                    // 更新列表視圖
                    UpdatePhotoListView();
                    break;
                case PhotoManager.PhotoCollectionChangedEventArgs.ChangeType.Clear:
                    // 清空列表視圖
                    lvPhotos.VirtualListSize = 0;
                    lvPhotos.Refresh();
                    // 清除預覽
                    if (pbPhotoPreview.Image != null)
                    {
                        pbPhotoPreview.Image.Dispose();
                        pbPhotoPreview.Image = null;
                    }
                    txtPhotoDescription.Text = string.Empty;
                    txtPhotoDescription.Enabled = false;
                    break;
            }
        }

        /// <summary>
        /// 更新照片列表視圖
        /// </summary>
        private void UpdatePhotoListView()
        {
            // 更新虛擬列表大小
            lvPhotos.VirtualListSize = PhotoManager.Instance.Count;
            lvPhotos.Refresh();

            // 更新狀態欄顯示照片數量
            UpdateStatusBar($"共有 {PhotoManager.Instance.Count} 張照片");
        }

        /// <summary>
        /// 照片描述變更事件
        /// </summary>
        private void TxtPhotoDescription_TextChanged(object sender, EventArgs e)
        {
            // 如果有選取的照片，更新其描述
            if (lvPhotos.SelectedIndices.Count > 0)
            {
                int index = lvPhotos.SelectedIndices[0];
                if (index >= 0 && index < PhotoManager.Instance.Count)
                {
                    PhotoManager.Instance.UpdatePhotoDescription(index, txtPhotoDescription.Text);
                }
            }
        }

        /// <summary>
        /// 照片列表選擇變更事件
        /// </summary>
        private void LvPhotos_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 當清單選擇改變時更新右側顯示和描述文字框
            if (lvPhotos.SelectedIndices.Count > 0)
            {
                int index = lvPhotos.SelectedIndices[0];
                if (index >= 0 && index < PhotoManager.Instance.Count)
                {
                    var photo = PhotoManager.Instance.GetPhoto(index);
                    if (photo != null)
                    {
                        // 啟用文字框和顯示照片描述
                        txtPhotoDescription.Enabled = true;
                        txtPhotoDescription.Text = photo.Description;

                        // 顯示照片預覽
                        try
                        {
                            if (File.Exists(photo.FilePath))
                            {
                                // 釋放之前的圖像
                                if (pbPhotoPreview.Image != null)
                                {
                                    pbPhotoPreview.Image.Dispose();
                                    pbPhotoPreview.Image = null;
                                }

                                // 使用新的GetThumbnail方法創建適合顯示的縮圖
                                int previewWidth = pbPhotoPreview.Width;
                                int previewHeight = pbPhotoPreview.Height;

                                // 獲取縮圖
                                Image thumbnail = photo.GetThumbnail(previewWidth, previewHeight);

                                if (thumbnail != null)
                                {
                                    pbPhotoPreview.Image = thumbnail;

                                    // 顯示尺寸信息
                                    string sizeInfo = $"尺寸: {photo.Width}x{photo.Height} 像素 | 大小: {photo.GetFormattedFileSize()}";
                                    lblPhotoInfo.Text = sizeInfo;
                                }
                                else
                                {
                                    // 如果縮圖創建失敗，使用傳統方法
                                    using (var stream = new FileStream(photo.FilePath, FileMode.Open, FileAccess.Read))
                                    {
                                        pbPhotoPreview.Image = Image.FromStream(stream);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            pbPhotoPreview.Image = null;
                            Logger.Log($"載入照片預覽時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
                            MessageBox.Show($"載入照片時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            else
            {
                // 如果沒有選擇，停用文字框和清除預覽
                txtPhotoDescription.Enabled = false;
                txtPhotoDescription.Text = string.Empty;
                lblPhotoInfo.Text = "";

                // 釋放之前的圖像
                if (pbPhotoPreview.Image != null)
                {
                    pbPhotoPreview.Image.Dispose();
                    pbPhotoPreview.Image = null;
                }
            }
        }

        /// <summary>
        /// 生成文件點擊事件
        /// </summary>
        private void BtnGenerate_Click(object sender, EventArgs e)
        {
            // 驗證表單並顯示錯誤
            if (!ValidateFormData(true))
                return;

            // 使用改進的SaveDocument方法
            if (!SaveDocument())
            {
                Logger.Log("生成文件失敗", Logger.LogLevel.Warning);
            }
        }

        /// <summary>
        /// 關於點擊事件
        /// </summary>
        private void MenuHelpAbout_Click(object sender, EventArgs e)
        {
            MessageBox.Show($"警察照片證據生成器 v{APP_VERSION}\n\n用於生成包含照片的證據文件。\n\n新竹市警察局刑大科偵隊",
                "關於", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// 驗證表單數據
        /// </summary>
        /// <param name="showErrors">是否顯示錯誤訊息</param>
        /// <returns>是否驗證通過</returns>
        private bool ValidateFormData(bool showErrors = false)
        {
            bool isValid = true;
            string errorMessage = "";

            // 檢查是否選擇了大單位和小單位
            if (cmbMainUnit.SelectedIndex < 0)
            {
                isValid = false;
                errorMessage += "• 請選擇機關\n";
                if (showErrors) errorProvider.SetError(cmbMainUnit, "請選擇機關");
            }

            if (cmbSubUnit.SelectedIndex < 0)
            {
                isValid = false;
                errorMessage += "• 請選擇單位\n";
                if (showErrors) errorProvider.SetError(cmbSubUnit, "請選擇單位");
            }

            if (string.IsNullOrWhiteSpace(txtCase.Text))
            {
                isValid = false;
                errorMessage += "• 請填寫案由欄位\n";
                if (showErrors) errorProvider.SetError(txtCase, "案由是必填欄位");
            }

            // 檢查是否有照片
            if (PhotoManager.Instance.Count == 0)
            {
                if (showErrors)
                {
                    DialogResult result = MessageBox.Show("尚未添加任何照片，確定要繼續嗎？", "確認",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.No)
                    {
                        return false;
                    }
                }
            }

            // 顯示驗證錯誤
            if (!isValid && showErrors)
            {
                MessageBox.Show("請修正以下問題:\n" + errorMessage, "驗證錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return isValid;
        }

        // 修改 SaveDocument 方法，收集表單數據時獲取完整單位名稱
        private bool SaveDocument()
        {
            try
            {
                // 收集表單數據
                string mainUnit = cmbMainUnit.SelectedItem as string ?? "";
                string subUnit = cmbSubUnit.SelectedItem as string ?? "";
                string unit = $"{mainUnit} {subUnit}".Trim();
                string caseDescription = txtCase.Text.Trim();
                string time = dtpDateTime.Text.Trim();
                string address = txtLocation.Text.Trim();
                string name = txtPhotographer.Text.Trim();

                // 儲存設定
                settings.LastUnit = $"{mainUnit} {subUnit}";
                settings.LastPhotographer = name;
                settings.Save();

                // 生成新檔案名稱
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string defaultFileName = $"警察證據照片_{timestamp}.docx";
                string initialDir = string.IsNullOrEmpty(settings.LastSaveDirectory)
                    ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    : settings.LastSaveDirectory;

                // 使用SaveFileDialog讓用戶選擇儲存位置
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Word文件 (*.docx)|*.docx";
                    saveFileDialog.Title = "儲存證據文件";
                    saveFileDialog.FileName = defaultFileName;
                    saveFileDialog.InitialDirectory = initialDir;

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string saveFilePath = saveFileDialog.FileName;

                        // 更新最後儲存目錄
                        settings.LastSaveDirectory = Path.GetDirectoryName(saveFilePath);
                        settings.Save();

                        // 查找有效的模板路徑（使用應用程序目錄）
                        string appPath = Application.StartupPath;
                        string templatePath = Path.Combine(appPath, "template.docx");

                        // 檢查範本是否存在
                        if (!File.Exists(templatePath))
                        {
                            Logger.Log("找不到範本檔案", Logger.LogLevel.Error);
                            MessageBox.Show($"找不到範本檔案！\n\n請確認應用程式目錄下是否有template.docx檔案。",
                                "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }

                        // 顯示進度表單
                        progressForm = new ProgressForm();
                        progressForm.Show(this);

                        // 獲取照片列表
                        var photos = PhotoManager.Instance.GetAllPhotos();

                        // 使用 WordInteropHelper 生成文檔（非同步）
                        Task.Run(async () =>
                        {
                            bool success = await WordInteropHelper.GenerateDocumentAsync(
                                templatePath,
                                saveFilePath,
                                unit,
                                caseDescription,
                                time,
                                address,
                                name,
                                photos,
                                ProgressReportCallback);

                            // 回到UI線程處理結果
                            this.Invoke(new Action(() =>
                            {
                                // 關閉進度表單
                                progressForm.Close();
                                progressForm = null;

                                if (success)
                                {
                                    // 保存設定
                                    settings.Save();

                                    UpdateStatusBar("文件已成功生成");
                                    MessageBox.Show($"文件已成功生成！\n儲存路徑: {saveFilePath}", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    Logger.Log($"文件已成功生成: {saveFilePath}");

                                    // 詢問是否開啟已儲存的文件
                                    if (MessageBox.Show("是否立即開啟文件？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                    {
                                        try
                                        {
                                            System.Diagnostics.Process.Start(saveFilePath);
                                        }
                                        catch (Exception ex)
                                        {
                                            Logger.Log($"無法開啟已生成的文件: {ex.Message}", Logger.LogLevel.Error);
                                            MessageBox.Show($"無法開啟文件: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                    }
                                }
                                else
                                {
                                    UpdateStatusBar("生成文件失敗");
                                }
                            }));
                        });

                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"生成文件時發生錯誤: {ex.Message}\n{ex.StackTrace}", Logger.LogLevel.Error);
                MessageBox.Show($"生成文件時發生錯誤: {ex.Message}\n\n錯誤類型: {ex.GetType().FullName}",
                    "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatusBar("生成文件時發生錯誤");
            }

            return false;
        }

        /// <summary>
        /// 進度報告回調
        /// </summary>
        private void ProgressReportCallback(int progress, string message)
        {
            if (progressForm != null && !progressForm.IsDisposed)
            {
                this.Invoke(new Action(() =>
                {
                    progressForm.UpdateProgress(progress, message);
                }));
            }
        }

        /// <summary>
        /// 添加照片按鈕點擊事件
        /// </summary>
        private void BtnAddPhoto_Click(object sender, EventArgs e)
        {
            if (PhotoManager.Instance.Count >= PhotoManager.MAX_PHOTOS)
            {
                MessageBox.Show($"最多只能添加{PhotoManager.MAX_PHOTOS}張照片！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "圖片檔案|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.tif|所有檔案|*.*";
                openFileDialog.Title = "選擇照片";
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    int addedCount = 0;
                    int failedCount = 0;

                    foreach (string file in openFileDialog.FileNames)
                    {
                        if (PhotoManager.Instance.Count >= PhotoManager.MAX_PHOTOS) break;

                        // 使用PhotoManager添加照片
                        if (PhotoManager.Instance.AddPhoto(file))
                        {
                            addedCount++;
                        }
                        else
                        {
                            failedCount++;
                        }
                    }

                    // 提供添加結果反饋
                    if (addedCount > 0)
                    {
                        string message = $"已添加 {addedCount} 張照片，目前共有 {PhotoManager.Instance.Count} 張照片";
                        if (failedCount > 0)
                        {
                            message += $"\n有 {failedCount} 張照片無法添加";
                        }
                        MessageBox.Show(message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Logger.Log(message);
                    }
                    else if (failedCount > 0)
                    {
                        MessageBox.Show($"所選的 {failedCount} 張照片均無法添加", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        /// <summary>
        /// 移除照片按鈕點擊事件
        /// </summary>
        private void BtnRemovePhoto_Click(object sender, EventArgs e)
        {
            if (lvPhotos.SelectedIndices.Count > 0)
            {
                int index = lvPhotos.SelectedIndices[0];

                // 使用PhotoManager移除照片
                if (PhotoManager.Instance.RemovePhoto(index))
                {
                    MessageBox.Show("照片已移除", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Logger.Log($"已移除索引為 {index} 的照片");
                }
            }
            else
            {
                MessageBox.Show("請先選擇要移除的照片", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// 顯示批量設定描述對話框
        /// </summary>
        private void ShowBatchDescriptionDialog()
        {
            if (PhotoManager.Instance.Count == 0)
            {
                MessageBox.Show("沒有照片可以設定描述", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 創建一個簡單的輸入對話框
            using (var form = new Form())
            {
                form.Text = "批量設定照片描述";
                form.Size = new Size(500, 240);
                form.StartPosition = FormStartPosition.CenterParent;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                var lblInfo = new Label
                {
                    Text = "您可以使用以下標記創建描述模板：",
                    Location = new Point(10, 15),
                    Size = new Size(460, 20)
                };

                var lblTags = new Label
                {
                    Text = "{INDEX} - 照片序號\n{DATE} - 拍攝日期\n{TIME} - 拍攝時間\n{FILENAME} - 檔案名稱",
                    Location = new Point(20, 35),
                    Size = new Size(460, 60)
                };

                var lblTemplate = new Label
                {
                    Text = "描述模板：",
                    Location = new Point(10, 95),
                    Size = new Size(100, 20)
                };

                var textBox = new TextBox
                {
                    Text = "照片 {INDEX} - 拍攝於 {DATE}",
                    Location = new Point(20, 115),
                    Size = new Size(450, 20),
                    Multiline = true,
                    Height = 40
                };

                var okButton = new Button
                {
                    Text = "確定",
                    DialogResult = DialogResult.OK,
                    Location = new Point(300, 170),
                    Width = 80
                };

                var cancelButton = new Button
                {
                    Text = "取消",
                    DialogResult = DialogResult.Cancel,
                    Location = new Point(390, 170),
                    Width = 80
                };

                form.Controls.Add(lblInfo);
                form.Controls.Add(lblTags);
                form.Controls.Add(lblTemplate);
                form.Controls.Add(textBox);
                form.Controls.Add(okButton);
                form.Controls.Add(cancelButton);

                form.AcceptButton = okButton;
                form.CancelButton = cancelButton;

                // 顯示對話框
                if (form.ShowDialog() == DialogResult.OK)
                {
                    // 批量更新描述
                    if (PhotoManager.Instance.BatchUpdateDescriptions(textBox.Text))
                    {
                        MessageBox.Show("已更新所有照片描述", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdatePhotoListView();
                    }
                    else
                    {
                        MessageBox.Show("更新照片描述失敗", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        /// <summary>
        /// 顯示照片詳細資訊對話框
        /// </summary>
        private void ShowPhotoDetailDialog()
        {
            if (lvPhotos.SelectedIndices.Count == 0)
            {
                MessageBox.Show("請先選擇一張照片", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int index = lvPhotos.SelectedIndices[0];
            var photo = PhotoManager.Instance.GetPhoto(index);

            if (photo == null)
                return;

            // 創建詳細資訊對話框
            using (var form = new Form())
            {
                form.Text = $"照片詳細資訊 - {Path.GetFileName(photo.FilePath)}";
                form.Size = new Size(600, 500);
                form.StartPosition = FormStartPosition.CenterParent;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                // 照片預覽
                var preview = new PictureBox
                {
                    Location = new Point(10, 10),
                    Size = new Size(200, 200),
                    SizeMode = PictureBoxSizeMode.Zoom,
                    BorderStyle = BorderStyle.FixedSingle
                };

                try
                {
                    preview.Image = photo.GetThumbnail(200, 200);
                }
                catch
                {
                    // 忽略預覽錯誤
                }

                // 基本資訊
                var lblBasicInfo = new Label
                {
                    Text = "基本資訊：",
                    Location = new Point(220, 10),
                    Size = new Size(100, 20),
                    Font = new Font(this.Font, FontStyle.Bold)
                };

                var txtBasicInfo = new TextBox
                {
                    Location = new Point(220, 30),
                    Size = new Size(350, 180),
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Vertical,
                    Text = $"文件名稱: {Path.GetFileName(photo.FilePath)}\r\n" +
                           $"文件路徑: {photo.FilePath}\r\n" +
                           $"拍攝時間: {photo.GetFormattedTime()}\r\n" +
                           $"圖片尺寸: {photo.Width}x{photo.Height} 像素\r\n" +
                           $"檔案大小: {photo.GetFormattedFileSize()}\r\n" +
                           $"描述: {photo.Description}"
                };

                // EXIF資訊
                var lblExif = new Label
                {
                    Text = "EXIF資訊：",
                    Location = new Point(10, 220),
                    Size = new Size(100, 20),
                    Font = new Font(this.Font, FontStyle.Bold)
                };

                var txtExif = new TextBox
                {
                    Location = new Point(10, 240),
                    Size = new Size(560, 170),
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Both,
                    WordWrap = false
                };

                // 獲取EXIF數據
                var exifData = photo.GetExifData();
                StringBuilder sb = new StringBuilder();

                if (exifData.Count > 0)
                {
                    foreach (var pair in exifData)
                    {
                        sb.AppendLine($"{pair.Key}: {pair.Value}");
                    }
                    txtExif.Text = sb.ToString();
                }
                else
                {
                    txtExif.Text = "沒有發現EXIF資訊";
                }

                // 關閉按鈕
                var closeButton = new Button
                {
                    Text = "關閉",
                    DialogResult = DialogResult.Cancel,
                    Location = new Point(490, 420),
                    Size = new Size(80, 30)
                };

                form.Controls.Add(preview);
                form.Controls.Add(lblBasicInfo);
                form.Controls.Add(txtBasicInfo);
                form.Controls.Add(lblExif);
                form.Controls.Add(txtExif);
                form.Controls.Add(closeButton);

                form.AcceptButton = closeButton;
                form.CancelButton = closeButton;

                form.ShowDialog();

                // 釋放資源
                preview.Image?.Dispose();
            }
        }

        /// <summary>
        /// 更新狀態列
        /// </summary>
        private void UpdateStatusBar(string message)
        {
            // 記錄狀態到日誌
            Logger.Log($"狀態: {message}", Logger.LogLevel.Debug);
            statusLabel.Text = message;
        }

        /// <summary>
        /// 匯出照片清單
        /// </summary>
        private void ExportPhotoList()
        {
            if (PhotoManager.Instance.Count == 0)
            {
                MessageBox.Show("沒有照片可供匯出", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Filter = "文字檔案 (*.txt)|*.txt|所有檔案 (*.*)|*.*";
                dlg.Title = "匯出照片清單";
                dlg.FileName = $"照片清單_{DateTime.Now:yyyyMMdd}.txt";

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    // 詢問是否包含EXIF資訊
                    DialogResult result = MessageBox.Show("是否包含詳細的EXIF資訊？", "確認",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    bool includeExif = (result == DialogResult.Yes);

                    // 匯出照片列表
                    if (PhotoManager.Instance.ExportPhotoList(dlg.FileName, includeExif))
                    {
                        MessageBox.Show($"照片清單已成功匯出到:\n{dlg.FileName}", "匯出成功",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // 詢問是否開啟檔案
                        result = MessageBox.Show("是否立即開啟匯出的檔案？", "確認",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (result == DialogResult.Yes)
                        {
                            try
                            {
                                System.Diagnostics.Process.Start(dlg.FileName);
                            }
                            catch (Exception ex)
                            {
                                Logger.Log($"開啟檔案時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
                                MessageBox.Show($"無法開啟檔案: {ex.Message}", "錯誤",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("匯出照片清單失敗", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        /// <summary>
        /// 表單關閉時的處理
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);

            // 儲存使用者設定
            if (settings != null)
            {
                settings.Save();
            }

            // 釋放資源
            if (PhotoManager.Instance != null && PhotoManager.Instance.Count > 0)
            {
                // 詢問是否保存更改
                if (MessageBox.Show("是否在退出前生成文件？", "確認",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    // 如果用戶選擇保存，則調用SaveDocument方法
                    if (!SaveDocument())
                    {
                        // 如果保存失敗，詢問是否仍要退出
                        if (MessageBox.Show("生成文件失敗，是否仍要退出？", "確認",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                        {
                            // 取消關閉操作
                            e.Cancel = true;
                            return;
                        }
                    }
                }
            }

            // 清理資源
            try
            {
                // 釋放圖片預覽資源
                if (pbPhotoPreview.Image != null)
                {
                    pbPhotoPreview.Image.Dispose();
                    pbPhotoPreview.Image = null;
                }

                // 其他資源釋放
                // ...
            }
            catch (Exception ex)
            {
                Logger.Log($"釋放資源時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
            }

            // 記錄應用程式結束
            Logger.Log("應用程式結束");
        }

        private void TblUnitLayout_Paint(object sender, PaintEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
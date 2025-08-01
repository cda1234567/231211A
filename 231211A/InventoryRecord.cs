using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace _231211A
{
    /// <summary>
    /// 庫存記錄管理類別
    /// </summary>
    public class InventoryRecord
    {
        public string PartNumber { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public double CurrentStock { get; set; }
        public double PreviousStock { get; set; }
        public double DeductedQuantity { get; set; }
        public DateTime LastUpdateTime { get; set; }
        public string UpdateSource { get; set; } = string.Empty;
        public InventoryStatus Status { get; set; }

        public enum InventoryStatus
        {
            Normal,      // 正常庫存
            LowStock,    // 低庫存
            ZeroStock,   // 零庫存
            Negative     // 負庫存
        }

        /// <summary>
        /// 根據庫存數量更新狀態
        /// </summary>
        public void UpdateStatus()
        {
            if (CurrentStock < 0)
                Status = InventoryStatus.Negative;
            else if (CurrentStock == 0)
                Status = InventoryStatus.ZeroStock;
            else if (CurrentStock < 10) // 低庫存閾值
                Status = InventoryStatus.LowStock;
            else
                Status = InventoryStatus.Normal;
        }

        /// <summary>
        /// 取得狀態顯示文字
        /// </summary>
        public string GetStatusText()
        {
            return Status switch
            {
                InventoryStatus.Normal => "正常",
                InventoryStatus.LowStock => "低庫存",
                InventoryStatus.ZeroStock => "零庫存",
                InventoryStatus.Negative => "負庫存",
                _ => "未知"
            };
        }

        /// <summary>
        /// 取得狀態顏色
        /// </summary>
        public System.Drawing.Color GetStatusColor()
        {
            return Status switch
            {
                InventoryStatus.Normal => System.Drawing.Color.Green,
                InventoryStatus.LowStock => System.Drawing.Color.Orange,
                InventoryStatus.ZeroStock => System.Drawing.Color.Gold,
                InventoryStatus.Negative => System.Drawing.Color.Red,
                _ => System.Drawing.Color.Black
            };
        }
    }

    /// <summary>
    /// 庫存記錄管理器
    /// </summary>
    public static class InventoryManager
    {
        private static readonly string INVENTORY_LOG_FILE = "inventory_log.csv";
        private static readonly List<InventoryRecord> _inventoryHistory = new();

        /// <summary>
        /// 記錄庫存變動
        /// </summary>
        public static void LogInventoryChange(string partNumber, string description, 
            double previousStock, double currentStock, string source)
        {
            var record = new InventoryRecord
            {
                PartNumber = partNumber,
                Description = description,
                PreviousStock = previousStock,
                CurrentStock = currentStock,
                DeductedQuantity = previousStock - currentStock,
                LastUpdateTime = DateTime.Now,
                UpdateSource = source
            };
            
            record.UpdateStatus();
            _inventoryHistory.Add(record);
            
            // 寫入CSV檔案
            WriteToLogFile(record);
        }

        /// <summary>
        /// 寫入記錄檔案
        /// </summary>
        private static void WriteToLogFile(InventoryRecord record)
        {
            try
            {
                bool fileExists = File.Exists(INVENTORY_LOG_FILE);
                using (var writer = new StreamWriter(INVENTORY_LOG_FILE, true, System.Text.Encoding.UTF8))
                {
                    // 如果檔案不存在，寫入標題列
                    if (!fileExists)
                    {
                        writer.WriteLine("時間,料號,描述,原庫存,現庫存,扣減數量,來源檔案,狀態");
                    }
                    
                    writer.WriteLine($"{record.LastUpdateTime:yyyy-MM-dd HH:mm:ss}," +
                                   $"{record.PartNumber}," +
                                   $"{record.Description}," +
                                   $"{record.PreviousStock}," +
                                   $"{record.CurrentStock}," +
                                   $"{record.DeductedQuantity}," +
                                   $"{record.UpdateSource}," +
                                   $"{record.GetStatusText()}");
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"寫入記錄檔案失敗：{ex.Message}", "錯誤");
            }
        }

        /// <summary>
        /// 取得庫存歷史記錄
        /// </summary>
        public static List<InventoryRecord> GetInventoryHistory()
        {
            return _inventoryHistory.ToList();
        }

        /// <summary>
        /// 清除歷史記錄
        /// </summary>
        public static void ClearHistory()
        {
            _inventoryHistory.Clear();
        }

        /// <summary>
        /// 從CSV檔案載入歷史記錄
        /// </summary>
        public static void LoadHistoryFromFile()
        {
            try
            {
                if (!File.Exists(INVENTORY_LOG_FILE))
                    return;

                var lines = File.ReadAllLines(INVENTORY_LOG_FILE, System.Text.Encoding.UTF8);
                
                // 跳過標題列
                for (int i = 1; i < lines.Length; i++)
                {
                    var parts = lines[i].Split(',');
                    if (parts.Length >= 8)
                    {
                        var record = new InventoryRecord
                        {
                            LastUpdateTime = DateTime.Parse(parts[0]),
                            PartNumber = parts[1],
                            Description = parts[2],
                            PreviousStock = double.Parse(parts[3]),
                            CurrentStock = double.Parse(parts[4]),
                            DeductedQuantity = double.Parse(parts[5]),
                            UpdateSource = parts[6]
                        };
                        
                        record.UpdateStatus();
                        _inventoryHistory.Add(record);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"載入記錄檔案失敗：{ex.Message}", "錯誤");
            }
        }

        /// <summary>
        /// 取得庫存統計
        /// </summary>
        public static InventoryStatistics GetStatistics(DataTable inventoryData)
        {
            var stats = new InventoryStatistics();
            
            if (inventoryData?.Rows.Count > 0)
            {
                stats.TotalItems = inventoryData.Rows.Count;
                
                foreach (DataRow row in inventoryData.Rows)
                {
                    if (row.ItemArray.Length > 0)
                    {
                        var lastValue = row.ItemArray[row.ItemArray.Length - 1];
                        if (double.TryParse(lastValue?.ToString(), out double stock))
                        {
                            if (stock < 0)
                                stats.NegativeStock++;
                            else if (stock == 0)
                                stats.ZeroStock++;
                            else if (stock < 10)
                                stats.LowStock++;
                            else
                                stats.NormalStock++;
                                
                            stats.TotalValue += stock;
                        }
                    }
                }
            }
            
            return stats;
        }
    }

    /// <summary>
    /// 庫存統計資訊
    /// </summary>
    public class InventoryStatistics
    {
        public int TotalItems { get; set; }
        public int NormalStock { get; set; }
        public int LowStock { get; set; }
        public int ZeroStock { get; set; }
        public int NegativeStock { get; set; }
        public double TotalValue { get; set; }

        public override string ToString()
        {
            return $"總計：{TotalItems} | 正常：{NormalStock} | 低庫存：{LowStock} | 零庫存：{ZeroStock} | 負庫存：{NegativeStock}";
        }
    }
}
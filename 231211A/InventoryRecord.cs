using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace _231211A
{
    /// <summary>
    /// �w�s�O���޲z���O
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
            Normal,      // ���`�w�s
            LowStock,    // �C�w�s
            ZeroStock,   // �s�w�s
            Negative     // �t�w�s
        }

        /// <summary>
        /// �ھڮw�s�ƶq��s���A
        /// </summary>
        public void UpdateStatus()
        {
            if (CurrentStock < 0)
                Status = InventoryStatus.Negative;
            else if (CurrentStock == 0)
                Status = InventoryStatus.ZeroStock;
            else if (CurrentStock < 10) // �C�w�s�H��
                Status = InventoryStatus.LowStock;
            else
                Status = InventoryStatus.Normal;
        }

        /// <summary>
        /// ���o���A��ܤ�r
        /// </summary>
        public string GetStatusText()
        {
            return Status switch
            {
                InventoryStatus.Normal => "���`",
                InventoryStatus.LowStock => "�C�w�s",
                InventoryStatus.ZeroStock => "�s�w�s",
                InventoryStatus.Negative => "�t�w�s",
                _ => "����"
            };
        }

        /// <summary>
        /// ���o���A�C��
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
    /// �w�s�O���޲z��
    /// </summary>
    public static class InventoryManager
    {
        private static readonly string INVENTORY_LOG_FILE = "inventory_log.csv";
        private static readonly List<InventoryRecord> _inventoryHistory = new();

        /// <summary>
        /// �O���w�s�ܰ�
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
            
            // �g�JCSV�ɮ�
            WriteToLogFile(record);
        }

        /// <summary>
        /// �g�J�O���ɮ�
        /// </summary>
        private static void WriteToLogFile(InventoryRecord record)
        {
            try
            {
                bool fileExists = File.Exists(INVENTORY_LOG_FILE);
                using (var writer = new StreamWriter(INVENTORY_LOG_FILE, true, System.Text.Encoding.UTF8))
                {
                    // �p�G�ɮפ��s�b�A�g�J���D�C
                    if (!fileExists)
                    {
                        writer.WriteLine("�ɶ�,�Ƹ�,�y�z,��w�s,�{�w�s,����ƶq,�ӷ��ɮ�,���A");
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
                System.Windows.Forms.MessageBox.Show($"�g�J�O���ɮץ��ѡG{ex.Message}", "���~");
            }
        }

        /// <summary>
        /// ���o�w�s���v�O��
        /// </summary>
        public static List<InventoryRecord> GetInventoryHistory()
        {
            return _inventoryHistory.ToList();
        }

        /// <summary>
        /// �M�����v�O��
        /// </summary>
        public static void ClearHistory()
        {
            _inventoryHistory.Clear();
        }

        /// <summary>
        /// �qCSV�ɮ׸��J���v�O��
        /// </summary>
        public static void LoadHistoryFromFile()
        {
            try
            {
                if (!File.Exists(INVENTORY_LOG_FILE))
                    return;

                var lines = File.ReadAllLines(INVENTORY_LOG_FILE, System.Text.Encoding.UTF8);
                
                // ���L���D�C
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
                System.Windows.Forms.MessageBox.Show($"���J�O���ɮץ��ѡG{ex.Message}", "���~");
            }
        }

        /// <summary>
        /// ���o�w�s�έp
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
    /// �w�s�έp��T
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
            return $"�`�p�G{TotalItems} | ���`�G{NormalStock} | �C�w�s�G{LowStock} | �s�w�s�G{ZeroStock} | �t�w�s�G{NegativeStock}";
        }
    }
}
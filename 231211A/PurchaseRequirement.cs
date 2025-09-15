using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace _231211A
{
    /// <summary>
    /// 缺料需求狀態
    /// </summary>
    public enum PurchaseRequirementStatus
    {
        Open,
        HasPO,
        Closed,
        Ignored
    }

    /// <summary>
    /// 缺料需求紀錄
    /// </summary>
    public class PurchaseRequirement
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public string PartNumber { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public double RequestQty { get; set; }
        public DateTime CreatedTime { get; set; } = DateTime.Now;
        public PurchaseRequirementStatus Status { get; set; } = PurchaseRequirementStatus.Open;
        public string Reason { get; set; } = "ShortByIssue"; // 可固定
        public string Note { get; set; } = string.Empty;
    }

    /// <summary>
    /// 缺料需求集中管理 (記憶體 + 匯出 CSV)
    /// </summary>
    public static class PurchaseRequirementManager
    {
        private static readonly List<PurchaseRequirement> _requirements = new();
        private const string FILE_NAME = "purchase_requirements.csv";

        public static IReadOnlyList<PurchaseRequirement> Requirements => _requirements;

        public static PurchaseRequirement AddRequirement(string part, string desc, double qty, PurchaseRequirementStatus status)
        {
            var existing = _requirements.FirstOrDefault(r => r.PartNumber.Equals(part, StringComparison.OrdinalIgnoreCase)
                                                             && r.Status == PurchaseRequirementStatus.Open);
            if (existing != null)
            {
                existing.RequestQty = Math.Max(existing.RequestQty, qty); // 保留較大需求
                return existing;
            }

            var req = new PurchaseRequirement
            {
                PartNumber = part,
                Description = desc,
                RequestQty = qty,
                Status = status
            };
            _requirements.Add(req);
            return req;
        }

        public static void MarkHasPO(string part)
        {
            var r = _requirements.FirstOrDefault(x => x.PartNumber.Equals(part, StringComparison.OrdinalIgnoreCase) && x.Status == PurchaseRequirementStatus.Open);
            if (r == null)
            {
                // 建一筆狀態即為 HasPO 以供追溯
                AddRequirement(part, string.Empty, 0, PurchaseRequirementStatus.HasPO);
            }
            else
            {
                r.Status = PurchaseRequirementStatus.HasPO;
            }
        }

        public static void IgnoreOnce(string part, double qty)
        {
            AddRequirement(part, string.Empty, qty, PurchaseRequirementStatus.Ignored);
        }

        public static void ExportCsv()
        {
            try
            {
                using var sw = new StreamWriter(FILE_NAME, false, System.Text.Encoding.UTF8);
                sw.WriteLine("Id,PartNumber,Description,RequestQty,Status,CreatedTime,Reason,Note");
                foreach (var r in _requirements.OrderBy(r => r.PartNumber))
                {
                    sw.WriteLine(string.Join(',', new[]
                    {
                        r.Id.ToString(),
                        r.PartNumber,
                        r.Description.Replace(',', ' '),
                        r.RequestQty.ToString(),
                        r.Status.ToString(),
                        r.CreatedTime.ToString("yyyy-MM-dd HH:mm:ss"),
                        r.Reason,
                        r.Note.Replace(',', ' ')
                    }));
                }
            }
            catch { }
        }
    }
}

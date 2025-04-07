using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class SystemChange
    {
        public int SysChgId { get; set; }
        public string? SysDbver { get; set; }
        public string? SysAppVer { get; set; }
        public string? SysDbverPrev { get; set; }
        public string? SysAppVerPrev { get; set; }
        public string? SysTableName { get; set; }
        public string? SysColName { get; set; }
        public string? SysActivity { get; set; }
        public string? SysFieldType { get; set; }
        public string? Notes { get; set; }
        public DateTime? DateCreate { get; set; }
    }
}

using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Loft
    {
        public string? Code { get; set; }
        public string? Name { get; set; }
        public string? Ownername { get; set; }
        public string? Owneradd1 { get; set; }
        public string? Owneradd2 { get; set; }
        public string? Ownercity { get; set; }
        public string? Ownerpc { get; set; }
        public string? Ownerstate { get; set; }
        public string? Ownerph { get; set; }
        public string? License { get; set; }
        public int? LatDeg { get; set; }
        public int? LatMin { get; set; }
        public double? LatSec { get; set; }
        public int? LngDeg { get; set; }
        public int? LngMin { get; set; }
        public double? LngSec { get; set; }
        public string? Ownermark { get; set; }
        public int? Bscmonth { get; set; }
        public string? OwnerClub { get; set; }
        public string? Checksum { get; set; }
        public string? PhSupport { get; set; }
        public int? RaceCode { get; set; }
        public int? Rslpp { get; set; }
        public int? RsraceFormType { get; set; }
        public string? Dbversion { get; set; }
        public DateTime? LastBackup { get; set; }
        public string? Copyright { get; set; }
        public string? DistanceType { get; set; }
        public string? LastBackupName { get; set; }
        public DateTime? LastBackupDateTime { get; set; }
        public string? OwnerFed { get; set; }
        public int Id { get; set; }
    }
}

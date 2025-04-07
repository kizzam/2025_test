using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Master
    {
        public int Yr { get; set; }
        public string Mark { get; set; } = null!;
        public string Ringno { get; set; } = null!;
        public string? Birdname { get; set; }
        public string? Bredby { get; set; }
        public string? Bredfor { get; set; }
        public string? Bloodline { get; set; }
        public string? Dob { get; set; }
        public string? Codesex { get; set; }
        public string? Codecolour { get; set; }
        public string? Codeeye { get; set; }
        public string? Codecage { get; set; }
        public string? Racelast { get; set; }
        public DateTime? RaceLastDte { get; set; }
        public DateTime? BredLastDte { get; set; }
        public bool? Indracelast { get; set; }
        public bool? Indstock { get; set; }
        public bool? Indrace { get; set; }
        public bool? Indsoh { get; set; }
        public bool? Indref { get; set; }
        public bool? Indbpa { get; set; }
        public bool? Indbpb { get; set; }
        public bool? Indsp1 { get; set; }
        public bool? Indsp2 { get; set; }
        public bool? Indstatus { get; set; }
        public bool? Indbredforstock { get; set; }
        public string? Status { get; set; }
        public string? Notes { get; set; }
        public int? Sireyr { get; set; }
        public string? Siremark { get; set; }
        public string? Sireringno { get; set; }
        public int? Damyr { get; set; }
        public string? Dammark { get; set; }
        public string? Damringno { get; set; }
        public string? Notespedg { get; set; }
        public string? Codes { get; set; }
        public int? DobLi { get; set; }
        public DateTime? AqDate { get; set; }
        public decimal? AqValue { get; set; }
        public string? AqStatus { get; set; }
        public DateTime? DisDate { get; set; }
        public decimal? DisValue { get; set; }
        public DateTime? DDob { get; set; }
        public string? DisStatus { get; set; }
        public string? Codebline { get; set; }
        public int? Rating { get; set; }
        public int? RatingPerf { get; set; }
        public int? RatingBloodLine { get; set; }
        public int? RatingCustom { get; set; }
        public int? RatingBreeder { get; set; }
        public bool? IndDnacert { get; set; }
        public bool? IndBredSnglPen { get; set; }
        public bool? IndPedgOnly { get; set; }
        public string? IndBf { get; set; }
        public string? IndBfcol { get; set; }
        public string? IndTf { get; set; }
        public string? IndThroat { get; set; }
        public string? IndBredSglPen { get; set; }
        public string? IndImport { get; set; }
        public double? Seq { get; set; }
        public double? Cat { get; set; }
        public int? LotNo { get; set; }
        public int? SeqRef { get; set; }
        public int? SeqSale { get; set; }
    }
}

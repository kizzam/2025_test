using Microsoft.VisualBasic;

namespace SalesMgmntApp.Entities
{
    public class Bird
    {
        public int BirdId { get; set; }

        public int RingYr { get; set; } = 2023;
        public string? RingMark { get; set; } = string.Empty;
        public string? RingNo { get; set; } = string.Empty;

        public string? BirdName { get; set; } = string.Empty;
        public DateTime DOB { get; set; } = DateTime.MinValue;
        public string Gender { get; set; } = "U";
        public string? Breeder { get; set; } = "";
        public string? BredFor { get; set; } = "";
        public int BloodLineId { get; set; } = 0;

        public int ColourId { get; set; } = 0;
        public string Eye_Code { get; set; } = "NA";

        public Boolean IsSOH { get; set; } = true;
        public Boolean IsRace { get; set; } = true;
        public Boolean IsStock { get; set; } = false;
        public Boolean IsOLR { get; set; } = false;
        public Boolean IsBP { get; set; } = false;
        public Boolean IsRef { get; set; } = false;
        public Boolean IsBred4Stock { get; set; } = false;
        public Boolean IsDNACert { get; set; } = false;
        public Boolean IsSnglPen { get; set; } = false;

        public String ImagePath { get; set; } = string.Empty;

        public int SireRingYr { get; set; } = 0;
        public string SireRingMark { get; set; } = string.Empty;
        public string SireRingNo { get; set; } = string.Empty;

        public int DamRingYr { get; set; } = 0;
        public string DamRingMark { get; set; } = string.Empty;
        public string DamRingNo { get; set; } = string.Empty;

        public int LoftId { get; set; } = 0;

        public string PedgNotes { get; set; } = string.Empty;

        public DateTime DT_Modif { get; set; } = DateTime.Now;
        public DateTime DT_Create { get; set; } = DateTime.Now;
    }
}

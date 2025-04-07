namespace SalesMgmntApp.Models
{
    public class Performance
    {
        public int PerfId { get; set; }
        public int BirdId { get; set; }
        public int RingYr { get; set; } = 0;
        public string RingNo { get; set; } = string.Empty;
        public string RingMark { get; set; } = string.Empty;
        public DateTime PerfDate { get; set; } = DateTime.Now;
        public string RacePoint { get; set; } = string.Empty;
        public string RaceType { get; set; } = string.Empty;
        public string RaceName { get; set; } = string.Empty;
        public int RaceDistance { get; set; } = 0;
        public string PositionInternat { get; set; } = string.Empty;
        public string PositionNat { get; set; } = string.Empty;
        public string PositionFed { get; set; } = string.Empty;
        public string PositionClub { get; set; } = string.Empty;
        public string PositionPools { get; set; } = string.Empty;
        public string PositionNotes { get; set; } = string.Empty;
        public int NobirdsInterNAT { get; set; } = 0;
        public int NobirdsNAT { get; set; } = 0;
        public int NobirdsFED { get; set; } = 0;
        public int NobirdsClub { get; set; } = 0;
        public string ImagePath { get; set; } = string.Empty;
        public int status { get; set; } = 0;
        public DateTime DT_Modif { get; set; } = DateTime.Now;
        public DateTime DT_Create { get; set; } = DateTime.Now;
    }
}

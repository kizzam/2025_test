using System.Runtime;

namespace SalesMgmntApp.Entities
{
    public class Loft
    {
        public int LoftId { get; set; }
        public string LoftName { get; set; } = string.Empty;
        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; } = string.Empty;
        public string LoftAddr1 { get; set; } = string.Empty;
        public string LoftAddr2 { get; set; } = string.Empty;
        public string LoftAddr3 { get; set; } = string.Empty;
        public string City { get; set; } = string.Empty;
        public string State { get; set; } = string.Empty;
        public string PostCode { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;
        public string Phone1 { get; set; } = string.Empty;
        public string Phone2 { get; set; } = string.Empty;
        public string eMAIL { get; set; } = string.Empty;
        public string LoftMark { get; set; } = string.Empty;
        public string LoftClub { get; set; } = string.Empty;
        public string LoftFed { get; set; } = string.Empty;
        public int BscMonth { get; set; } = 5;
        public string ETS_Brand { get; set; } = string.Empty;
        public string ETS_Model { get; set; } = string.Empty;
        public string Latitude { get; set; } = string.Empty;
        public string Longitude { get; set; } = string.Empty;
        public float LatDeg { get; set; } = float.MinValue;
        public float LngDeg { get; set; } = float.MinValue;
        public string LicCode { get; set; } = string.Empty;
        public string SerialNo { get; set; } = string.Empty;
        public string dBVersion { get; set; } = "0.0";
        public DateTime LastBackupDT { get; set; } = DateTime.Parse("01 Jan 1900");
        public DateTime LastRunDT { get; set; } = DateTime.Parse("01 Jan 1900");
        public DateTime DT_Modif { get; set; } = DateTime.Now;
        public DateTime DT_Create { get; set; } = DateTime.Now;
    }
}

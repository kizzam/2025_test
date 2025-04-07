namespace SalesMgmntApp.Entities
{
    public class BloodLine
    {
        public int BloodLineId { get; set; }
        public string BloodLineName { get; set; } = string.Empty;
        public int Status { get; set; } = 0;
        public DateTime DT_Modif { get; set; } = DateTime.Now;
        public DateTime DT_Create { get; set; } = DateTime.Now;
    }
}

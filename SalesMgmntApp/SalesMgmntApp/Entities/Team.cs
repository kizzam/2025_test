namespace SalesMgmntApp.Entities
{
    public class Team
    {
        public int TeamId { get; set; }
        public string TeamCode { get; set; } = string.Empty;
        public string TeamName { get; set; } = string.Empty;
        public string TeamDescription { get; set; } = string.Empty;
        public int TeamStatus { get; set; } = 0;
        public DateTime DT_Modif { get; set; } = DateTime.Now;
        public DateTime DT_Create { get; set; } = DateTime.Now;
    }
}

namespace SalesMgmntApp.Entities
{
    public class TeamBird
    {
        public int TeamBirdId { get; set; }
        public int BirdId { get; set; }
        public int TeamId { get; set; }
        public int Status { get; set; }
        public DateTime DT_Modif { get; set; } = DateTime.Now;
        public DateTime DT_Create { get; set; } = DateTime.Now;
    }
}

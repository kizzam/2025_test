namespace SalesMgmntApp.Entities
{
    public class Colour
    {
        public int ColourId { get; set; }
        public string ColourCode { get; set; } = string.Empty;
        public string ColourFull { get; set; } = string.Empty;
        public string ColourDescription { get; set; } = string.Empty;
        public int Status { get; set; } = 0;
        public DateTime DT_Modif { get; set; } = DateTime.Now;
        public DateTime DT_Create { get; set; } = DateTime.Now;
    }
}

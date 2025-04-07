namespace SalesMgmntApp.Entities
{
    public class Catalogue
    {
        public int CatalogueID { get; set; }
        public string CatalogueName { get; set; } = string.Empty;
        public string CatalogueDescription { get; set; } = string.Empty;
        public int TypeId { get; set; } = 0;
        public int Status { get; set; } = 0;
        public DateTime DT_Modif { get; set; } = DateTime.Now;
        public DateTime DT_Create { get; set; } = DateTime.Now;
    }
}

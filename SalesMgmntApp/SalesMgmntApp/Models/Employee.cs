namespace SalesMgmntApp.Models
{
    public class Employee
    {
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string JobTitle { get; set; }
            public string Gender { get; set; }
            public DateTime DateOfBirth { get; set; }
            public int? ReportToEmpId { get; set; }
    }
}

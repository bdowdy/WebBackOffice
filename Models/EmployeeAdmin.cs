namespace MSaleWebServer.Models
{
    public class EmployeeAdminModel
    {
        public string OriginalFirstName { get; set; }
        public string OriginalLastName { get; set; }

        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string LastName { get; set; }
        public string NickName { get; set; }

        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Zip { get; set; }

        public string AccessLevel { get; set; }
        public string AccessNumber { get; set; }

        public string HomePhone { get; set; }
        public string Phone2 { get; set; }

        public string Language { get; set; }
        public string EmployeeId { get; set; }
        public string CompanyId { get; set; }

        public string HireDate { get; set; }
        public string Dob { get; set; }

        public bool Terminated { get; set; }
    }
}

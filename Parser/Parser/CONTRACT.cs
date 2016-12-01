namespace Parser
{
    public class CONTRACT : Serializer
    {
        public string Status = "A";
        public string DealerCode = "A";
        public string DealerPointCode = "A";
        public string DealerContractCode = "A";
        public string DealerContractDate = "A";
        public string ABSContractCode = "A";
        public Customer CUSTOMER;
        private string BANKPROPLIST = "";

        public CONTRACT()
        {
            CUSTOMER = new Customer();
        }
    }

    public class Customer
    {
        public string CUSTOMERTYPESId = "A";
        public string SPHERESId = "A";
        public string Resident = "A";
        public string Ratepayer = "A";
        public Person PERSON;
        public Address ADDRESS;

        public Customer()
        {
            PERSON = new Person();
            ADDRESS = new Address();
        }
    }

    public class Person
    {
        public string PERSONTYPESId = "A";
        public PersonName PERSONNAME;
        public Document DOCUMENT;
        public string INN = "";

        public Person()
        {
            PERSONNAME = new PersonName();
            DOCUMENT = new Document();
        }
    }

    public class PersonName
    {
        public string SEXTYPESId = "A"; //??? Пол?
        public string LastName = "A";
        public string FirstName = "A";
        public string SecondName = "A";
    }

    public class Document
    {
        public string DOCTYPESId = "A";
        public string Seria = "A";
        public string Number = "A";
        public string GivenBy = "A";
        public string Date = "A";
        public string Birthday = "A";
    }

    public class Address
    {
        public string ZIP = "A";
        public string Country = "A";
        public string Area = "A";
        public string Region = "A";
        public string PLACETYPESId = "A";
        public string PlaceName = "A";
        public string STREETTYPESId = "A";
        public string StreetName = "A";
        public string House = "A";
        public string BUILDINGTYPESId = "A";
        public string Building = "A";
        public string ROOMTYPESId = "";
        public string Room = "A";
    }
}

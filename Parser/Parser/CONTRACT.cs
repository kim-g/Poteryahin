using System.Collections.Generic;

namespace Parser
{
    public class CONTRACT : Serializer
    {
        public string Status;
        public string DealerCode;
        public string DealerPointCode;
        public string DealerContractCode;
        public string DealerContractDate;
        public string ABSContractCode;
        public Customer CUSTOMER = new Customer();
        public string BANKPROPLIST;
        public Delivery DELIVERY = new Delivery();
        public Contact CONTACT = new Contact();
        public Connections CONNECTIONS = new Connections();
        public List<LOGPARAM> LOGPARAMS = new List<LOGPARAM>();
        public string Comments;
        public string CLIENTVER;

        //Сделал полное внесение всех данных в карточку. Осталось сделать вывод в XML
    }

    public class Customer
    {
        public string CUSTOMERTYPESId;
        public string SPHERESId;
        public string Resident;
        public string Ratepayer;
        public Person PERSON = new Person();
        public Address ADDRESS = new Address();
    }

    public class Person
    {
        public string PERSONTYPESId;
        public PersonName PERSONNAME = new PersonName();
        public Document DOCUMENT = new Document();
        public string INN;
    }

    public class PersonName
    {
        public string SEXTYPESId; //??? Пол?
        public string LastName;
        public string FirstName;
        public string SecondName;
    }

    public class Document
    {
        public string DOCTYPESId;
        public string Seria;
        public string Number;
        public string GivenBy;
        public string Date;
        public string Birthday;
    }

    public class Address
    {
        public string ZIP;
        public string Country;
        public string Area;
        public string Region;
        public string PLACETYPESId;
        public string PlaceName;
        public string STREETTYPESId;
        public string StreetName;
        public string House;
        public string BUILDINGTYPESId;
        public string Building;
        public string ROOMTYPESId;
        public string Room;
    }

    public class Delivery
    {
        public string DELIVERYTYPESId;
        public Address ADDRESS = new Address();
        public string Notes;
    }

    public class Contact
    {
        public PersonName PERSONNAME = new PersonName();
        public string PhonePrefix;
        public string Phone;
        public string FaxPrefix;
        public string Fax;
        public string EMail;
        public string PagerOperatorPrefix;
        public string PagerOperator;
        public string PagerAbonent;
        public string Notes;
    }

    public class Connections
    {
        public Connection CONNECTION = new Connection();
    }

    public class Connection
    {
        public string PAYSYSTEMSId;
        public string BILLCYCLESId;
        public string CELLNETSId;
        public string PRODUCTSId;
        public string PhoneOwner;
        public string SerNumber;
        public string SimLock;
        public string IMSI;
        public Mobiles MOBILES = new Mobiles();
    }

    public class Mobiles
    {
        public Mobile MOBILE = new Mobile();
    }

    public class Mobile
    {
        public string CHANNELTYPESId;
        public string CHANNELLENSId;
        public string SNB;
        public string BILLPLANSId;
        public List<string> SERVICES = new List<string>();
    }

    public class LOGPARAM
    {
        public string LOGPARAMSId;
        public string LOGPARAMSCode;
        public string LOGPARAMSValue;
    }

}

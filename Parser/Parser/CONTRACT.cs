using System;
using System.Collections.Generic;

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
        public Customer CUSTOMER = new Customer();
        public string BANKPROPLIST = "A";
        public Delivery DELIVERY = new Delivery();
        public Contact CONTACT = new Contact();
        public Connections CONNECTIONS = new Connections();
        public List<LOGPARAM> LOGPARAMS = new List<LOGPARAM>();
        public string Comments = "A";
        public string CLIENTVER = "A";

        public CONTRACT()
        {
            LOGPARAMS.Add(new LOGPARAM());
            LOGPARAMS.Add(new LOGPARAM());
            LOGPARAMS.Add(new LOGPARAM());
            LOGPARAMS.Add(new LOGPARAM());
            LOGPARAMS.Add(new LOGPARAM());
        } 
    }

    public class Customer
    {
        public string CUSTOMERTYPESId = "A";
        public string SPHERESId = "A";
        public string Resident = "A";
        public string Ratepayer = "A";
        public Person PERSON = new Person();
        public Address ADDRESS = new Address();
    }

    public class Person
    {
        public string PERSONTYPESId = "A";
        public PersonName PERSONNAME = new PersonName();
        public Document DOCUMENT = new Document();
        public string INN = "";
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

    public class Delivery
    {
        public string DELIVERYTYPESId = "A";
        public Address ADDRESS = new Address();
        public string Notes = "A";
    }

    public class Contact
    {
        public PersonName PERSONNAME = new PersonName();
        public string PhonePrefix = "A";
        public string Phone = "A";
        public string FaxPrefix = "A";
        public string Fax = "A";
        public string EMail = "A";
        public string PagerOperatorPrefix = "A";
        public string PagerOperator = "A";
        public string PagerAbonent = "A";
        public string Notes = "A";
    }

    public class Connections
    {
        public Connection CONNECTION = new Connection();
    }

    public class Connection
    {
        public string PAYSYSTEMSId = "A";
        public string BILLCYCLESId = "A";
        public string CELLNETSId = "A";
        public string PRODUCTSId = "A";
        public string PhoneOwner = "A";
        public string SerNumber = "A";
        public string SimLock = "A";
        public string IMSI = "";
        public Mobiles MOBILES = new Mobiles();
    }

    public class Mobiles
    {
        public Mobile MOBILE = new Mobile();
    }

    public class Mobile
    {
        public string CHANNELTYPESId = "A";
        public string CHANNELLENSId = "A";
        public string SNB = "A";
        public string BILLPLANSId = "A";
        public List<string> SERVICES = new List<string>();

        public Mobile()
        {
            SERVICES.Add("1");
            SERVICES.Add("6");
            SERVICES.Add("8");
            SERVICES.Add("34");
        }
    }

    public class LOGPARAM
    {
        public string LOGPARAMSId = "A";
        public string LOGPARAMSCode = "A";
        public string LOGPARAMSValue = "A";
    }

}

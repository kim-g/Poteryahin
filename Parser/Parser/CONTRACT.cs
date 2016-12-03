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

        public string ToXMLString(string Example)
        {
            string Text = Example;

            // CONTRACT
            Text = Text.Replace("@C_Status@", Status);
            Text = Text.Replace("@C_DealerCode@", DealerCode);
            Text = Text.Replace("@C_DealerPointCode@", DealerPointCode);
            Text = Text.Replace("@C_DealerContractCode@", DealerContractCode);
            Text = Text.Replace("@C_DealerContractDate@", DealerContractDate);
            Text = Text.Replace("@C_ABSContractCode@", ABSContractCode);
            Text = Text.Replace("@C_BANKPROPLIST@", BANKPROPLIST == "" ? "<BANKPROPLIST/>" : "<BANKPROPLIST>" + BANKPROPLIST + "</BANKPROPLIST>");
            Text = Text.Replace("@C_Comments@", Comments);
            Text = Text.Replace("@C_CLIENTVER@", CLIENTVER);

            // CONTRACT.CUSTOMER
            Text = Text.Replace("@C_C_CUSTOMERTYPESId@", CUSTOMER.CUSTOMERTYPESId);
            Text = Text.Replace("@C_C_SPHERESId@", CUSTOMER.SPHERESId);
            Text = Text.Replace("@C_C_Resident@", CUSTOMER.Resident);
            Text = Text.Replace("@C_C_Ratepayer@", CUSTOMER.Ratepayer);

            // CONTRACT.CUSTOMER.PERSON
            Text = Text.Replace("@C_C_P_PERSONTYPESId@", CUSTOMER.PERSON.PERSONTYPESId);
            Text = Text.Replace("@C_C_P_INN@", CUSTOMER.PERSON.INN);

            // CONTRACT.CUSTOMER.PERSON.PERSONNAME
            Text = Text.Replace("@C_C_P_P_SEXTYPESId@", CUSTOMER.PERSON.PERSONNAME.SEXTYPESId);
            Text = Text.Replace("@C_C_P_P_LastName@", CUSTOMER.PERSON.PERSONNAME.LastName);
            Text = Text.Replace("@C_C_P_P_FirstName@", CUSTOMER.PERSON.PERSONNAME.FirstName);
            Text = Text.Replace("@C_C_P_P_SecondName@", CUSTOMER.PERSON.PERSONNAME.SecondName);

            // CONTRACT.CUSTOMER.PERSON.DOCUMENT
            Text = Text.Replace("@C_C_P_D_DOCTYPESId@", CUSTOMER.PERSON.DOCUMENT.DOCTYPESId);
            Text = Text.Replace("@C_C_P_D_Seria@", CUSTOMER.PERSON.DOCUMENT.Seria);
            Text = Text.Replace("@C_C_P_D_Number@", CUSTOMER.PERSON.DOCUMENT.Number);
            Text = Text.Replace("@C_C_P_D_GivenBy@", CUSTOMER.PERSON.DOCUMENT.GivenBy);
            Text = Text.Replace("@C_C_P_D_Date@", CUSTOMER.PERSON.DOCUMENT.Date);
            Text = Text.Replace("@C_C_P_D_Birthday@", CUSTOMER.PERSON.DOCUMENT.Birthday);

            // CONTRACT.CUSTOMER.ADDRESS
            Text = Text.Replace("@C_C_A_ZIP@", CUSTOMER.ADDRESS.ZIP);
            Text = Text.Replace("@C_C_A_Country@", CUSTOMER.ADDRESS.Country);
            Text = Text.Replace("@C_C_A_Area@", CUSTOMER.ADDRESS.Area);
            Text = Text.Replace("@C_C_A_Region@", CUSTOMER.ADDRESS.Region);
            Text = Text.Replace("@C_C_A_PLACETYPESId@", CUSTOMER.ADDRESS.PLACETYPESId);
            Text = Text.Replace("@C_C_A_PlaceName@", CUSTOMER.ADDRESS.PlaceName);
            Text = Text.Replace("@C_C_A_STREETTYPESId@", CUSTOMER.ADDRESS.STREETTYPESId);
            Text = Text.Replace("@C_C_A_StreetName@", CUSTOMER.ADDRESS.StreetName);
            Text = Text.Replace("@C_C_A_House@", CUSTOMER.ADDRESS.House);
            Text = Text.Replace("@C_C_A_BUILDINGTYPESId@", CUSTOMER.ADDRESS.BUILDINGTYPESId);
            Text = Text.Replace("@C_C_A_Building@", CUSTOMER.ADDRESS.Building);
            Text = Text.Replace("@C_C_A_ROOMTYPESId@", CUSTOMER.ADDRESS.ROOMTYPESId);
            Text = Text.Replace("@C_C_A_Room@", CUSTOMER.ADDRESS.Room);

            // CONTRACT.DELIVERY
            Text = Text.Replace("@C_D_DELIVERYTYPESId@", DELIVERY.DELIVERYTYPESId);
            Text = Text.Replace("@C_D_Notes@", DELIVERY.Notes);

            // CONTRACT.DELIVERY.ADDRESS
            Text = Text.Replace("@C_D_A_ZIP@", DELIVERY.ADDRESS.ZIP);
            Text = Text.Replace("@C_D_A_Country@", DELIVERY.ADDRESS.Country);
            Text = Text.Replace("@C_D_A_Area@", DELIVERY.ADDRESS.Area);
            Text = Text.Replace("@C_D_A_Region@", DELIVERY.ADDRESS.Region);
            Text = Text.Replace("@C_D_A_PLACETYPESId@", DELIVERY.ADDRESS.PLACETYPESId);
            Text = Text.Replace("@C_D_A_PlaceName@", DELIVERY.ADDRESS.PlaceName);
            Text = Text.Replace("@C_D_A_STREETTYPESId@", DELIVERY.ADDRESS.STREETTYPESId);
            Text = Text.Replace("@C_D_A_StreetName@", DELIVERY.ADDRESS.StreetName);
            Text = Text.Replace("@C_D_A_House@", DELIVERY.ADDRESS.House);
            Text = Text.Replace("@C_D_A_BUILDINGTYPESId@", DELIVERY.ADDRESS.BUILDINGTYPESId);
            Text = Text.Replace("@C_D_A_Building@", DELIVERY.ADDRESS.Building);
            Text = Text.Replace("@C_D_A_ROOMTYPESId@", DELIVERY.ADDRESS.ROOMTYPESId);
            Text = Text.Replace("@C_D_A_Room@", DELIVERY.ADDRESS.Room);

            // CONTRACT.CONTACT
            Text = Text.Replace("@C_Co_PhonePrefix@", CONTACT.PhonePrefix);
            Text = Text.Replace("@C_Co_Phone@", CONTACT.Phone);
            Text = Text.Replace("@C_Co_FaxPrefix@", CONTACT.FaxPrefix);
            Text = Text.Replace("@C_Co_Fax@", CONTACT.Fax);
            Text = Text.Replace("@C_Co_EMail@", CONTACT.EMail);
            Text = Text.Replace("@C_Co_PagerOperatorPrefix@", CONTACT.PagerOperatorPrefix);
            Text = Text.Replace("@C_Co_PagerOperator@", CONTACT.PagerOperator);
            Text = Text.Replace("@C_Co_PagerAbonent@", CONTACT.PagerAbonent);
            Text = Text.Replace("@C_Co_Notes@", CONTACT.Notes);

            // CONTRACT.CONTACT.PERSONNAME
            Text = Text.Replace("@C_Co_P_SEXTYPESId@", CONTACT.PERSONNAME.SEXTYPESId);
            Text = Text.Replace("@C_Co_P_LastName@", CONTACT.PERSONNAME.LastName);
            Text = Text.Replace("@C_Co_P_FirstName@", CONTACT.PERSONNAME.FirstName);
            Text = Text.Replace("@C_Co_P_SecondName@", CONTACT.PERSONNAME.SecondName);

            // CONTRACT.CONNECTIONS.CONNECTION
            Text = Text.Replace("@C_Con_C_PAYSYSTEMSId@", CONNECTIONS.CONNECTION.PAYSYSTEMSId);
            Text = Text.Replace("@C_Con_C_BILLCYCLESId@", CONNECTIONS.CONNECTION.BILLCYCLESId);
            Text = Text.Replace("@C_Con_C_CELLNETSId@", CONNECTIONS.CONNECTION.CELLNETSId);
            Text = Text.Replace("@C_Con_C_PRODUCTSId@", CONNECTIONS.CONNECTION.PRODUCTSId);
            Text = Text.Replace("@C_Con_C_PhoneOwner@", CONNECTIONS.CONNECTION.PhoneOwner);
            Text = Text.Replace("@C_Con_C_SerNumber@", CONNECTIONS.CONNECTION.SerNumber);
            Text = Text.Replace("@C_Con_C_SimLock@", CONNECTIONS.CONNECTION.SimLock);
            Text = Text.Replace("@C_Con_C_IMSI@", CONNECTIONS.CONNECTION.IMSI);

            // CONTRACT.CONNECTIONS.CONNECTION.MOBILES.MOBILE
            Text = Text.Replace("@C_Con_C_M_M_CHANNELTYPESId@", CONNECTIONS.CONNECTION.MOBILES.MOBILE.CHANNELTYPESId);
            Text = Text.Replace("@C_Con_C_M_M_CHANNELLENSId@", CONNECTIONS.CONNECTION.MOBILES.MOBILE.CHANNELLENSId);
            Text = Text.Replace("@C_Con_C_M_M_SNB@", CONNECTIONS.CONNECTION.MOBILES.MOBILE.SNB);
            Text = Text.Replace("@C_Con_C_M_M_BILLPLANSId@", CONNECTIONS.CONNECTION.MOBILES.MOBILE.BILLPLANSId);

            // CONTRACT.CONNECTIONS.CONNECTION.MOBILES.MOBILE.SERVICES
            string Serv;
            if (CONNECTIONS.CONNECTION.MOBILES.MOBILE.SERVICES.Count == 0)
                Serv = "<SERVICES/>";
            else
            {
                Serv = "<SERVICES>";
                for (int i = 0; i < CONNECTIONS.CONNECTION.MOBILES.MOBILE.SERVICES.Count; i++)
                    Serv += "<SERVICESId>" + CONNECTIONS.CONNECTION.MOBILES.MOBILE.SERVICES[i] + "</SERVICESId>";
                Serv += "</SERVICES>";
            }
            Text = Text.Replace("@C_Con_C_M_M_SERVICES@", Serv);

            // CONTRACT.LOGPARAMS
            string LogP;
            if (LOGPARAMS.Count == 0)
                LogP= "<LOGPARAMS/>";
            else
            {
                LogP= "<LOGPARAMS>";
                for (int i = 0; i < LOGPARAMS.Count; i++)
                    LogP += "<LOGPARAM><LOGPARAMSId>" + LOGPARAMS[i].LOGPARAMSId + 
                        "</LOGPARAMSId><LOGPARAMSCode>" + LOGPARAMS[i].LOGPARAMSCode +
                        "</LOGPARAMSCode><LOGPARAMSValue>" + LOGPARAMS[i].LOGPARAMSValue +
                        "</LOGPARAMSValue></LOGPARAM>";
                LogP += "</LOGPARAMS>";
            }
            Text = Text.Replace("@C_LOGPARAMS@", LogP);


            return Text;
        }
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

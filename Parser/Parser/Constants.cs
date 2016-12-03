using System.Collections.Generic;

namespace Parser
{
    public class Constants
    {
        // Напрямую из CONTRACT
        public string Status = "0";
        public string DealerCode = "3T33";
        public string DealerPointCode = "3T33001";
        public string DealerContractCode = "1063";
        public string DealerContractDate = "2016-04-23";
        public string ABSContractCode = "";
        public string BANKPROPLIST = "";
        public string Comments = "";
        public string CLIENTVER = "10.5.0.1024";

        //CONTRACT.CUSTOMER
        public string CUSTOMERTYPESId = "0";
        public string SPHERESId = "12";
        public string Resident = "1";
        public string Ratepayer = "1";
        public string INN = "";

        //CONTRACT.CUSTOMER.PERSON
        public string PERSONTYPESId = "0";

        //CONTRACT.CUSTOMER.PERSON.PERSONNAME
        public string SEXTYPESId = "0";

        //CONTRACT.CUSTOMER.ADDRESS
        public string Region = "";

        //CONTRACT.DELIVERY
        public string DELIVERYTYPESId = "1";
        public string Notes = "";

        //CONTRACT.CONTACT
        public string PhonePrefix = "999";
        public string Phone = "9999999";
        public string FaxPrefix = "";
        public string Fax = "";
        public string EMail = "";
        public string PagerOperatorPrefix = "";
        public string PagerOperator = "";
        public string PagerAbonent = "";
        public string Contact_Notes = "";

        //CONTRACT.CONTACT.PERSONNAME
        public string CP_FirstName = "";
        public string CP_SecondName = "";

        //CONTRACT.CONNECTIONS.CONNECTION
        public string PAYSYSTEMSId = "3";
        public string BILLCYCLESId = "-1";
        public string CELLNETSId = "2";
        public string PRODUCTSId = "1000";
        public string PhoneOwner = "1";
        public string SerNumber = "000000000000000";
        public string SimLock = "0";

        //CONTRACT.CONNECTIONS.CONNECTION.MOBILES.MOBILE
        public string CHANNELTYPESId = "1";
        public string CHANNELLENSId = "0";

        //CONTRACT.LOGPARAMS
        public List<LOGPARAM> LOGPARAMS = new List<LOGPARAM>();
    }
}

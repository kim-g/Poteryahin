using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Parser
{
    public class Constants : Serializer
    {
        // Напрямую из CONTRACT
        public string Status = "0";
        public string DealerCode = "3T33";
        public string DealerPointCode = "3T33001";
        public string DealerContractCode = "1063";
        public string DealerContractDate = "2016-04-23";
        public string ABSContractCode = "";

        //CONTRACT.CUSTOMER
        public string CUSTOMERTYPESId = "0";
        public string SPHERESId = "12";
        public string Resident = "1";
        public string Ratepayer = "1";

        //CONTRACT.CUSTOMER.PERSON
        public string PERSONTYPESId = "0";

        //CONTRACT.CUSTOMER.PERSON.PERSONNAME
        public string SEXTYPESId = "0";

    }
}

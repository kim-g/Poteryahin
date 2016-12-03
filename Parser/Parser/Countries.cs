using System.Collections.Generic;

namespace Parser
{
    public class Countries : Serializer
    {
        public List<Country> CountryList = new List<Country>();

        public string GetCountryName(string ID)
        {
            var found = CountryList.FindAll(p => p.ID == ID);
            if (found.Count == 0) return null;
            return found[0].Name;
        }
    }
}

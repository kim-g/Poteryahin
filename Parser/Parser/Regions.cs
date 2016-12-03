using System.Collections.Generic;

namespace Parser
{
    public class Regions : Serializer
    {
        public List<Region> RegionList = new List<Region>();

        public string GetZIP(string ID)
        {
            var found = RegionList.FindAll(p => p.ID == ID);
            if (found.Count == 0) return null;
            return found[0].ZIP;
        }

        public string GetRegion(string ID)
        {
            var found = RegionList.FindAll(p => p.ID == ID);
            if (found.Count == 0) return null;
            return found[0].State;
        }
    }

    
}

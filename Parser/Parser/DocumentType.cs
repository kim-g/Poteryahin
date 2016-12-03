using System.Collections.Generic;

namespace Parser
{
    public class DocumentType : Serializer
    {
        public List<DocumentInfo> DocumentList = new List<DocumentInfo>();

        public string GetID(string Name)
        {
            var found = DocumentList.FindAll(p => p.Name == Name);
            if (found.Count == 0) return null;
            return found[0].ID;
        }
    }


}

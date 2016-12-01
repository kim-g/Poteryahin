using System;
using System.IO;
using System.Xml.Serialization;

namespace Parser
{
    public class Colomn_Numbers :Serializer
    {
        public string Region ="A";
        public string BAN = "A";
        public string CTN = "A";
        public string TITLE = "A";
        public string LAST_N = "A";
        public string FIRST_N = "A";
        public string PATRONYME = "A";
        public string GENDER = "A";
        public string BIRTH_DATE = "A";
        public string DOCUMENTTYPE = "A";
        public string DOCUMENT_N = "A";
        public string DOCUMENTID = "A";
        public string DOCISSUDATE = "A";
        public string DOCISSUORIGINE = "A";
        public string COUNTRY = "A";
        public string PLACETYPE = "A";
        public string PLACENAMECITY = "A";
        public string STREETTYPE = "A";
        public string STREETNAME = "A";
        public string HOUSE_NO = "A";
        public string BUILDINGTYPE = "A";
        public string BUILDING_NO = "A";
        public string APARTTYPE = "A";
        public string APPARTEMENT_NO = "A";

        /*public void SaveToXML(String FileName)
        {
            using (Stream writer = new FileStream(FileName, FileMode.Create))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Colomn_Numbers));
                serializer.Serialize(writer, this);
            }
        }

        public static Colomn_Numbers LoadFromXML(String FileName)
        {
            // загружаем данные из файла FileName
            using (Stream stream = new FileStream(FileName, FileMode.Open))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Colomn_Numbers));

                // в тут же созданную копию класса Serializer под именем ser
                Colomn_Numbers ser = (Colomn_Numbers)serializer.Deserialize(stream);
                return ser;
            }
        }*/
    }
}

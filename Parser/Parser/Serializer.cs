using System;
using System.IO;
using System.Xml.Serialization;

namespace Parser
{
    public class Serializer
    {
        public void SaveToXML(String FileName)
        {
            using (Stream writer = new FileStream(FileName, FileMode.Create))
            {
                XmlSerializer serializer = new XmlSerializer(this.GetType());
                serializer.Serialize(writer, this);
            }
        }

        public static Serializer LoadFromXML(String FileName, Type type)
        {
            // загружаем данные из файла FileName
            using (Stream stream = new FileStream(FileName, FileMode.Open))
            {
                XmlSerializer serializer = new XmlSerializer(type);

                // в тут же созданную копию класса Serializer под именем ser
                Serializer ser = (Serializer)serializer.Deserialize(stream);
                return ser;
            }
        }
    }
}

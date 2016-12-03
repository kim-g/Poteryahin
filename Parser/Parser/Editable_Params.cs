using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Parser
{
    public class Editable_Params : Serializer
    {
        public Constants Const = new Constants();   // Константы карточек
        public List<Country> CountryList = new List<Country>(); // Список стран
        public List<DocumentInfo> DocumentList = new List<DocumentInfo>();  // Список типов документов
        public List<Region> RegionList = new List<Region>();    // Список регионов
        public List<CityType> CityTypes = new List<CityType>(); // Списое номеров типов нас. пунктов (Город, деревня, посёлок...)
        public List<StreetType> StreetTypes = new List<StreetType>(); // Списое номеров типов улиц (Улица, переулок, проспект, бульвар...)
        public List<BuildingType> BuildingTypeList = new List<BuildingType>();    // Список типов зданий
        public List<Building> BuildingList = new List<Building>(); // Список номеров строений
        public List<RoomType> RoomTypeList = new List<RoomType>(); // Списое типов комнат

        // Получить страну по ID
        public string GetCountryName(string ID)
        {
            var found = CountryList.FindAll(p => p.ID == ID);
            if (found.Count == 0) return null;
            return found[0].Name;
        }

        // Получить документ по ID
        public string GetID(string Name)
        {
            var found = DocumentList.FindAll(p => p.Name == Name);
            if (found.Count == 0) return null;
            return found[0].ID;
        }

        // Получить почтовый индекс по ID региона
        public string GetZIP(string ID)
        {
            var found = RegionList.FindAll(p => p.ID == ID);
            if (found.Count == 0) return null;
            return found[0].ZIP;
        }

        //Получить область по ID региона
        public string GetRegion(string ID)
        {
            var found = RegionList.FindAll(p => p.ID == ID);
            if (found.Count == 0) return null;
            return found[0].State;
        }

        //Получить ID типа нас. пункта по названию 
        public string GetCityID(string Name)
        {
            var found = CityTypes.FindAll(p => p.Name == Name);
            if (found.Count == 0) return null;
            return found[0].ID;
        }

        //Получить ID типа улицы по названию 
        public string GetStreetID(string Name)
        {
            var found = StreetTypes.FindAll(p => p.Name == Name);
            if (found.Count == 0) return null;
            return found[0].ID;
        }

        //Получить ID типа здания по названию 
        public string GetBuildingTypeID(string Name)
        {
            var found = BuildingTypeList.FindAll(p => p.Name == Name);
            if (found.Count == 0) return null;
            return found[0].ID;
        }

        //Получить ID типа строения по названию 
        public string GetBuildingID(string Name)
        {
            var found = BuildingList.FindAll(p => p.Name == Name);
            if (found.Count == 0) return null;
            return found[0].ID;
        }

        //Получить ID типа Помещения по названию 
        public string GetRoomTypeID(string Name)
        {
            var found = RoomTypeList.FindAll(p => p.Name == Name);
            if (found.Count == 0) return null;
            return found[0].ID;
        }
    }

    public class Country
    {
        public string ID;
        public string Name;
    }

    public class DocumentInfo
    {
        public string Name;
        public string ID;
    }

    public class Region
    {
        public string ID;
        public string ZIP;
        public string State;
    }

    public class CityType
    {
        public string ID;
        public string Name;
    }

    public class StreetType
    {
        public string ID;
        public string Name;
    }

    public class BuildingType
    {
        public string ID;
        public string Name;
    }

    public class Building
    {
        public string ID;
        public string Name;
    }

    public class RoomType
    {
        public string ID;
        public string Name;
    }
}

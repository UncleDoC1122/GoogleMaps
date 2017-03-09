using Google.Maps.Geocoding;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Google_Maps_Address_Qualifier
{
    class Savers
    {
        private void SaveToXmlShort(List<Result> input, List<string> original, string path)
        {
            XmlDocument Doc = new XmlDocument();
            XmlTextWriter Writer = new XmlTextWriter(path, Encoding.UTF8);
            Writer.WriteStartDocument();
            Writer.WriteEndDocument();
            Writer.Close();
            Doc.Load(path);
            int i = 0;

            foreach (var element in input)
            {
                if (element.FormattedAddress == null)
                {
                    XmlNode Element = Doc.CreateElement("BusinessPartner");
                    Doc.DocumentElement.AppendChild(Element);

                    XmlNode Intcode = Doc.CreateElement("Intcode");
                    Intcode.InnerText = original[i];
                    Element.AppendChild(Intcode);

                    XmlNode Region = Doc.CreateElement("Region");
                    Region.InnerText = "-";
                    Element.AppendChild(Region);

                    XmlNode Index = Doc.CreateElement("Index");
                    Index.InnerText = "-";
                    Element.AppendChild(Index);

                    XmlNode City = Doc.CreateElement("City");
                    City.InnerText = "-";
                    Element.AppendChild(City);

                    XmlNode Street = Doc.CreateElement("Street");
                    Street.InnerText = "-";
                    Element.AppendChild(Street);

                    XmlNode House = Doc.CreateElement("House");
                    House.InnerText = "-";
                    Element.AppendChild(House);

                    XmlNode Apartments = Doc.CreateElement("Apartments");
                    Apartments.InnerText = "-";
                    Element.AppendChild(Apartments);

                    XmlNode Full = Doc.CreateElement("Full");
                    Full.InnerText = "-";
                    Element.AppendChild(Full);

                    XmlNode IsFound = Doc.CreateElement("IsFound");
                    IsFound.InnerText = "false";
                    Element.AppendChild(IsFound);

                    i++;
                }
                else
                {
                    XmlNode Intcode = Doc.CreateElement("Intcode");
                    XmlNode Region = Doc.CreateElement("Region");
                    XmlNode Index = Doc.CreateElement("Index");
                    XmlNode City = Doc.CreateElement("City");
                    XmlNode Street = Doc.CreateElement("Street");
                    XmlNode House = Doc.CreateElement("House");
                    XmlNode Apartments = Doc.CreateElement("Apartments");
                    XmlNode Full = Doc.CreateElement("Full");
                    XmlNode IsFound = Doc.CreateElement("IsFound");

                    Intcode.InnerText = original[i];
                    Region.InnerText = "-";
                    Index.InnerText = "-";
                    City.InnerText = "-";
                    Street.InnerText = "-";
                    House.InnerText = "-";
                    Apartments.InnerText = "-";
                    Full.InnerText = "-";
                    IsFound.InnerText = "true";

                    var components = element.AddressComponents;

                    foreach (var component in components)
                    {
                        var types = component.Types;

                        Full.InnerText = element.FormattedAddress;

                        foreach (var type in types)
                        {
                            switch (type)
                            {
                                case Google.Maps.Shared.AddressType.PostalCode:
                                    Index.InnerText = component.LongName;
                                    break;

                                case Google.Maps.Shared.AddressType.Country:
                                    Region.InnerText = component.LongName;
                                    break;

                                case Google.Maps.Shared.AddressType.Locality:
                                    City.InnerText = component.LongName;
                                    break;

                                case Google.Maps.Shared.AddressType.Sublocality:
                                    if (City.InnerText.Equals("-"))
                                        City.InnerText = component.LongName;
                                    break;

                                case Google.Maps.Shared.AddressType.Route:
                                    Street.InnerText = component.LongName;
                                    break;

                                case Google.Maps.Shared.AddressType.Intersection:
                                    if (Street.InnerText.Equals("-"))
                                        Street.InnerText = component.LongName;
                                    break;

                                case Google.Maps.Shared.AddressType.Premise:
                                    if (Street.InnerText.Equals("-"))
                                        Street.InnerText = component.LongName;
                                    break;

                                case Google.Maps.Shared.AddressType.StreetNumber:
                                    House.InnerText = component.LongName;
                                    break;

                                case Google.Maps.Shared.AddressType.Room:
                                    Apartments.InnerText = component.LongName;
                                    break;
                            }
                        }
                    }

                    XmlNode Element = Doc.CreateElement("BusinessPartner");
                    Doc.DocumentElement.AppendChild(Element);



                    Element.AppendChild(Intcode);
                    Element.AppendChild(Region);
                    Element.AppendChild(Index);
                    Element.AppendChild(City);
                    Element.AppendChild(Street);
                    Element.AppendChild(House);
                    Element.AppendChild(Apartments);
                    Element.AppendChild(Full);
                    Element.AppendChild(IsFound);

                    i++;
                }
            }
        }

        private void SaveToXlsShort(List<Result> input, List<string> original, string cellIndex, string cellRegion, string cellCity,
                               string cellStreet, string cellHouse, string cellApp, string cellFull, string cellIsFound)
        {

        }

        private void SaveToXmlFull(List<Result> input, List<List<string>> original, List<string> headers, string path)
        {
            XmlDocument Doc = new XmlDocument();
            XmlTextWriter Writer = new XmlTextWriter(path, Encoding.UTF8);
            Writer.WriteStartDocument();
            Writer.WriteEndDocument();
            Writer.Close();
            Doc.Load(path);
            int i = 0;
            

            foreach (var element in input)
            {
                if (element.FormattedAddress == null)
                {
                    XmlNode Element = Doc.CreateElement("BusinessPartner");
                    Doc.DocumentElement.AppendChild(Element);

                    for (int j = 0; j < original[i].Count; j++) 
                    {
                        XmlNode Original = Doc.CreateElement(headers[j]);
                        Original.InnerText = original[i][j];
                        Element.AppendChild(Original);
                    }

                    XmlNode Region = Doc.CreateElement("Region");
                    Region.InnerText = "-";
                    Element.AppendChild(Region);

                    XmlNode Index = Doc.CreateElement("Index");
                    Index.InnerText = "-";
                    Element.AppendChild(Index);

                    XmlNode City = Doc.CreateElement("City");
                    City.InnerText = "-";
                    Element.AppendChild(City);

                    XmlNode Street = Doc.CreateElement("Street");
                    Street.InnerText = "-";
                    Element.AppendChild(Street);

                    XmlNode House = Doc.CreateElement("House");
                    House.InnerText = "-";
                    Element.AppendChild(House);

                    XmlNode Apartments = Doc.CreateElement("Apartments");
                    Apartments.InnerText = "-";
                    Element.AppendChild(Apartments);

                    XmlNode Full = Doc.CreateElement("Full");
                    Full.InnerText = "-";
                    Element.AppendChild(Full);

                    XmlNode IsFound = Doc.CreateElement("IsFound");
                    IsFound.InnerText = "false";
                    Element.AppendChild(IsFound);

                    i++;
                }
                else
                {
                    XmlNode Intcode = Doc.CreateElement("Intcode");
                    XmlNode Region = Doc.CreateElement("Region");
                    XmlNode Index = Doc.CreateElement("Index");
                    XmlNode City = Doc.CreateElement("City");
                    XmlNode Street = Doc.CreateElement("Street");
                    XmlNode House = Doc.CreateElement("House");
                    XmlNode Apartments = Doc.CreateElement("Apartments");
                    XmlNode Full = Doc.CreateElement("Full");
                    XmlNode IsFound = Doc.CreateElement("IsFound");

                    XmlNode Element = Doc.CreateElement("BusinessPartner");
                    Doc.DocumentElement.AppendChild(Element);

                    for (int j = 0; j < original[i].Count; j++)
                    {
                        XmlNode Original = Doc.CreateElement(headers[j]);
                        Original.InnerText = original[i][j];
                        Element.AppendChild(Original);
                    }
                    Region.InnerText = "-";
                    Index.InnerText = "-";
                    City.InnerText = "-";
                    Street.InnerText = "-";
                    House.InnerText = "-";
                    Apartments.InnerText = "-";
                    Full.InnerText = "-";
                    IsFound.InnerText = "true";

                    var components = element.AddressComponents;

                    foreach (var component in components)
                    {
                        var types = component.Types;

                        Full.InnerText = element.FormattedAddress;

                        foreach (var type in types)
                        {
                            switch (type)
                            {
                                case Google.Maps.Shared.AddressType.PostalCode:
                                    Index.InnerText = component.LongName;
                                    break;

                                case Google.Maps.Shared.AddressType.Country:
                                    Region.InnerText = component.LongName;
                                    break;

                                case Google.Maps.Shared.AddressType.Locality:
                                    City.InnerText = component.LongName;
                                    break;

                                case Google.Maps.Shared.AddressType.Sublocality:
                                    if (City.InnerText.Equals("-"))
                                        City.InnerText = component.LongName;
                                    break;

                                case Google.Maps.Shared.AddressType.Route:
                                    Street.InnerText = component.LongName;
                                    break;

                                case Google.Maps.Shared.AddressType.Intersection:
                                    if (Street.InnerText.Equals("-"))
                                        Street.InnerText = component.LongName;
                                    break;

                                case Google.Maps.Shared.AddressType.Premise:
                                    if (Street.InnerText.Equals("-"))
                                        Street.InnerText = component.LongName;
                                    break;

                                case Google.Maps.Shared.AddressType.StreetNumber:
                                    House.InnerText = component.LongName;
                                    break;

                                case Google.Maps.Shared.AddressType.Room:
                                    Apartments.InnerText = component.LongName;
                                    break;
                            }
                        }
                    }

                   



                    Element.AppendChild(Intcode);
                    Element.AppendChild(Region);
                    Element.AppendChild(Index);
                    Element.AppendChild(City);
                    Element.AppendChild(Street);
                    Element.AppendChild(House);
                    Element.AppendChild(Apartments);
                    Element.AppendChild(Full);
                    Element.AppendChild(IsFound);

                    i++;
                }
            }
        }

    }
}

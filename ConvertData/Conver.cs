using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml;
using Newtonsoft.Json;


namespace ConvertData
{

    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface iconvert
    {
        bool xml2json(string InPath, string OutPath, bool considerXMLRootObj);
        bool json2xml(string InPath, string OutPath, string rootNodeXML);
        string Base64ToString(string Base64Text);

        //Property
        string Status { get; set; }

    }


    [ClassInterface(ClassInterfaceType.None)]
    public class Convert : iconvert
    {
        //Globals
        public string Status { get; set; }


        public bool xml2json(string InPath, string OutPath, bool considerXMLRootObj)
        {

            XmlDocument _domDoc = new XmlDocument();
            string _jsonTxt;
            
            try
            {
                _domDoc.Load(InPath);

                //convert XML2Json
                if (considerXMLRootObj)
                {
                    _jsonTxt = JsonConvert.SerializeXmlNode(_domDoc);
                }
                else
                {
                    _jsonTxt = JsonConvert.SerializeXmlNode(_domDoc,Newtonsoft.Json.Formatting.None, true);
                }

                Newtonsoft.Json.Linq.JObject _jsonObj = Newtonsoft.Json.Linq.JObject.Parse(_jsonTxt);
                File.WriteAllText(OutPath, _jsonObj.ToString());

                Status = "okay";
                return (true);
            }
            catch (Exception ex)
            {
                Status = ex.ToString();
                return (false);                
            }

        }

        public bool json2xml(string InPath, string OutPath, string rootNodeXML)
        {

            string _jsonTxt;            
            System.Xml.Linq.XNode xml;

            try
            {
                //convert json2xml
                _jsonTxt = File.ReadAllText(InPath);
                xml = JsonConvert.DeserializeXNode(_jsonTxt, rootNodeXML).FirstNode;
                File.WriteAllText(OutPath, xml.ToString());

                Status = "okay";
                return (true);
            }
            catch (Exception ex)
            {
                Status = ex.ToString();
                return (false);
            }

        }

        public string Base64ToString(string Base64Text)
        {
            try
            {
                Byte []b = System.Convert.FromBase64String(Base64Text);
                Status = "okay";
                return (Encoding.UTF8.GetString(b));
            }
            catch (Exception ex)
            {
                Status = ex.ToString();
                return (string.Empty);                
            }
        }

    }
}

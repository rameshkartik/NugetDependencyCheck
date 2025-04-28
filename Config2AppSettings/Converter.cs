using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using System.Xml;

namespace Config2AppSettings
{
    public class ConverterLibrary
    {
        private string sInputFileLocation = string.Empty;
        private string sOutputFileLocation = string.Empty;
        JObject jParent = new JObject();

        public ConverterLibrary() { }

        public ConverterLibrary(string InputFile,string OutputFile) 
        {
            File.WriteAllText(OutputFile, string.Empty);
            sInputFileLocation = InputFile;
            sOutputFileLocation = OutputFile;
        }

        public void Convert2AppSettings()
        {
           LoadAppSettings();

           LoadCacheSettings();
        }

        private void LoadAppSettings()
        {
            var jsonObj = LoadAndSetUpKeyValuePairs("/configuration/appSettings/add", "appSettings");

            Write2JsonFile(jsonObj, "appSettings");
        }

        private void LoadCacheSettings()
        {
            var jsonObj = LoadAndSetUpKeyValuePairs("/configuration/CacheConfiguration/Caches/Cache", "CacheConfiguration");

            Write2JsonFile(jsonObj, "CacheConfiguration");
        }


        private void Write2JsonFile(JObject keyValuePairs,string jobj)
        {
            JsonConvert.SerializeObject(keyValuePairs, Newtonsoft.Json.Formatting.Indented);
            File.WriteAllText(sOutputFileLocation, keyValuePairs.ToString());
        }
        private JObject LoadAndSetUpKeyValuePairs(string xpath,string section)
        {
            XmlDocument xmlDocConfig = new XmlDocument();
            xmlDocConfig.Load(sInputFileLocation);


            var appSettingsNodes = xmlDocConfig.SelectNodes(xpath);

            var dict = new Dictionary<string, string>();
            var appKey = string.Empty;
            JObject jObjChild = new JObject();
            JObject jObjSubChild = new JObject();
            string strParent = string.Empty;

            if (appSettingsNodes != null)
            {
                foreach (XmlNode xmlNode in appSettingsNodes)
                {
                    if (xmlNode is not null && xmlNode?.Attributes?.Count > 0 && xmlNode.Attributes.Count == 2)
                    {
                        jObjChild.Add(xmlNode?.Attributes[0]?.Value, xmlNode?.Attributes[1]?.Value);

                    }
                    else if (xmlNode is not null && xmlNode?.Attributes?.Count > 0 && xmlNode.Attributes.Count > 2)
                    {
                        int iCnt = 0;
                        jObjSubChild = new JObject();

                        foreach (XmlAttribute xmlAttribute in xmlNode.Attributes)
                        {
                            if (iCnt == 0) { iCnt++; strParent = xmlAttribute.Value;  continue; }

                            jObjSubChild.Add(xmlAttribute.Name, xmlAttribute.Value);

                        }
                        jObjChild.Add(strParent, jObjSubChild);

                    }
                }

                jParent.Add(section, jObjChild);
            }

            return jParent;
        }
    }
}

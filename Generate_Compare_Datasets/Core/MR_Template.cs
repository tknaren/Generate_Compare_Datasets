using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Generate_Compare_Datasets.Core
{
    [Serializable]
    public class MRColumnDef
    {
        public string nameInExcel { get; set; }
        public string nameInXML { get; set; }
        public string dataType { get; set; }
        public int order { get; set; }
        public string? pair { get; set; }
        public bool? isNullAllowed { get; set; }
    }

    [Serializable]
    [XmlRoot("template")]
    public class MRTemplate
    {
        [XmlElement]
        public List<MasterRecipe> masterRecipes { get; set; }
       
    }

    [Serializable]
    [XmlRoot("masterRecipe")]
    public class MasterRecipe
    {
        [XmlAttribute]
        public string tag { get; set; }
        [XmlAttribute]
        public int version { get; set; }
        [XmlAttribute]
        public string description { get; set; }
        [XmlAttribute]
        public string state { get; set; }
        [XmlAttribute]
        public string effective { get; set; }
        [XmlAttribute]
        public string expiration { get; set; }
        [XmlAttribute]
        public string UId { get; set; }
        [XmlAttribute]
        public string orig_version { get; set; }
        [XmlAttribute]
        public string resPlan { get; set; }
        [XmlAttribute]
        public string resActv { get; set; }
        [XmlAttribute]
        public string RPL_tag { get; set; }
        [XmlAttribute]
        public string RPL_ver { get; set; }
        [XmlElement]
        public List<MRParam> mrParams { get; set; }
    }

    [Serializable]
    [XmlRoot("mrParam")]
    public class MRParam
    {
        [XmlAttribute]
        public string tag { get; set; }
        [XmlAttribute]
        public int version { get; set; }
        [XmlAttribute]
        public string value { get; set; }
        [XmlAttribute]
        public string resPlan { get; set; }
        [XmlAttribute]
        public string resActv { get; set; }
    }

}

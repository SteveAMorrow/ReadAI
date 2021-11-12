using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BmfCat = Bentley.Plant.MF.Catalog;
using Bentley.ECObjects.Instance;
using Bentley.ECObjects.Schema;
using System.Windows.Forms;
using Bom    = Bentley.Plant.ObjectModel;

namespace Bentley.Plant.App.Pid.AIImport
{
    public class ReadAI
    {
        public ReadAI(string aiFileName)
        {
            LoadJson(aiFileName);            
        }

        public void LoadJson(string aiFileName)
        {
            if (string.IsNullOrEmpty(aiFileName) || !File.Exists(aiFileName))
            {
                Valid = false;
                return;
            }
            DeserializeConfiguration(aiFileName);
        }

        public bool Valid { get; set; } = true;
        public string ErrorText { get; private set; } = string.Empty;
        public AIRoot AIObject { get; set; }
        // public Object OBJ { get; set; }

        public AIRoot DeserializeConfiguration(string aiFileName)
        {
            if (string.IsNullOrEmpty(aiFileName) || !File.Exists(aiFileName))
                return null;

            try
            {
                JObject jsonData = JObject.Parse(File.ReadAllText(aiFileName));
                AIObject = JsonConvert.DeserializeObject<AIRoot>(jsonData.ToString());

                // OBJ = JsonConvert.DeserializeObject(jsonData.ToString()); 
                return AIObject;
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                if (ex.InnerException != null)
                    msg = msg + Environment.NewLine + ex.InnerException.Message;
                ErrorText = ErrorText + aiFileName + Environment.NewLine;
                ErrorText = ErrorText + msg;
                return null;
            }
        }

        public IDictionary<string, JValue> GetProperties (Region ri)
        {
            if (!Valid || ri == null) return null;
            IDictionary<string, JValue> properties = new Dictionary<string,JValue> ();
            
            JObject props = ri.properties as JObject;
            if (props == null) return null;

            foreach (JProperty x in (JToken)props) 
            {
                string name = x.Name;
                JValue value = x.Value as JValue;
                properties.Add(name, value);
            }

            return properties;
            
        }
        public string GetStringValue (string name, Region ri)
        {
            IDictionary<string, JValue> properties = GetProperties(ri);
            if (properties == null) return "";
            if (!properties.TryGetValue(name, out JValue val))
                return "";
            return val.ToString();
        }
        public JToken GetNativeValue (string name, Region ri)
        {
            IDictionary<string, JValue> properties = GetProperties(ri);
            if (properties == null) return null;
            if (!properties.TryGetValue(name, out JValue val))
                return null;
            return val;
        }

        public string GetClassName (Region ri)
        {
            string mainKey = Mapping.Instance.PropertyNames.FirstOrDefault(x => x.Value == Mapping.CLASSNAME).Key;
            string altKey = Mapping.Instance.PropertyNames.FirstOrDefault(x => x.Value == Mapping.ALT_CLASSNAME).Key;
            
            if (string.IsNullOrEmpty(mainKey)) return string.Empty;
            IDictionary<string, JValue> properties = GetProperties(ri);
            if (properties == null) return string.Empty;
            JValue mainCls;
            if (!properties.TryGetValue(mainKey, out mainCls))
            {
                // try alternate key 
                if (!properties.TryGetValue(altKey, out mainCls))
                    return string.Empty;
            }

            JValue altCls;
            if (properties.TryGetValue(altKey, out altCls))
                mainCls = altCls;

            if (mainCls == null) return string.Empty;

            string keyName = mainCls.Value.ToString();
            if (Mapping.Instance.ClassNames.TryGetValue(keyName, out string clsName) && !string.IsNullOrEmpty(clsName))
                return clsName;
            
            return keyName;
        }

        public IList<string> GetAssociatedId (Region ri)        
        {

            IList<string> ids = new List<string>();
            if (ri.associatedTo == null || ri.associatedTo.Count == 0) return null;
            foreach (AssociatedTo a in ri.associatedTo)
            {
                ids.Add(a.id);
            }
            return ids;
        }

        public Region FindRelatedId (string id)
        {
            foreach (Region ri in AIObject.regions)
            {
                if (ri.id.Equals(id))
                {
                    return ri;
                }
            }
            return null;
        }

        public void CreateInstance (Region ri)
        {
            string tag = ri.userLabel;
            string clsName = GetClassName(ri);
            IECClass cls = BmfCat.CatalogServer.GetECClass (clsName);
            if (cls == null)
            {
                MessageBox.Show("class not found");
                return;
            }

            Bom.IComponent oppidComponent = PIDUtilities.CreateComponentFromClassName (clsName, APPIDUtilities.ProjectSchemaName);
            if (null == oppidComponent)
                return;
            oppidComponent.BusinessKey = tag;
            Bom.Workspace.ActiveModel.InsertComponent (oppidComponent, null);
            if (null == oppidComponent)
                return;
        }

    }

}

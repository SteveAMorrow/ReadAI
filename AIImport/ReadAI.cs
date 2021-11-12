using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

                foreach (Region r in AIObject.regions)
                {
                    var val = r.GetType();
                    var fi = val.GetFields();

                    Console.Write(r.associatedTo);
                    Console.Write(r.boundingBox);
                }

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
    }

// Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse); 

}

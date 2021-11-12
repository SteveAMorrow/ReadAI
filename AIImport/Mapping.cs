using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bentley.Plant.App.Pid.AIImport
{
    public class Mapping
    {
        private static readonly object mutex = new object ();  
        private static Mapping m_mapping = null;
        public const string CLASSNAME = "CLASSNAME";
        public const string ALT_CLASSNAME = "ALT_CLASSNAME ";

        public IDictionary<string,string> ClassNames {get;set;}
        public IDictionary<string,string> PropertyNames {get;set;}

        public static Mapping Instance
            {
            get
                {
                lock(mutex)
                    {
                    if(m_mapping == null)
                        m_mapping = new Mapping ();
                    return m_mapping;
                    }
                }
            }
        private Mapping()
        {
            setMappings ();
        }


        private void setMappings()
        {
            //todo read mapping for a schema or json
            // or add custom attributes
            // or a new schema
            ClassNames = new Dictionary<string,string>();            
            ClassNames.Add("gate", "GATE_VALVE");
            ClassNames.Add("control_valve", "CONTROL_VALVE"); //direct match??

            // todo class and their properties
            PropertyNames = new Dictionary<string,string>();
            PropertyNames.Add("type", CLASSNAME);
            PropertyNames.Add("instrument type", ALT_CLASSNAME); //seems this is for control valves

            PropertyNames.Add("openness", "A");
            PropertyNames.Add("operation", "A");
            PropertyNames.Add("nb inlets","A");
            PropertyNames.Add("regulation","A");
            PropertyNames.Add("angle",    "A");
            PropertyNames.Add("fail type","A");
            PropertyNames.Add("location","A");
            PropertyNames.Add("measured quantity","A");
            
        }
    }
}

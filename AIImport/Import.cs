using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CTLog = Bentley.Plant.CommonTools.Logging;

namespace Bentley.Plant.App.Pid.AIImport
{
    public class Import
    {
    private bool autoSelect = false;
        public Import ()
        {

        }

        public bool Run ()
            {
            ReadAI ai = new ReadAI(selectAiFile);
            if (!ai.Valid) return false;
            
            Display dd = new Display(ai);
            dd.Show();
            return ai.Valid;
            }

        private string selectAiFile
        {
            get
            {
                // for debugging
                if (autoSelect)
                    return @"E:\Work\Innovation (OPPID ReadAI)\PIDP400_0_new.json";

                string filter = string.Format("{0} (*.json)|*.json", CTLog.FrameworkLogger.Instance.GetLocalizedString("JsonFile"));
                OpenFileDialog fileDialog = new OpenFileDialog();
                fileDialog.Title = CTLog.FrameworkLogger.Instance.GetLocalizedString("SelectAiCaption");
                fileDialog.Multiselect = false;
                fileDialog.Filter = filter;
                fileDialog.FileName = string.Empty;
                DialogResult result = fileDialog.ShowDialog();
                return (result != DialogResult.OK) ? null : fileDialog.FileName;
            }
        }
        

    }
}

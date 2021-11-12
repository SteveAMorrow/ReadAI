using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Bentley.Plant.App.Pid.AIImport
{
    public partial class Display : Form
    {
        public Display(ReadAI ai)
        {
            InitializeComponent();
            populate (ai);
        }

        public void populate (ReadAI ai)
        {
            if (!ai.Valid) return;
            
            foreach (Region ri in ai.AIObject.regions)
            {
                if (string.IsNullOrEmpty(ri.iModelClass)) continue;
                ListViewItem lv = new ListViewItem(ri.userLabel);
                lv.SubItems.Add(ri.iModelClass);
                listView.Items.Add(lv);
            }
        }

    }
}

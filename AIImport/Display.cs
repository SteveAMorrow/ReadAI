using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MW = Bentley.MstnPlatformNET.WinForms;

namespace Bentley.Plant.App.Pid.AIImport
{
    //public partial class Display : Form
    public partial class Display : MW.Adapter
    {
        private ReadAI m_ai;
        public Display(ReadAI ai)
        {
            this.AttachAsTopLevelForm (null, false, Text);
            InitializeComponent();
            m_ai = ai;
            populate ();
        }

        public void populate ()
        {
            if (!m_ai.Valid) return;
            
            foreach (Region ri in m_ai.AIObject.regions)
            {
                if (string.IsNullOrEmpty(ri.iModelClass)) continue;
                ListViewItem lv = new ListViewItem(ri.id);
                lv.SubItems.Add(ri.userLabel);
                lv.SubItems.Add(ri.iModelClass);
                lv.SubItems.Add(m_ai.GetClassName(ri));
                lv.Tag = ri;
                listView.Items.Add(lv);
            }
        }

        private void listView_SelectedIndexChanged(object sender, EventArgs e)
        {
            listViewProps.Items.Clear();
            tabControl.Controls.Clear();
            if (listView.SelectedItems == null || listView.SelectedItems.Count == 0) return;
            Region ri = listView.SelectedItems[0].Tag as Region;
            if (ri == null) return;

            var props = m_ai.GetProperties(ri);
            foreach (string prop in props.Keys)
            {
                ListViewItem lv = new ListViewItem(prop);
                lv.SubItems.Add(props[prop].Value.ToString());
                listViewProps.Items.Add(lv);
            }

            IList<string> associatedIds = m_ai.GetAssociatedId (ri);
            if (associatedIds == null || associatedIds.Count == 0) return;

            foreach(string id in associatedIds)
            {                
                Region ari = m_ai.FindRelatedId(id);
                if (ari == null) continue;
                TabPage tabPage = new TabPage(id);
                ListView pglistView = new ListView();
                pglistView.Columns.Add(GetColumnHeader("Name"));
                pglistView.Columns.Add(GetColumnHeader("Value"));
                pglistView.View = View.Details;
                pglistView.Items.Add(setLvValue("type",ari.type));
                pglistView.Items.Add(setLvValue("text",ari.text));
                pglistView.Items.Add(setLvValue("orientation",ari.orientation));
                pglistView.Items.Add(setLvValue("symmetry",ari.symmetry));

                pglistView.Items.Add(setLvValue("",""));
                pglistView.Items.Add(setLvValue("boundingBox",""));
                pglistView.Items.Add(setLvValue("height",ari.boundingBox.height.ToString()));
                pglistView.Items.Add(setLvValue("width",ari.boundingBox.width.ToString()));
                pglistView.Items.Add(setLvValue("left",ari.boundingBox.left.ToString()));
                pglistView.Items.Add(setLvValue("top",ari.boundingBox.top.ToString()));

                pglistView.Items.Add(setLvValue("",""));

                if (ari.points != null && ari.points.Count > 0)
                {
                    int i = 0;
                    ListViewItem lv = new ListViewItem("points");
                    lv.SubItems.Add("x,y");
                    pglistView.Items.Add(lv);
                    foreach (Point p in ari.points)
                    {
                        lv = new ListViewItem(string.Format("Point-{0}",i++));
                        lv.SubItems.Add(string.Format("{0},{1}",p.x, p.y));
                        pglistView.Items.Add(lv);
                    }
                }

                tabPage.Controls.Add(pglistView);
                pglistView.Dock = DockStyle.Fill;
                tabControl.Controls.Add(tabPage);                
            }

        }

        private ListViewItem setLvValue(string name, string value)
        {
            ListViewItem lv = null;
            lv = new ListViewItem(name);
            lv.SubItems.Add(value);
            return lv;
        }

        private ColumnHeader GetColumnHeader (string name)
        {
            ColumnHeader colName = new ColumnHeader();
            colName.Text = name;
            colName.Width = 100;
            return colName;
        }

        private void buttonCreate_Click(object sender, EventArgs e)
        {
            if (listView.SelectedItems == null || listView.SelectedItems.Count == 0) return;
            Region ri = listView.SelectedItems[0].Tag as Region;
            if (ri == null) return;
            m_ai.CreateInstance(ri);
        }
    }
}

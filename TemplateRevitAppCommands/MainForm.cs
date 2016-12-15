using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Form = System.Windows.Forms.Form;

namespace Revit_Spec_VK
{
  
    public partial class MainForm : Form
    {
        public static MainForm mf ;
        public static string EventName;
        private UIApplication app;
        //public bool isMyEvent = false;

        private List<ErrorMessage> errorMessagesrrors;
        public MainForm(UIApplication appRevit,List<ErrorMessage> errors)
        {
            InitializeComponent();
            mf = this;
            app = appRevit;
            errorMessagesrrors = errors.GroupBy(x => x.ID).Select(g => g.First()).ToList();
            //isMyEvent = false;
           foreach (var er in errorMessagesrrors)
            {
                dg.Rows.Add();
                dg[0, dg.RowCount - 1].Value = er.Message;
                dg[1, dg.RowCount - 1].Value = er.ID;
                dg.Rows[dg.RowCount - 1].HeaderCell.Value = (dg.RowCount).ToString();
            }
           
        }

        public MainForm()
        {
            InitializeComponent();
            mf = this; 

        }


        private void dg_CellMouseDoubleClick(object sender, System.Windows.Forms.DataGridViewCellMouseEventArgs e)
        {
            //if (!isMyEvent) return;
            if (dg.SelectedRows.Count == 0) return;
            int id = Convert.ToInt32(dg[1, dg.SelectedRows[0].Index].Value);
            if (id == 0) return;
            app.ActiveUIDocument.ShowElements(app.ActiveUIDocument.Document.GetElement(new ElementId(id)));
            List<ElementId> ids = new List<ElementId>();
            ids.Add(new ElementId(id));
            app.ActiveUIDocument.Selection.SetElementIds(ids);
        }

        private void выделитьВМоделиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dg.SelectedRows.Count == 0) return;
            List<ElementId> ids = new List<ElementId>();
            for (int i = 0; i < dg.SelectedRows.Count; i++)
            {
                int id = Convert.ToInt32(dg[1, dg.SelectedRows[i].Index].Value);
                if (id == 0) continue;
                ids.Add(new ElementId(id));
            }
            if (ids.Count == 0)
                return;
            app.ActiveUIDocument.ShowElements(ids);
            app.ActiveUIDocument.Selection.SetElementIds(ids);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace Sets
{
    public partial class ThisAddIn
    {
        //private OptionsPane _mOptionsPane;
        //private Microsoft.Office.Tools.CustomTaskPane _mTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //_mOptionsPane = new OptionsPane();
            //_mTaskPane = this.CustomTaskPanes.Add(_mOptionsPane, "Sets Options");
            //_mTaskPane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

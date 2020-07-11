using Microsoft.Office.Tools;
using System;
using System.IO;

namespace PowerWorshipVSTO
{
    public partial class ThisAddIn
    {
        public static String appDataPath;// = $@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\PowerWorship";

        // Called from ribbon constructor
        public static void PreInitialize()
        {
            // Get the assembly information
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();

            // CodeBase is the location of the DLL
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
            appDataPath = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString()) + "\\Data";
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            
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

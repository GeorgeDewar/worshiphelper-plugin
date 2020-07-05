using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
//using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
//using Office = Microsoft.Office.Core;

namespace PowerWorshipVSTO
{
    public partial class TestRibbonItem
    {
        private void TestRibbonItem_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void btnInsertSong_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.PowerPoint.Application app = Globals.ThisAddIn.Application;

            FileDialog dialog = app.FileDialog[MsoFileDialogType.msoFileDialogOpen];
            dialog.Title = "Select Song or Presentation";
            if (dialog.Show() == -1) // If user selected a file
            {
                foreach(string item in dialog.SelectedItems)
                {
                    var sourcePresentation = app.Presentations.Open(item, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                    sourcePresentation.Slides.Range().Copy();
                    app.CommandBars.ExecuteMso("PasteSourceFormatting");
                }
            }
        }
    }
}

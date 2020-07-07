﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using static Microsoft.Office.Core.MsoTriState;
//using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
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
            Application app = Globals.ThisAddIn.Application;

            FileDialog dialog = app.FileDialog[MsoFileDialogType.msoFileDialogOpen];
            dialog.Title = "Select Song or Presentation";
            if (dialog.Show() == -1) // If user selected a file
            {
                foreach(string item in dialog.SelectedItems)
                {
                    var sourcePresentation = app.Presentations.Open(item, msoTrue, msoFalse, msoFalse);
                    sourcePresentation.Slides.Range().Copy();
                    sourcePresentation.Close();
                    app.CommandBars.ExecuteMso("PasteSourceFormatting");
                }
            }
        }

        private void btnInsertScripture_Click(object sender, RibbonControlEventArgs e)
        {
            new InsertScriptureForm().Show();
        }

        private void btnInsertOneClick_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;

            var fileName = (sender as RibbonButton).Tag as string;
            var sourcePresentation = app.Presentations.Open(fileName, msoTrue, msoFalse, msoFalse);
            sourcePresentation.Slides.Range().Copy();
            sourcePresentation.Close();
            app.CommandBars.ExecuteMso("PasteSourceFormatting");
        }
    }
}

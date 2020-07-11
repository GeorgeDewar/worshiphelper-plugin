using Microsoft.Office.Core;
using static Microsoft.Office.Core.MsoTriState;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Win32;
using System.IO;

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

            var registryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\PowerWorship");
            var lastSongLocation = registryKey.GetValue("LastSongLocation") as string;
            
            FileDialog dialog = app.FileDialog[MsoFileDialogType.msoFileDialogOpen];
            dialog.Title = "Select Song or Presentation";
            dialog.AllowMultiSelect = false;
            if (lastSongLocation != null) dialog.InitialFileName = lastSongLocation;
            if (dialog.Show() == -1) // If user selected a file
            {
                var selectedDirectory = Path.GetDirectoryName(dialog.SelectedItems.Item(1));
                registryKey.SetValue("LastSongLocation", selectedDirectory);
                foreach (string item in dialog.SelectedItems)
                {
                    var sourcePresentation = app.Presentations.Open(item, msoTrue, msoFalse, msoFalse);
                    sourcePresentation.Slides.Range().Copy();
                    sourcePresentation.Close();
                    app.CommandBars.ExecuteMso("PasteSourceFormatting");
                    ScriptureManager.goToEnd();
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

            var fileName = (sender as RibbonControl).Tag as string;
            var sourcePresentation = app.Presentations.Open(fileName, msoTrue, msoFalse, msoFalse);
            sourcePresentation.Slides.Range().Copy();
            sourcePresentation.Close();
            app.CommandBars.ExecuteMso("PasteSourceFormatting");
            ScriptureManager.goToEnd();
        }

        private void btnRemoveOneClick_Click(object sender, RibbonControlEventArgs e)
        {
            var favRegistryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\PowerWorship\Favourites");
            var fileName = (sender as RibbonControl).Tag as string;
            try
            {
                favRegistryKey.DeleteValue(fileName);
            } catch (System.ArgumentException)
            {
                System.Windows.Forms.MessageBox.Show("This item appears to have already been deleted - try restarting PowerPoint.");
            }
            ((sender as RibbonControl).Parent as RibbonControl).Visible = false;
        }

        private void btnAddFavourite_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;

            var registryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\PowerWorship");
            var favRegistryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\PowerWorship\Favourites");
            var lastSongLocation = registryKey.GetValue("LastSongLocation") as string;

            FileDialog dialog = app.FileDialog[MsoFileDialogType.msoFileDialogOpen];
            dialog.Title = "Select Song or Presentation";
            if (lastSongLocation != null) dialog.InitialFileName = lastSongLocation;
            if (dialog.Show() == -1) // If user selected a file
            {
                foreach (string item in dialog.SelectedItems)
                {
                    favRegistryKey.SetValue(item, item);
                }
                System.Windows.Forms.MessageBox.Show(
                    "Your new favourite has been added, but it won't appear until you restart PowerPoint.",
                    "Favourite added");
            }

            
        }
    }
}

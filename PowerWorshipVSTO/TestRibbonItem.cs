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
            var favRegistryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\PowerWorship\Favourites");
            var favouriteCount = favRegistryKey.GetValueNames().Length;
            for (int i=0; i < favouriteCount; i++)
            {
                var file = favRegistryKey.GetValueNames()[i];
                var slideButton = favouriteButtons[i];
                
                var pathParts = file.Split(new char[] { '\\' });
                slideButton.Label = pathParts[pathParts.Length - 1].Replace(".pptx", "").Replace(".ppt", "");
                slideButton.Tag = file;
                slideButton.Visible = true;
            }

            // Hide the unused buttons
            for (int i = favouriteCount; i < favouriteButtons.Count; i++)
            {
                var slideButton = favouriteButtons[i];
                slideButton.Visible = false;
            }
            btnAddFavourite.Enabled = favouriteCount < favouriteButtons.Count;

            #if !DEBUG
            grpDebug.Visible = false;
            #endif
        }

        private void btnInsertSong_Click(object sender, RibbonControlEventArgs e)
        {
            new SongManager().InsertSong();
        }

        private void btnInsertScripture_Click(object sender, RibbonControlEventArgs e)
        {
            new InsertScriptureForm().Show();
        }

        private void btnInsertOneClick_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;

            string fileName;
            if(sender is RibbonSplitButton)
            {
                // This IS the parent SplitButton
                fileName = (sender as RibbonControl).Tag as string;
            } else
            {
                // Get it from the tag of the parent SplitButton
                fileName = (sender as RibbonControl).Parent.Tag as string;
            }

            new SongManager().InsertSongFromFile(fileName);
        }

        private void btnRemoveOneClick_Click(object sender, RibbonControlEventArgs e)
        {
            var favRegistryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\PowerWorship\Favourites");
            var fileName = (sender as RibbonControl).Parent.Tag as string;
            try
            {
                favRegistryKey.DeleteValue(fileName);
            } catch (System.ArgumentException)
            {
                System.Windows.Forms.MessageBox.Show("This item appears to have already been deleted - try restarting PowerPoint.");
            }

            // Force a refresh
            TestRibbonItem_Load(null, null);
        }

        private void btnAddFavourite_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;

            var registryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\PowerWorship");
            var favRegistryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\PowerWorship\Favourites");

            if (favRegistryKey.GetValueNames().Length >= 5) {
                System.Windows.Forms.MessageBox.Show("No more favourites can be added");
            }
                        
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
            }

            // Force a refresh
            TestRibbonItem_Load(null, null);
        }

        private void btnSelfTest_Click(object sender, RibbonControlEventArgs e)
        {
            new SelfTestManager().SelfTest();
        }
    }
}

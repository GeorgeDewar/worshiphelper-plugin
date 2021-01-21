using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Win32;
using System.IO;
using static Microsoft.Office.Core.MsoTriState;

namespace PowerWorshipVSTO
{
    class SongManager
    {
        Application app = Globals.ThisAddIn.Application;

        public void InsertSong()
        {
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
                    InsertSongFromFile(item);
                }
            }
        }

        public void InsertSongFromFile(string filePath)
        {
            var sourcePresentation = app.Presentations.Open(filePath, msoTrue, msoFalse, msoFalse);
            sourcePresentation.Slides.Range().Copy();
            sourcePresentation.Close();
            app.CommandBars.ExecuteMso("PasteSourceFormatting");
            ScriptureManager.goToEnd();
        }
    }
}

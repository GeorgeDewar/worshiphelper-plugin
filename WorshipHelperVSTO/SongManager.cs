using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using static Microsoft.Office.Core.MsoTriState;

namespace WorshipHelperVSTO
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
            Debug.WriteLine($"Inserting presentation from file at {filePath}");

            try
            {
                var selectionManager = new SelectionManager();
                var insertIndex = selectionManager.GetNextSlideIndex();
                if (insertIndex > 1)
                {
                // Paste will already paste in the slide *after* the selected one
                    selectionManager.GoToSlide(insertIndex - 1);
                } else
                {
                // There must be no slides, no need to navigate
                }

                var sourcePresentation = app.Presentations.Open(filePath, msoTrue, msoFalse, msoFalse);
                sourcePresentation.Slides.Range().Copy();
                sourcePresentation.Close();

                app.CommandBars.ExecuteMso("PasteSourceFormatting");
                //ScriptureManager.goToEnd();
            } catch (Exception e)
            {
                Debug.WriteLine(e.Message);
                System.Windows.Forms.MessageBox.Show($"An error occurred while inserting the song or presentation.\n\n{e.Message}");
            }
        }
    }
}

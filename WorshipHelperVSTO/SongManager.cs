using log4net;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Win32;
using System;
using System.IO;
using static Microsoft.Office.Core.MsoTriState;

namespace WorshipHelperVSTO
{
    class SongManager
    {
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        Application app = Globals.ThisAddIn.Application;

        public void InsertSong()
        {
            log.Info("Inserting song...");
            var registryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\WorshipHelper");
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
            log.Info($"Inserting presentation from file at {filePath}");

            try
            {
                var selectionManager = new SelectionManager();
                var insertIndex = selectionManager.GetNextSlideIndex();
                if (insertIndex > 1)
                {
                    // Paste will already paste in the slide *after* the selected one
                    log.Debug($"Navigating to slide {insertIndex - 1} to insert the presentation");
                    selectionManager.GoToSlide(insertIndex - 1);
                } else
                {
                    // There must be no slides, no need to navigate
                    log.Debug("No slides present, no need to navigate");
                }

                var sourcePresentation = app.Presentations.Open(filePath, msoTrue, msoFalse, msoFalse);
                sourcePresentation.Slides.Range().Copy();
                sourcePresentation.Close();

                app.CommandBars.ExecuteMso("PasteSourceFormatting");
                //ScriptureManager.goToEnd();
                log.Debug("Presentation inserted successfully");
            } catch (Exception e)
            {
                log.Error("An error occurred while inserting the presentation", e);
                System.Windows.Forms.MessageBox.Show($"An error occurred while inserting the song or presentation.\n\n{e.Message}");
            }
        }
    }
}
